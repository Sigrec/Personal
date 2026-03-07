function Push-GitCommit {
    [CmdletBinding()]
    param (
        [string] $Title,
        [string] $Desc,
        [switch] $AI
    )

    $Branch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $Branch) {
        Write-Host "Not inside a git repository." -ForegroundColor Red
        return
    }

    if (-not $AI -and -not $Title) {
        Write-Host "Either -Title or -AI is required." -ForegroundColor Red
        return
    }

    if ($AI -and -not (Get-Command claude -ErrorAction SilentlyContinue)) {
        Write-Warning "Claude Code CLI ('claude') not found in PATH."
        if (-not $Title) { return }
        $AI = $false
    }

    git add -A

    if ($AI) {
        Write-Host "Generating AI commit message..." -ForegroundColor Cyan
        $stagedFiles = @(git diff --cached --name-only 2>$null) -join "`n"
        $stagedDiff  = git diff --cached 2>$null | Out-String
        if ($stagedDiff.Length -gt 8000) {
            $stagedDiff = $stagedDiff.Substring(0, 8000) + "`n...(truncated)"
        }

        $repoRoot      = git rev-parse --show-toplevel 2>$null
        $commitSkill   = if ($repoRoot) { Join-Path $repoRoot ".claude/skills/commit.md" } else { $null }
        $commitConvs   = if ($commitSkill -and (Test-Path $commitSkill)) {
            Write-Host "Using commit conventions from $commitSkill" -ForegroundColor DarkGray
            "The following skill file defines the conventions for this repo — follow them exactly:`n`n$(Get-Content $commitSkill -Raw)"
        } else {
            "Follow conventional commits format (feat, fix, chore, docs, refactor, test).`nTitle max 72 chars, imperative mood, no trailing period."
        }

        $prompt = @"
Generate a git commit message for these staged changes.

Staged files:
$stagedFiles

Diff (may be truncated):
$stagedDiff

$commitConvs

Respond in this exact format (no extra text):
TITLE: <commit title>
DESC: <optional 1-2 sentence description, or leave blank>
"@

        try {
            $OutputEncoding = [System.Text.Encoding]::UTF8
            $aiOutput = ($prompt | claude --print) -join "`n"
            if ($LASTEXITCODE -eq 0 -and $aiOutput) {
                if ($aiOutput -match "(?m)^TITLE:\s*(.+)$") { $Title = $Matches[1].Trim() }
                if ($aiOutput -match "(?m)^DESC:\s*(.+)$")  { $Desc  = $Matches[1].Trim() }
            } else {
                Write-Warning "Claude returned no output."
            }
        } catch {
            Write-Warning "AI call failed: $($_.Exception.Message)"
        }

        if (-not $Title) {
            Write-Host "Could not generate commit title — aborting." -ForegroundColor Red
            return
        }

        Write-Host "Title: $Title" -ForegroundColor Green
        if ($Desc) { Write-Host "Desc:  $Desc" -ForegroundColor Gray }
    }

    if ($Desc) {
        git commit -m $Title -m $Desc
    } else {
        git commit -m $Title
    }
    git push origin $Branch
}

function Sync-GitBranch {
    $branch = git rev-parse --abbrev-ref HEAD
    git pull origin $branch
    git push origin $branch
}

function Undo-GitCommit {
    git reset --soft HEAD~1
}

function Edit-GitCommit {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]$NewMessage
    )
    git commit --amend -m "$NewMessage"
}

function Show-GitDiff {
    param (
        [string]$BasePath = $(git rev-parse --show-toplevel 2>$null)
    )

    if (-not $BasePath) {
        Write-Host "❌ Not inside a Git repository." -ForegroundColor Red
        return
    }

    # Gather all file paths
    $changed    = git diff --name-only HEAD 2>$null
    $staged     = git diff --cached --name-only 2>$null
    $untracked  = git ls-files --others --exclude-standard 2>$null
    $deleted    = git diff --name-only --diff-filter=D HEAD 2>$null

    $existingPaths = $changed + $staged + $untracked | Sort-Object -Unique
    $deletedPaths  = $deleted | Sort-Object -Unique

    function Build-TreeFromPaths {
        param ([string[]]$Paths)

        $tree = @{}

        foreach ($path in $Paths) {
            $parts = $path -split '[\\/]+' 
            $currentLevel = $tree

            for ($i = 0; $i -lt $parts.Count; $i++) {
                $part = $parts[$i]

                $isLeaf = ($i -eq $parts.Count - 1)

                if (-not $currentLevel.ContainsKey($part)) {
                    if ($isLeaf) {
                        # Leaf node = file, set value to $null
                        $currentLevel[$part] = $null
                    } else {
                        # Folder = nested hashtable
                        $currentLevel[$part] = @{}
                    }
                }

                if (-not $isLeaf) {
                    $currentLevel = $currentLevel[$part]
                }
            }
        }

        return $tree
    }

    function Print-Tree {
        param (
            [hashtable]$Node,
            [string]$ParentPath,
            [string]$Indent = "",
            [bool]$IsLast = $true,
            [bool]$DeletedMode = $false
        )

        $keys = $Node.Keys | Sort-Object
        $count = $keys.Count

        for ($i = 0; $i -lt $count; $i++) {
            $key = $keys[$i]
            $isLastChild = ($i -eq $count - 1)
            $pathParts = $key -split '[\\/]'
            $fullPath = Resolve-RelativePath -Base $ParentPath -Relative $key

            $prefix = if ($Indent) {
                $Indent + ($isLastChild ? "└─ " : "├─ ")
            } else {
                ""
            }

            if ($DeletedMode) {
                $simulatedPath = Resolve-RelativePath -Base $ParentPath -Relative $key

                # Guess if it's a directory or file based on node count or extension
                $isLikelyFolder = $null -ne $Node[$key] -and $Node[$key].Count -gt 0
                if ($isLikelyFolder) {
                    $fakeItem = [System.IO.DirectoryInfo]::new($simulatedPath)
                } else {
                    $fakeItem = [System.IO.FileInfo]::new($simulatedPath)
                }

                $icon = $fakeItem | Format-TerminalIcons
                Write-Host "$prefix$icon" -ForegroundColor DarkRed
            }
            else {
                if (-not (Test-Path -LiteralPath $fullPath)) {
                    Write-Host "⚠️  Skipping missing: $fullPath" -ForegroundColor DarkGray
                    continue
                }

                try {
                    $item = Get-Item -LiteralPath $fullPath -ErrorAction Stop
                    $icon = $item | Format-TerminalIcons
                } catch {
                    Write-Host "❌ Error accessing: $fullPath" -ForegroundColor Red
                    continue
                }

                Write-Host "$prefix$icon"

                if ($item.PSIsContainer -and $Node[$key] -ne $null -and $Node[$key].Count -gt 0) {
                    $newIndent = $Indent + ($isLastChild ? "   " : "│  ")
                    $test = $Node[$key]
                    Print-Tree -Node $Node[$key] -ParentPath $fullPath -Indent $newIndent -IsLast $isLastChild -DeletedMode:$false
                }
            }

            # Recurse in deleted mode
            if ($null -ne $Node[$key] -and $Node[$key].Count -gt 0) {
                $newIndent = $Indent + ($isLastChild ? "   " : "│  ")
                Print-Tree -Node $Node[$key] -ParentPath $fullPath -Indent $newIndent -IsLast $isLastChild -DeletedMode:$DeletedMode
            }
        }
    }

    if ($existingPaths.Count -gt 0) {
        Write-Host "📂 Changed/Staged/Untracked Files" -ForegroundColor Cyan
        $existingTree = Build-TreeFromPaths -Paths $existingPaths
        Print-Tree -Node $existingTree -ParentPath $BasePath -DeletedMode:$false
        Write-Host "`n"
    }

    if ($deletedPaths.Count -gt 0) {
        Write-Host "🗑️ Deleted Files" -ForegroundColor Magenta
        $deletedTree = Build-TreeFromPaths -Paths $deletedPaths
        Print-Tree -Node $deletedTree -ParentPath $BasePath -DeletedMode:$true
        Write-Host "`n"
    }

    if ($existingPaths.Count -eq 0 -and $deletedPaths.Count -eq 0) {
        Write-Host "✅ No changes detected." -ForegroundColor Green
    }
}

function Show-GitStatusTree {
    git log --oneline --graph --decorate --all
}

function Git-PR {
    [CmdletBinding()]
    param (
        [string]$Base = "master",
        [switch]$AI,
        [switch]$DryRun
    )

    $currentBranch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $currentBranch -or $currentBranch -eq $Base) {
        Write-Host "Already on base branch '$Base'. Checkout a feature branch first." -ForegroundColor Red
        return
    }

    # Gather commits since diverging from base — normalize to clean unique array
    $rawCommits = @(git log "$Base...HEAD" --pretty=format:"%s" 2>$null)
    $commitList = $rawCommits | Where-Object { $_ -ne "" } | ForEach-Object { $_.Trim() } | Select-Object -Unique
    if ($commitList.Count -eq 0) {
        Write-Host "No commits found between '$Base' and '$currentBranch'." -ForegroundColor Yellow
        return
    }

    # Changed files and diff stat
    $changedFiles = @(git diff "$Base...HEAD" --name-only 2>$null) | Where-Object { $_ -ne "" }
    $diffStat = git diff "$Base...HEAD" --stat 2>$null | Select-Object -Last 1

    # --- Derive fallback PR title from branch name ---
    $titleFromBranch = $currentBranch -replace "^[^/]+/", ""
    $titleFromBranch = $titleFromBranch -replace "^(feature|fix|chore|docs|refactor|test|hotfix)[/_-]", ""
    $titleFromBranch = $titleFromBranch -replace "[_/-]", " "
    $titleFromBranch = (Get-Culture).TextInfo.ToTitleCase($titleFromBranch.ToLower())

    $prTitle = if ($commitList.Count -eq 1) { $commitList[0] } else { $titleFromBranch }

    # --- Load skill file for PR conventions ---
    $repoRoot  = git rev-parse --show-toplevel 2>$null
    $skillFile = if ($repoRoot) { Join-Path $repoRoot ".claude/skills/pr.md" } else { $null }
    $skillText = if ($skillFile -and (Test-Path $skillFile)) {
        Write-Host "Using PR conventions from $skillFile" -ForegroundColor DarkGray
        Get-Content $skillFile -Raw
    } else { $null }

    # --- Build PR body ---
    if ($AI -and -not (Get-Command claude -ErrorAction SilentlyContinue)) {
        Write-Warning "Claude Code CLI ('claude') not found in PATH — falling back to basic body."
        $AI = $false
    }

    if ($AI) {
        Write-Host "Generating AI summary..." -ForegroundColor Cyan
        $fullDiff = git diff "$Base...HEAD" 2>$null | Out-String
        if ($fullDiff.Length -gt 12000) { $fullDiff = $fullDiff.Substring(0, 12000) + "`n...(truncated)" }

        $conventionsBlock = if ($skillText) {
            "The following skill file defines the conventions for this repo — follow them exactly:`n`n$skillText"
        } else {
            @"
Title rules:
- Title Case phrases joined by " + " (e.g. "Refactor Account Service + Add Integration Tests")
- Imperative mood, no trailing period
- Derived from commit messages, not branch name

Body sections:
## Summary
(2-4 bullet points describing what changed and why)

## Changes
(bullet list of notable file/area changes)

## Test Plan
(bulleted checklist of how to verify this works — omit if purely docs/config)

## Notes
(caveats, breaking changes, follow-ups — omit if nothing relevant)
"@
        }

        $prompt = @"
You are writing a GitHub pull request title and description.
Based on the information below, write a concise, professional PR title and body.

Branch: $currentBranch -> $Base
Commits:
$($commitList -join "`n")

Changed files:
$($changedFiles -join "`n")

Diff (may be truncated):
$fullDiff

$conventionsBlock

Respond in this exact format (no extra text before TITLE):
TITLE: <title>
LABELS: <comma-separated labels from the skill file>
---
<body markdown>
"@

        try {
            $OutputEncoding = [System.Text.Encoding]::UTF8
            $aiOutput = ($prompt | claude --print) -join "`n"
            if ($LASTEXITCODE -eq 0 -and $aiOutput) {
                if ($aiOutput -match "(?m)^TITLE:\s*(.+)$")  { $prTitle   = $Matches[1].Trim() }
                if ($aiOutput -match "(?m)^LABELS:\s*(.+)$") { $aiLabels  = $Matches[1].Trim() }
                $body = $aiOutput -replace "(?m)^TITLE:.*$\r?\n?",  ""
                $body = $body    -replace "(?m)^LABELS:.*$\r?\n?", ""
                $body = $body    -replace "(?m)^---\s*$\r?\n?",     ""
                $body = $body.Trim()
                if (-not $body) {
                    Write-Warning "Could not parse AI body — falling back to basic body."
                    $AI = $false
                }
            } else {
                Write-Warning "Claude returned no output — falling back to basic body."
                $AI = $false
            }
        } catch {
            Write-Warning "AI call failed: $($_.Exception.Message) — falling back to basic body."
            $AI = $false
        }
    }

    if (-not $AI) {
        $commitLines = ($commitList | ForEach-Object { "- $_" }) -join "`n"
        $fileLines   = ($changedFiles | Select-Object -First 20 | ForEach-Object { "- $_" }) -join "`n"
        if ($changedFiles.Count -gt 20) { $fileLines += "`n- _(and more...)_" }

        $body = @"
## Summary

$commitLines

## Changes

$fileLines

$( if ($diffStat) { "**Diff:** $diffStat" } )
"@
    }

    # --- Derive GitHub repo path from remote ---
    $remoteUrl = git remote get-url origin 2>$null
    $repoPath  = $null
    if ($remoteUrl -match "github\.com[:/](.+?)(?:\.git)?$") {
        $repoPath = $Matches[1]
    }

    # --- Check for existing PR on this branch ---
    $hasGh       = [bool](Get-Command gh -ErrorAction SilentlyContinue)
    $existingPr  = $null
    $existingUrl = $null
    if ($hasGh) {
        $prJson = gh pr view --json number,url 2>$null
        if ($LASTEXITCODE -eq 0 -and $prJson) {
            $existingPr  = ($prJson | ConvertFrom-Json).number
            $existingUrl = ($prJson | ConvertFrom-Json).url
        }
    }

    if ($DryRun) {
        Write-Host "=== DRY RUN ===" -ForegroundColor Cyan
        Write-Host "Title    : $prTitle" -ForegroundColor Yellow
        Write-Host "Base     : $Base <- $currentBranch" -ForegroundColor Yellow
        Write-Host "Assignee : @me" -ForegroundColor Yellow
        if ($aiLabels) { Write-Host "Labels   : $aiLabels" -ForegroundColor Yellow }
        Write-Host "Body:`n$body" -ForegroundColor Gray
        if ($existingPr) {
            Write-Host "Action : UPDATE existing PR #$existingPr ($existingUrl)" -ForegroundColor DarkCyan
        } elseif ($repoPath) {
            Write-Host "Action : CREATE new PR" -ForegroundColor DarkCyan
            Write-Host "URL    : https://github.com/$repoPath/compare/$Base...$currentBranch" -ForegroundColor DarkCyan
        }
        return
    }

    # Push branch if not already on remote
    $remoteExists = git ls-remote --heads origin $currentBranch 2>$null
    if (-not $remoteExists) {
        Write-Host "Pushing '$currentBranch' to origin..." -ForegroundColor Cyan
        git push -u origin $currentBranch
    }

    Write-Host "Title  : $prTitle" -ForegroundColor Green
    if ($aiLabels) { Write-Host "Labels : $aiLabels" -ForegroundColor Cyan }

    if ($existingPr) {
        # Update existing PR via gh then open it
        Write-Host "Updating PR #$existingPr..." -ForegroundColor Cyan
        $editArgs = @("pr", "edit", $existingPr, "--title", $prTitle, "--body", $body, "--add-assignee", "@me")
        if ($aiLabels) { $editArgs += @("--label", $aiLabels) }
        gh @editArgs
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PR #$existingPr updated: $existingUrl" -ForegroundColor Green
            Start-Process $existingUrl
        } else {
            Write-Warning "gh pr edit failed. Open the PR manually: $existingUrl"
        }
    } elseif ($hasGh) {
        # Create via gh so labels and assignee are applied immediately, then open it
        Write-Host "Creating PR via gh..." -ForegroundColor Cyan
        $ghArgs = @("pr", "create", "--base", $Base, "--title", $prTitle, "--body", $body, "--assignee", "@me")
        if ($aiLabels) { $ghArgs += @("--label", $aiLabels) }
        $newPrUrl = gh @ghArgs 2>&1 | Select-Object -Last 1
        if ($LASTEXITCODE -eq 0 -and $newPrUrl -match "^https://") {
            Write-Host "PR created: $newPrUrl" -ForegroundColor Green
            Start-Process $newPrUrl
        } else {
            Write-Warning "gh pr create failed — falling back to browser."
            if ($repoPath) {
                $encodedTitle = [Uri]::EscapeDataString($prTitle)
                $encodedBody  = [Uri]::EscapeDataString($body)
                Start-Process "https://github.com/$repoPath/compare/$Base...$($currentBranch)?expand=1&title=$encodedTitle&body=$encodedBody"
            }
        }
    } elseif ($repoPath) {
        $encodedTitle = [Uri]::EscapeDataString($prTitle)
        $encodedBody  = [Uri]::EscapeDataString($body)
        $prUrl = "https://github.com/$repoPath/compare/$Base...$($currentBranch)?expand=1&title=$encodedTitle&body=$encodedBody"
        Write-Host "Opening GitHub PR page in browser..." -ForegroundColor Cyan
        Start-Process $prUrl
    } else {
        Write-Warning "Could not parse GitHub URL from remote '$remoteUrl'. Push complete — open GitHub manually."
    }
}