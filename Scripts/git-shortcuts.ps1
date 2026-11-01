function Push-GitCommit {
    <#
    .SYNOPSIS
    Stages everything, commits and pushes the current branch.

    .DESCRIPTION
    Runs git add -A, commits with the given title (and optional description),
    then pushes to origin. With -AI, Claude Code writes the title and
    description from the staged diff, following .claude/skills/commit.md in the
    repo when that file exists.

    .PARAMETER Title
    Commit title. Required unless -AI is used.

    .PARAMETER Desc
    Optional commit description.

    .PARAMETER AI
    Generate the title and description with the Claude Code CLI.

    .EXAMPLE
    gcp 'fix: handle empty cart'

    .EXAMPLE
    gcp -AI

    .NOTES
    Alias: gcp
    #>
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

        if (-not $stagedFiles -and -not $stagedDiff.Trim()) {
            Write-Host "Nothing to commit — no staged changes found after 'git add -A'." -ForegroundColor Red
            return
        }

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

        $aiOutput = $null
        try {
            $OutputEncoding = [System.Text.Encoding]::UTF8
            $aiOutput = ($prompt | claude --print) -join "`n"
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Claude CLI exited with code $LASTEXITCODE — aborting." -ForegroundColor Red
                if ($aiOutput) { Write-Host "Output: $aiOutput" -ForegroundColor Gray }
                return
            }
            if (-not $aiOutput) {
                Write-Host "Claude returned empty output — aborting." -ForegroundColor Red
                return
            }
            if ($aiOutput -match "(?m)^TITLE:\s*(.+)$") { $Title = $Matches[1].Trim() }
            if ($aiOutput -match "(?m)^DESC:\s*(.+)$")  { $Desc  = $Matches[1].Trim() }
        } catch {
            Write-Host "Claude CLI call failed: $($_.Exception.Message) — aborting." -ForegroundColor Red
            return
        }

        if (-not $Title) {
            Write-Host "Could not parse 'TITLE:' from Claude output — aborting." -ForegroundColor Red
            Write-Host "Raw output:" -ForegroundColor Yellow
            Write-Host $aiOutput -ForegroundColor Gray
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
    <#
    .SYNOPSIS
    Pulls then pushes the current branch against origin.

    .EXAMPLE
    gsync

    .NOTES
    Alias: gsync
    #>
    [CmdletBinding()]
    param ()

    $branch = git rev-parse --abbrev-ref HEAD
    git pull origin $branch
    git push origin $branch
}

function Undo-GitCommit {
    <#
    .SYNOPSIS
    Undoes the last commit but keeps its changes staged.

    .DESCRIPTION
    Runs git reset --soft HEAD~1, so nothing is lost; the files stay staged.

    .EXAMPLE
    gundo

    .NOTES
    Alias: gundo
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param ()

    if ($PSCmdlet.ShouldProcess('HEAD', 'Undo last commit (soft reset)')) {
        git reset --soft HEAD~1
    }
}

function Edit-GitCommit {
    <#
    .SYNOPSIS
    Rewrites the message of the last commit.

    .PARAMETER NewMessage
    The replacement commit message.

    .EXAMPLE
    gamend 'feat: add batch tracking'

    .NOTES
    Alias: gamend
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$NewMessage
    )
    if ($PSCmdlet.ShouldProcess('HEAD', "Amend commit message to '$NewMessage'")) {
        git commit --amend -m $NewMessage
    }
}

function Show-GitDiff {
    <#
    .SYNOPSIS
    Shows changed, staged, untracked and deleted files as an icon tree.

    .DESCRIPTION
    Builds a tree from git's changed, staged and untracked paths, and a
    separate red tree for deleted files. Uses Terminal-Icons for the icons.

    .PARAMETER BasePath
    Repo root to resolve paths against. Defaults to the current repo root.

    .EXAMPLE
    gdiff

    .NOTES
    Alias: gdiff
    #>
    [CmdletBinding()]
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

    function Write-TreeNode {
        param (
            [hashtable]$Node,
            [string]$ParentPath,
            [string]$Indent = "",
            [bool]$IsLast = $true,
            [bool]$DeletedMode = $false
        )

        # @() keeps a single child as a one-item array (a bare string would index to its first letter)
        $keys = @($Node.Keys | Sort-Object)
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
                    Write-TreeNode -Node $Node[$key] -ParentPath $fullPath -Indent $newIndent -IsLast $isLastChild -DeletedMode:$false
                }
            }

            # Existing paths recurse above (only real folders); deleted paths recurse here
            if ($DeletedMode -and $null -ne $Node[$key] -and $Node[$key].Count -gt 0) {
                $newIndent = $Indent + ($isLastChild ? "   " : "│  ")
                Write-TreeNode -Node $Node[$key] -ParentPath $fullPath -Indent $newIndent -IsLast $isLastChild -DeletedMode:$DeletedMode
            }
        }
    }

    if ($existingPaths.Count -gt 0) {
        Write-Host "📂 Changed/Staged/Untracked Files" -ForegroundColor Cyan
        $existingTree = Build-TreeFromPaths -Paths $existingPaths
        Write-TreeNode -Node $existingTree -ParentPath $BasePath -DeletedMode:$false
        Write-Host "`n"
    }

    if ($deletedPaths.Count -gt 0) {
        Write-Host "🗑️ Deleted Files" -ForegroundColor Magenta
        $deletedTree = Build-TreeFromPaths -Paths $deletedPaths
        Write-TreeNode -Node $deletedTree -ParentPath $BasePath -DeletedMode:$true
        Write-Host "`n"
    }

    if ($existingPaths.Count -eq 0 -and $deletedPaths.Count -eq 0) {
        Write-Host "✅ No changes detected." -ForegroundColor Green
    }
}

function Show-GitGraph {
    <#
    .SYNOPSIS
    Shows the commit history of all branches as a decorated graph.

    .EXAMPLE
    glg

    .NOTES
    Alias: glg
    #>
    [CmdletBinding()]
    param ()

    git log --oneline --graph --decorate --all
}

function New-GitPR {
    <#
    .SYNOPSIS
    Creates or updates a GitHub pull request with an AI-written title and body.

    .DESCRIPTION
    Collects the commits and diff against the base branch, detects linked
    issues from commits and the branch name, then has the Claude Code CLI write
    the title, labels and body (following .claude/skills/pr.md when present).
    Creates the PR through gh, or updates the existing one for the branch.
    Falls back to the browser compare page when gh is missing.

    .PARAMETER Base
    Branch to compare against. Defaults to master.

    .PARAMETER DryRun
    Print the title, labels and body without creating or changing anything.

    .EXAMPLE
    gpr -DryRun

    .EXAMPLE
    gpr -Base develop

    .NOTES
    Alias: gpr. Always uses AI for the body.
    #>
    [CmdletBinding()]
    param (
        [string]$Base = "master",
        [switch]$DryRun
    )

    $currentBranch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $currentBranch -or $currentBranch -eq $Base) {
        Write-Host "Already on base branch '$Base'. Checkout a feature branch first." -ForegroundColor Red
        return
    }

    # Gather commits since diverging from base — normalize to clean unique array
    $rawCommits = @(git log "$Base..HEAD" --pretty=format:"%s" 2>$null)
    $commitList = $rawCommits | Where-Object { $_ -ne "" } | ForEach-Object { $_.Trim() } | Select-Object -Unique
    if ($commitList.Count -eq 0) {
        Write-Host "No commits found between '$Base' and '$currentBranch'." -ForegroundColor Yellow
        return
    }

    # Changed files and diff stat
    $changedFiles = @(git diff "$Base...HEAD" --name-only 2>$null) | Where-Object { $_ -ne "" }
    $diffStat = git diff "$Base...HEAD" --stat 2>$null | Select-Object -Last 1

    # --- Collect referenced issue numbers from commits + branch name ---
    # GitHub auto-links and auto-closes issues when the PR body contains
    # "Fixes #N" / "Closes #N" / "Resolves #N". Scan commit messages and the
    # branch name for bare "#N" or conventional "WP-N"/"ISSUE-N" references.
    $issueNumbers = [System.Collections.Generic.HashSet[int]]::new()
    $issueRegex   = '(?<!\w)#(\d+)\b'

    $commitBodies = git log "$Base..HEAD" --pretty=format:"%B" 2>$null | Out-String
    foreach ($match in [regex]::Matches("$commitBodies $currentBranch", $issueRegex)) {
        [void]$issueNumbers.Add([int]$match.Groups[1].Value)
    }

    # Branch names like "fix/42-something" or "feature/issue-42-something"
    if ($currentBranch -match '(?<!\d)(?:issue[-_]?|#)(\d+)') {
        [void]$issueNumbers.Add([int]$Matches[1])
    } elseif ($currentBranch -match '^[^/]+/(\d+)[-_]') {
        [void]$issueNumbers.Add([int]$Matches[1])
    }

    $linkedIssues = @($issueNumbers | Sort-Object)
    if ($linkedIssues.Count -gt 0) {
        Write-Host "Detected linked issues: $(($linkedIssues | ForEach-Object { "#$_" }) -join ', ')" -ForegroundColor DarkGray
    }

    # --- Fetch open issues from GitHub so Claude can suggest additional links ---
    # The local scan above only catches issues already referenced in commits/branch.
    # Pulling the open issue list lets Claude correlate PR content with unreferenced
    # issues (e.g. a PR that fixes an issue but whose commit message never said so).
    $openIssuesJson = $null
    if (Get-Command gh -ErrorAction SilentlyContinue) {
        $ghIssues = gh issue list --state open --limit 100 --json number,title,labels 2>$null
        if ($LASTEXITCODE -eq 0 -and $ghIssues) {
            $openIssuesJson = $ghIssues | Out-String
            $openCount = ($ghIssues | ConvertFrom-Json).Count
            Write-Host "Fetched $openCount open issues from GitHub" -ForegroundColor DarkGray
        }
    }

    # --- Derive PR title from commit messages ---
    # Strip conventional commit prefixes to get clean descriptions
    $cleanCommits = $commitList | ForEach-Object {
        $_ -replace "^(feat|fix|chore|docs|refactor|test|ci|perf|style|build)(\(.+?\))?!?:\s*", ""
    } | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) }

    if ($commitList.Count -eq 1) {
        # Single commit — use the clean title-cased message
        $prTitle = $cleanCommits[0]
    } elseif ($cleanCommits.Count -le 3) {
        # 2-3 commits — join with " + "
        $prTitle = $cleanCommits -join " + "
        if ($prTitle.Length -gt 70) { $prTitle = $prTitle.Substring(0, 67) + "..." }
    } else {
        # Many commits — take the most common prefix type and summarize from branch name
        $titleFromBranch = $currentBranch -replace "^.+/", ""
        $titleFromBranch = $titleFromBranch -replace "[-_]", " "
        $prTitle = (Get-Culture).TextInfo.ToTitleCase($titleFromBranch.ToLower())
    }

    # --- Load skill file for PR conventions ---
    $repoRoot  = git rev-parse --show-toplevel 2>$null
    $skillFile = if ($repoRoot) { Join-Path $repoRoot ".claude/skills/pr.md" } else { $null }
    $skillText = if ($skillFile -and (Test-Path $skillFile)) {
        Write-Host "Using PR conventions from $skillFile" -ForegroundColor DarkGray
        Get-Content $skillFile -Raw
    } else { $null }

    # --- Build PR body via AI ---
    if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
        Write-Host "Claude Code CLI ('claude') not found in PATH. Cannot generate PR body." -ForegroundColor Red
        return
    }

    Write-Host "Generating AI summary..." -ForegroundColor Cyan

    # File list with status (A/M/D/R) — always included in full
    $nameStatus = git diff "$Base...HEAD" --name-status 2>$null | Out-String

    # Commit log with full descriptions (not just subjects)
    $commitLog = git log "$Base..HEAD" --pretty=format:"%h %s%n%b" 2>$null | Out-String

    # Full diff — no truncation, let Claude read everything
    $fullDiff = git diff "$Base...HEAD" 2>$null | Out-String
    Write-Host "Diff size: $([math]::Round($fullDiff.Length / 1024))KB ($($changedFiles.Count) files)" -ForegroundColor DarkGray

    $conventionsBlock = if ($skillText) {
        "The following skill file defines the conventions for this repo — follow them exactly:`n`n$skillText"
    } else {
        @"
Title rules:
- Title Case phrases joined by " + " (e.g. "Refactor Account Service + Add Integration Tests")
- Imperative mood, no trailing period
- Derived from commit messages, not branch name
- Never include usernames or author names in the title

Body sections:
## Summary
(2-5 sentences describing what this PR does and why — elevator pitch)

## Changes
(bulleted list grouped by feature/concern — describe behavioral changes, not filenames)

## Package Bumps
(table of changed dependencies with From/To versions — omit if none)

## Test Plan
(short checklist, 1-3 items)

## Notes
(caveats, breaking changes, follow-ups — omit if nothing relevant)
"@
    }

    # Write the diff to a temp file so Claude can read it via tools
    # instead of stuffing it into the prompt (which hits the input limit)
    $diffFile = Join-Path ([System.IO.Path]::GetTempPath()) "pr-diff-$([System.IO.Path]::GetRandomFileName()).patch"
    [System.IO.File]::WriteAllText($diffFile, $fullDiff, [System.Text.Encoding]::UTF8)
    Write-Host "Diff written to $diffFile ($([math]::Round((Get-Item $diffFile).Length / 1024))KB)" -ForegroundColor DarkGray

    $issuesBlock = ""
    if ($linkedIssues.Count -gt 0) {
        $issueList = ($linkedIssues | ForEach-Object { "#$_" }) -join ", "
        $issuesBlock += @"

Linked issues (already detected from commits / branch): $issueList
Include a "Closes $issueList" line at the bottom of the body so GitHub auto-links
and auto-closes these issues when the PR merges. Use "Closes" for each issue
(e.g. "Closes #1, closes #2") so every issue is linked individually.
"@
    }
    if ($openIssuesJson) {
        $issuesBlock += @"


Open GitHub issues (JSON array of {number,title,labels}):
$openIssuesJson

If any of the ABOVE open issues are clearly resolved by the diff you read, add
them to the closing line as well (e.g. "Closes #12, closes #17"). Only include
an issue if the diff unambiguously addresses its title — when in doubt, skip it.
"@
    }

    $prompt = @"
You are writing a GitHub pull request title and description.
Summarize what the code DOES DIFFERENTLY after this PR — do NOT list commit
messages, filenames, or SHAs as bullets. Group changes by feature/concern.

Branch: $currentBranch -> $Base
Stat: $diffStat

Commit log (for context/intent — do NOT copy as summary bullets):
$commitLog

Files changed (A=added, M=modified, D=deleted, R=renamed):
$nameStatus

The full diff is saved at: $diffFile
Read this file using the Read tool to understand every change. Read it in
chunks if needed (use offset/limit parameters). Do NOT skip reading the diff.
$issuesBlock
$conventionsBlock

After reading the diff, respond in this exact format (no extra text before TITLE):
TITLE: <title>
LABELS: <comma-separated labels from the skill file>
---
<body markdown>
"@

    $body = $null
    try {
        Write-Host "Sending to Claude (with Read tool access)..." -ForegroundColor DarkGray
        $aiOutput = ($prompt | claude --print --allowedTools "Read" --dangerously-skip-permissions) -join "`n"
        if ($LASTEXITCODE -ne 0 -or -not $aiOutput) {
            Write-Host "Claude exited with code $LASTEXITCODE." -ForegroundColor Red
            if ($aiOutput) { Write-Host "Output: $aiOutput" -ForegroundColor Gray }
            return
        }
        if ($aiOutput -match "(?m)^TITLE:\s*(.+)$")  { $prTitle  = $Matches[1].Trim() }
        if ($aiOutput -match "(?m)^LABELS:\s*(.+)$") { $aiLabels = $Matches[1].Trim() }
        $body = $aiOutput -replace "(?m)^TITLE:.*$\r?\n?",  ""
        $body = $body    -replace "(?m)^LABELS:.*$\r?\n?", ""
        $body = $body    -replace "(?m)^---\s*$\r?\n?",     ""
        $body = $body.Trim()
    } catch {
        Write-Host "AI call failed: $($_.Exception.Message)" -ForegroundColor Red
        return
    } finally {
        Remove-Item $diffFile -ErrorAction SilentlyContinue
    }

    if (-not $body) {
        Write-Host "Could not parse PR body from Claude output. Raw output:" -ForegroundColor Red
        Write-Host $aiOutput -ForegroundColor Gray
        return
    }

    # Safety net — if Claude didn't include closing keywords, append them so the
    # linked issues still get auto-closed on merge.
    if ($linkedIssues.Count -gt 0) {
        $missing = @()
        foreach ($n in $linkedIssues) {
            if ($body -notmatch "(?im)(close[ds]?|fix(e[ds])?|resolve[ds]?)\s+#$n\b") {
                $missing += "#$n"
            }
        }
        if ($missing.Count -gt 0) {
            $closingLine = "Closes " + (($missing | ForEach-Object { $_ }) -join ", closes ")
            $body = $body.TrimEnd() + "`n`n$closingLine`n"
            Write-Host "Appended closing keywords for: $($missing -join ', ')" -ForegroundColor DarkGray
        }
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
            $prInfo      = $prJson | ConvertFrom-Json
            $existingPr  = $prInfo.number
            $existingUrl = $prInfo.url
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
        $editBodyFile = Join-Path ([System.IO.Path]::GetTempPath()) "pr-edit-body-$([System.IO.Path]::GetRandomFileName()).md"
        [System.IO.File]::WriteAllText($editBodyFile, $body, [System.Text.Encoding]::UTF8)
        $editArgs = @("pr", "edit", $existingPr, "--title", $prTitle, "--body-file", $editBodyFile, "--add-assignee", "@me")
        if ($aiLabels) { foreach ($l in $aiLabels -split ',') { $editArgs += @("--add-label", $l.Trim()) } }
        gh @editArgs
        Remove-Item $editBodyFile -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PR #$existingPr updated: $existingUrl" -ForegroundColor Green
            Start-Process $existingUrl
        } else {
            Write-Warning "gh pr edit failed. Open the PR manually: $existingUrl"
        }
    } elseif ($hasGh) {
        # Create via gh so labels and assignee are applied immediately, then open it
        Write-Host "Creating PR via gh..." -ForegroundColor Cyan
        $bodyFile = Join-Path ([System.IO.Path]::GetTempPath()) "pr-body-$([System.IO.Path]::GetRandomFileName()).md"
        [System.IO.File]::WriteAllText($bodyFile, $body, [System.Text.Encoding]::UTF8)
        try {
            $ghArgs = @("pr", "create", "--base", $Base, "--title", $prTitle, "--body-file", $bodyFile, "--assignee", "@me")
            $newPrUrl = gh @ghArgs 2>&1 | Select-Object -Last 1
            # Add labels separately so missing labels don't fail the PR creation
            if ($LASTEXITCODE -eq 0 -and $aiLabels) {
                foreach ($l in $aiLabels -split ',') {
                    gh pr edit --add-label $l.Trim() 2>$null
                }
            }
        } finally {
            Remove-Item $bodyFile -ErrorAction SilentlyContinue
        }
        if ($newPrUrl -match "^https://") {
            Write-Host "PR created: $newPrUrl" -ForegroundColor Green
            Start-Process $newPrUrl
        } else {
            Write-Warning "gh pr create failed (exit code: $LASTEXITCODE). Output: $newPrUrl"
            if ($repoPath) {
                Set-Clipboard $body
                Write-Host "PR body copied to clipboard. Paste it in the description field." -ForegroundColor Yellow
                $truncatedTitle = if ($prTitle.Length -gt 200) { $prTitle.Substring(0, 197) + "..." } else { $prTitle }
                $encodedTitle = [Uri]::EscapeDataString($truncatedTitle)
                Start-Process "https://github.com/$repoPath/compare/$Base...$($currentBranch)?expand=1&title=$encodedTitle"
            }
        }
    } elseif ($repoPath) {
        Set-Clipboard $body
        Write-Host "PR body copied to clipboard. Paste it in the description field." -ForegroundColor Yellow
        $truncatedTitle = if ($prTitle.Length -gt 200) { $prTitle.Substring(0, 197) + "..." } else { $prTitle }
        $encodedTitle = [Uri]::EscapeDataString($truncatedTitle)
        Write-Host "Opening GitHub PR page in browser..." -ForegroundColor Cyan
        Start-Process "https://github.com/$repoPath/compare/$Base...$($currentBranch)?expand=1&title=$encodedTitle"
    } else {
        Write-Warning "Could not parse GitHub URL from remote '$remoteUrl'. Push complete — open GitHub manually."
    }
}

Set-Alias -Name gcp   -Value Push-GitCommit
Set-Alias -Name gsync -Value Sync-GitBranch
Set-Alias -Name gundo -Value Undo-GitCommit
Set-Alias -Name gamend -Value Edit-GitCommit
Set-Alias -Name gdiff -Value Show-GitDiff
Set-Alias -Name glg   -Value Show-GitGraph
Set-Alias -Name gpr   -Value New-GitPR
