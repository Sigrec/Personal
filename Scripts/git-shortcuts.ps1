function Git-AddCommitPush {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string] $Title,
        [string] $Desc,
        [Parameter(Mandatory = $true)] [string] $Branch
    )

    git add -A
    git commit -m $Title -m "$Desc"
    git push origin $Branch
}

function Git-SyncBranch {
    $branch = git rev-parse --abbrev-ref HEAD
    git pull origin $branch
    git push origin $branch
}

function Git-UndoLastCommit {
    git reset --soft HEAD~1
}

function Git-AmendLastCommit {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]$NewMessage
    )
    git commit --amend -m "$NewMessage"
}

function Git-Diff {
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

function Git-StatusTree {
    git log --oneline --graph --decorate --all
}