function Print-FolderStructure {
    [CmdletBinding()]
    param (
        [string]$Path=$(Get-Location).Path,
        [string[]]$ExcludeFolders = @(),
        [string]$Indent = "",
        [bool]$IsLast = $true
    )

    $item = Get-Item -LiteralPath $Path
    $iconEntry = $item | Format-TerminalIcons
    $prefix = if ($Indent) { $Indent + ($IsLast ? "└─ " : "├─ ") } else { "" }
    Write-Host "$prefix$($iconEntry)" 

    if ($item.PSIsContainer) {
        $children = Get-ChildItem -LiteralPath $item.FullName -ErrorAction SilentlyContinue |
            Where-Object { -not ($ExcludeFolders -contains $_.Name -and $_.PSIsContainer) } |
            Sort-Object @{Expression = { if ($_.PSIsContainer) { 0 } else { 1 } }}, Name

        $count = $children.Count
        for ($i = 0; $i -lt $count; $i++) {
            Print-FolderStructure -Path $children[$i].FullName -ExcludeFolders $ExcludeFolders `
                   -Indent ($Indent + ($IsLast ? "   " : "│  ")) -IsLast:($i -eq $count - 1)
        }
    }
}

function Touch {
    [CmdletBinding()]
    param (
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        New-Item -ItemType File -Path $Path | Out-Null
    } else {
        (Get-Item $Path).LastWriteTime = Get-Date
    }
}