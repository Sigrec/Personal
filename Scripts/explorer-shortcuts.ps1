function Show-FolderStructure {
    <#
    .SYNOPSIS
    Prints a folder as a tree with Terminal-Icons file icons.

    .DESCRIPTION
    Recursively lists Path with box-drawing connectors, folders first. Uses
    Terminal-Icons for the icons and colors, so that module must be loaded.

    .PARAMETER Path
    Folder or file to print. Defaults to the current folder.

    .PARAMETER ExcludeFolders
    Folder names to skip, such as bin, obj or node_modules.

    .PARAMETER Indent
    Internal: the prefix for the current depth.

    .PARAMETER IsLast
    Internal: whether this entry is the last child of its parent.

    .EXAMPLE
    Show-FolderStructure -ExcludeFolders bin, obj, node_modules

    .EXAMPLE
    ftree src

    .NOTES
    Alias: ftree
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]$Path = (Get-Location).Path,

        [string[]]$ExcludeFolders = @(),

        [string]$Indent = '',

        [bool]$IsLast = $true
    )

    $item = Get-Item -LiteralPath $Path
    $iconEntry = $item | Format-TerminalIcons
    $prefix = if ($Indent) { $Indent + ($IsLast ? '└─ ' : '├─ ') } else { '' }
    Write-Host "$prefix$iconEntry"

    if (-not $item.PSIsContainer) { return }

    $children = Get-ChildItem -LiteralPath $item.FullName -ErrorAction SilentlyContinue |
        Where-Object { -not ($_.PSIsContainer -and $ExcludeFolders -contains $_.Name) } |
        Sort-Object @{ Expression = { -not $_.PSIsContainer } }, Name

    $count = @($children).Count
    $childIndent = $Indent + ($IsLast ? '   ' : '│  ')

    for ($i = 0; $i -lt $count; $i++) {
        Show-FolderStructure -Path $children[$i].FullName -ExcludeFolders $ExcludeFolders `
            -Indent $childIndent -IsLast:($i -eq $count - 1)
    }
}

function Update-FileTimestamp {
    <#
    .SYNOPSIS
    Creates an empty file, or bumps its last-write time if it exists (touch).

    .DESCRIPTION
    Works like the Unix touch command. Accepts several paths and pipeline input.

    .PARAMETER Path
    One or more files to create or update.

    .EXAMPLE
    touch notes.txt, todo.txt

    .NOTES
    Alias: touch
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]$Path
    )

    process {
        foreach ($file in $Path) {
            if (-not $PSCmdlet.ShouldProcess($file, 'Create or update timestamp')) { continue }

            if (Test-Path -LiteralPath $file) {
                (Get-Item -LiteralPath $file).LastWriteTime = Get-Date
            }
            else {
                New-Item -ItemType File -Path $file | Out-Null
            }
        }
    }
}

Set-Alias -Name ftree -Value Show-FolderStructure
Set-Alias -Name touch -Value Update-FileTimestamp
