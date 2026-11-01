function Stop-Browser {
    <#
    .SYNOPSIS
    Kills a browser and its WebDriver process tree.

    .DESCRIPTION
    Force-kills every process of the chosen browser plus its driver
    (geckodriver for Firefox, chromedriver for Chrome). Useful when a Selenium
    or Playwright run leaves orphaned browser processes behind.

    .PARAMETER Browser
    Which browser to kill: firefox (ff) or chrome (c).

    .EXAMPLE
    Stop-Browser ff

    Kills Firefox and geckodriver.

    .EXAMPLE
    kb chrome -WhatIf

    Shows what would be killed without killing anything.

    .NOTES
    Alias: kb
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet('firefox', 'ff', 'chrome', 'c')]
        [string]$Browser
    )

    $images = if ($Browser -in 'firefox', 'ff') {
        'firefox.exe', 'geckodriver.exe'
    }
    else {
        'chrome.exe', 'chromedriver.exe'
    }

    foreach ($image in $images) {
        if ($PSCmdlet.ShouldProcess($image, 'Kill process tree')) {
            taskkill /F /IM $image /T
        }
    }
}

function Get-ModifiedFilesSince {
    <#
    .SYNOPSIS
    Lists the files modified since a point in time.

    .DESCRIPTION
    Returns the names of files under Path whose LastWriteTime is at or after
    Since. Defaults to the last 24 hours in the current folder.

    .PARAMETER Path
    Folder to search. Defaults to the current folder.

    .PARAMETER Since
    Cut-off date and time. Defaults to 24 hours ago.

    .PARAMETER Recurse
    Search subfolders too.

    .EXAMPLE
    Get-ModifiedFilesSince -Recurse

    Files changed in the last day, including subfolders.

    .EXAMPLE
    recent S:\Personal (Get-Date).AddHours(-2)

    Files in S:\Personal changed in the last two hours.

    .NOTES
    Alias: recent
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Position = 0)]
        [string]$Path = '.',

        [Parameter(Position = 1)]
        [datetime]$Since = (Get-Date).AddDays(-1),

        [switch]$Recurse
    )

    Get-ChildItem -Path $Path -File -Recurse:$Recurse |
        Where-Object { $_.LastWriteTime -ge $Since } |
        Select-Object -ExpandProperty Name
}

function Resolve-RelativePath {
    <#
    .SYNOPSIS
    Joins a base path and a relative path, whichever slash style it uses.

    .DESCRIPTION
    Splits Relative on either slash and joins each segment onto Base, so
    mixed separators from git output resolve cleanly. The path does not need
    to exist. A missing Base falls back to the current folder.

    .PARAMETER Base
    Starting folder. Defaults to the current folder when empty.

    .PARAMETER Relative
    Relative path, with or without separators. Empty returns Base unchanged.

    .EXAMPLE
    Resolve-RelativePath -Base 'S:\Personal' -Relative 'Scripts/maintenance.ps1'

    Returns S:\Personal\Scripts\maintenance.ps1.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [string]$Base,
        [string]$Relative
    )

    if (-not $Base) { $Base = (Get-Location).Path }
    if ([string]::IsNullOrWhiteSpace($Relative)) { return $Base }

    foreach ($segment in ($Relative -split '[\\/]' | Where-Object { $_ })) {
        $Base = Join-Path -Path $Base -ChildPath $segment
    }

    return $Base
}

Set-Alias -Name kb     -Value Stop-Browser
Set-Alias -Name recent -Value Get-ModifiedFilesSince
