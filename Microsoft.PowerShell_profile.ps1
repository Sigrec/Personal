# Start OMP styling
oh-my-posh init pwsh --config "S:\Personal\Themes\wild_west_vibes.omp.json" | Invoke-Expression

# Import modules
Import-Module -Name Terminal-Icons
Import-Module posh-git
Import-Module PSReadLine
Invoke-Expression (& { (zoxide init powershell | Out-String) })

# Import local scripts (with error isolation so profile doesn't die)
Get-ChildItem -Path "S:\Personal\Scripts" -Filter *.ps1 -Recurse | ForEach-Object {
    try {
        . $_.FullName
    }
    catch {
        Write-Warning ("Failed loading script: {0}`n{1}" -f $_.FullName, $_.Exception.Message)
    }
}

# Enable Read line
if ($host.Name -eq 'ConsoleHost')
{
    $PSReadLineOptions = @{
        HistoryNoDuplicates = $true
        HistorySearchCursorMovesToEnd = $true
        # Removed "Default" and used specific keys to avoid parsing errors
        Colors = @{
            "Command" = "#F7E0B4"
        }
    }
    Set-PSReadLineOption @PSReadLineOptions
}

# Command to kill browser instances
function kb([string]$Browser)
{
    $Browser = $Browser.ToLower()
    Switch($Browser)
    {
        {$_ -in "firefox", "ff"} { taskkill /F /IM firefox.exe /T ; taskkill /F /IM geckodriver.exe /T }
        {$_ -in "chrome", "c"} { taskkill /F /IM chrome.exe /T ; taskkill /F /IM chromedriver.exe /T }
        default { Write-Error "Invalid browser" }
    }
}

function Get-ModifiedFilesSince {
    [CmdletBinding()]
    param (
        [string]$Path = ".",
        [datetime]$Since = (Get-Date).AddDays(-1),
        [switch]$Recurse
    )

    Get-ChildItem -Path $Path -File -Recurse:$Recurse | Where-Object {
        $_.LastWriteTime -ge $Since
    } | Select-Object -ExpandProperty Name
}

# Import the Chocolatey Profile that contains the necessary code to enable
# tab-completions to function for `choco`.
# Be aware that if you are missing these lines from your profile, tab completion
# for `choco` will not function.
# See https://ch0.co/tab-completion for details.
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}

function Resolve-RelativePath {
    param ([string]$Base, [string]$Relative)

    # Added guard: if Base is null, use current location or return null
    if (-not $Base) { $Base = (Get-Location).Path }
    if ([string]::IsNullOrWhiteSpace($Relative)) { return $Base }

    if ($Relative -like "*[\\/]*") {
        $segments = $Relative -split '[\\/]' | Where-Object { $_ }
        foreach ($segment in $segments) {
            $Base = Join-Path -Path $Base -ChildPath $segment
        }
        return $Base
    } else {
        return Join-Path -Path $Base -ChildPath $Relative
    }
}