# Start OMP styling
oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH/remedy_dark.omp.json" | Invoke-Expression

# Import modules
Import-Module -Name Terminal-Icons
Import-Module posh-git
Import-Module PSReadLine
Import-Module z

# Import local scripts
Get-ChildItem -Path "$(Split-Path -Path $PROFILE)\Scripts\Internal" -Filter *.ps1 -Recurse | ForEach-Object { 
    . $_.FullName
}

# Enable Read line
if ($host.Name -eq 'ConsoleHost')
{
    $PSReadLineOptions = @{
        # EditMode = "Emacs"
        HistoryNoDuplicates = $true
        HistorySearchCursorMovesToEnd = $true
        Colors = @{
            "Default" = "#F7E0B4"
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
