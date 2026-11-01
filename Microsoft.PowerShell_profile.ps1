# Set UTF-8 encoding for shell and external process I/O
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

# Start OMP styling
oh-my-posh init pwsh --config "S:\Personal\Themes\wild_west_powerline.omp.json" | Invoke-Expression

# Modules needed before the first prompt (PSReadLine is already loaded by the host)
Import-Module CompletionPredictor
Invoke-Expression (& { (zoxide init powershell | Out-String) })

# Import local scripts (with error isolation so profile doesn't die)
Get-ChildItem -Path "S:\Personal\Scripts" -Filter *.ps1 | ForEach-Object {
    try {
        . $_.FullName
    }
    catch {
        Write-Warning ("Failed loading script: {0}`n{1}" -f $_.FullName, $_.Exception.Message)
    }
}

# Machine-specific values (API keys, webhook secrets) live in a git-ignored file
$LocalProfile = Join-Path $PSScriptRoot "profile.local.ps1"
if (Test-Path $LocalProfile) { . $LocalProfile }

# Enable Read line
if ($host.Name -eq 'ConsoleHost')
{
    $PSReadLineOptions = @{
        HistoryNoDuplicates = $true
        HistorySearchCursorMovesToEnd = $true
        PredictionSource = 'HistoryAndPlugin'
        PredictionViewStyle = 'InlineView'
        # Removed "Default" and used specific keys to avoid parsing errors
        Colors = @{
            "Command" = "#F7E0B4"
        }
    }
    Set-PSReadLineOption @PSReadLineOptions

    # Up/Down search history by the text already typed (like zsh)
    Set-PSReadLineKeyHandler -Key UpArrow   -Function HistorySearchBackward
    Set-PSReadLineKeyHandler -Key DownArrow -Function HistorySearchForward

    # Slow modules (about 1.3s together) load after the first prompt is drawn,
    # so the shell is usable right away: Terminal-Icons, posh-git, PSFzf and
    # Chocolatey tab-completion. Ctrl+R / Ctrl+T switch to fzf once PSFzf loads.
    $null = Register-EngineEvent -SourceIdentifier PowerShell.OnIdle -MaxTriggerCount 1 -Action {
        Import-Module Terminal-Icons, posh-git, PSFzf -Global
        Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'

        $chocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
        if ($env:ChocolateyInstall -and (Test-Path $chocolateyProfile)) {
            Import-Module $chocolateyProfile -Global
        }
    }
}

# System summary on a fresh interactive terminal (skipped for -Command / -File runs)
if ($host.Name -eq 'ConsoleHost' -and (Get-Command fastfetch -ErrorAction SilentlyContinue) -and
    -not ([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-(c|command|f|file|encodedcommand)$' })) {
    fastfetch --config "S:\Personal\fastfetch\config.jsonc"
}

# dotnetup SDK location
if ($env:PATH -notlike "*C:\Users\Sean\.dotnetup*") {
    $env:PATH = "C:\Users\Sean\.dotnetup;$env:PATH"
}
