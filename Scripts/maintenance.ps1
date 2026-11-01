function Update-DevEnvironment {
    <#
    .SYNOPSIS
    Updates PowerShell modules, Chocolatey packages and winget packages.

    .DESCRIPTION
    Runs each package source in turn. Chocolatey needs an elevated shell, so it
    is skipped with a warning when the shell is not elevated. winget also
    updates Oh My Posh, which is an MSIX install and cannot update itself.

    .PARAMETER SkipModules
    Do not update PowerShell modules.

    .PARAMETER SkipChoco
    Do not update Chocolatey packages.

    .PARAMETER SkipWinget
    Do not update winget packages.

    .EXAMPLE
    upall

    Updates everything.

    .EXAMPLE
    upall -SkipChoco

    Updates modules and winget packages only.

    .NOTES
    Alias: upall
    #>
    [CmdletBinding()]
    param (
        [switch]$SkipModules,
        [switch]$SkipChoco,
        [switch]$SkipWinget
    )

    Write-Host '🚀 Starting global update...' -ForegroundColor Cyan

    if (-not $SkipModules) {
        Write-Host '📦 Checking for PowerShell module updates...' -ForegroundColor Yellow
        try {
            # Update the package provider first to prevent handshake errors
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue
            Update-Module -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-Host '✅ PowerShell modules updated.' -ForegroundColor Green
        }
        catch {
            Write-Warning 'Some modules could not be updated (they may be in use or require Admin).'
        }
    }

    if (-not $SkipChoco -and (Get-Command choco -ErrorAction SilentlyContinue)) {
        $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Host '🍫 Checking Chocolatey packages...' -ForegroundColor Yellow
            choco upgrade all -y
        }
        else {
            Write-Warning 'Skipping Chocolatey: run this from an elevated shell to upgrade choco packages.'
        }
    }

    if (-not $SkipWinget -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host '📦 Checking winget packages...' -ForegroundColor Yellow
        winget upgrade --all --silent --accept-package-agreements --accept-source-agreements
    }

    Write-Host '✨ All systems updated!' -ForegroundColor Cyan
}

function Update-Profile {
    <#
    .SYNOPSIS
    Reloads the profile and every script in the Scripts folder.

    .DESCRIPTION
    Dot-sourcing a file from inside a function defines everything in that
    function's scope, so it vanishes on return. This re-defines each function
    and alias from the Scripts folder in the global scope instead, then
    re-runs the profile for everything else (key bindings, PATH, modules).

    .PARAMETER Clear
    Clear the screen afterwards.

    .EXAMPLE
    reload

    .NOTES
    Alias: reload
    #>
    [CmdletBinding()]
    param (
        [switch]$Clear
    )

    Write-Host '🔄 Reloading PowerShell profile...' -ForegroundColor Cyan

    $scriptsDir = Join-Path (Split-Path -Parent $PROFILE) 'Scripts'
    if (-not (Test-Path -LiteralPath $scriptsDir)) { $scriptsDir = 'S:\Personal\Scripts' }

    foreach ($file in Get-ChildItem -Path $scriptsDir -Filter *.ps1) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)

        foreach ($function in $ast.EndBlock.Statements |
                Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
            Set-Item -Path "function:global:$($function.Name)" -Value $function.Body.GetScriptBlock()
        }

        foreach ($command in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $false) |
                Where-Object { $_.GetCommandName() -eq 'Set-Alias' }) {
            Invoke-Expression "$($command.Extent.Text) -Scope Global -Force"
        }
    }

    . $PROFILE

    if ($Clear) { Clear-Host }
}

Set-Alias -Name upall  -Value Update-DevEnvironment
Set-Alias -Name reload -Value Update-Profile
