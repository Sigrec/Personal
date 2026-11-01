# Personal PowerShell setup

My PowerShell 7 profile, helper scripts and Oh My Posh themes.

| Path | What it is |
|------|------------|
| `Microsoft.PowerShell_profile.ps1` | The profile: prompt, modules, PSReadLine, PATH |
| `Scripts/` | Helper scripts; every `*.ps1` here is dot-sourced by the profile |
| `Themes/` | Oh My Posh themes (`wild_west_powerline.omp.json` is the active one; `wild_west_vibes` is the old plain prompt) |
| `fastfetch/` | fastfetch config and the ASCII logos (`options/2_triskele.txt` is active; the others are spares) |
| `settings.json` | Terminal settings |

## Scripts

Run `cmds` to list every custom command with its alias and description, or
`cmds git*` to filter. For full help on one command, run `help <command> -Full`.
Every function has comment-based help and an approved verb; each file ends with
its short aliases.

| File | Commands (alias) |
|------|------------------|
| `utility-functions.ps1` | `Stop-Browser` (`kb`), `Get-ModifiedFilesSince` (`recent`), `Resolve-RelativePath` |
| `maintenance.ps1` | `Update-DevEnvironment` (`upall`), `Update-Profile` (`reload`) |
| `project-shortcuts.ps1` | `Push-WpCode` (`wpc`), `Push-WpWiki` (`wpw`) |
| `explorer-shortcuts.ps1` | `Show-FolderStructure` (`ftree`), `Update-FileTimestamp` (`touch`) |
| `git-shortcuts.ps1` | `Push-GitCommit` (`gcp`), `Sync-GitBranch` (`gsync`), `Undo-GitCommit` (`gundo`), `Edit-GitCommit` (`gamend`), `Show-GitDiff` (`gdiff`), `Show-GitGraph` (`glg`), `New-GitPR` (`gpr`) |
| `command-index.ps1` | `Get-MyCommand` (`cmds`) |

## Setup on a new machine

Install the tools:

```powershell
winget install JanDeDobbeleer.OhMyPosh
winget install ajeetdsouza.zoxide
winget install junegunn.fzf
```

Install the PowerShell modules:

```powershell
Install-Module Terminal-Icons, posh-git, PSFzf, CompletionPredictor -Scope CurrentUser
```

Point the profile at this repo (PowerShell 7 loads `$PROFILE`, so either copy
the file there or dot-source it):

```powershell
# 1. Create the profile if it does not exist yet
if (-not (Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force }

# 2. Dot-source the one in this repo from it
Add-Content $PROFILE '. "S:\Personal\Microsoft.PowerShell_profile.ps1"'
```

If Windows blocks the scripts, right-click each file, open Properties, and tick
**Unblock** on the General tab, then open a new shell.

## Secrets

Never put API keys or webhook secrets in the profile; this repo is public.
Put them in `profile.local.ps1` next to the profile. It is git-ignored and
loaded automatically if it exists.

## Updating

`Update-DevEnvironment` (`upall`) updates PowerShell modules, winget packages and,
from an elevated shell, Chocolatey packages. Oh My Posh is an MSIX install, so
`oh-my-posh upgrade` does not work; winget updates it.

## Startup time

Terminal-Icons, posh-git, PSFzf and the Chocolatey completions take about 1.3s
to import, so the profile loads them on the first idle moment after the prompt
is drawn instead of before it.
