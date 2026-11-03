# Personal PowerShell setup

My PowerShell 7 profile, helper scripts, Oh My Posh themes, WezTerm config and
`mmv`, a viewer for large Mermaid diagrams.

| Path | What it is |
|------|------------|
| `Microsoft.PowerShell_profile.ps1` | The profile: prompt, modules, PSReadLine, PATH |
| `Scripts/` | Helper scripts; every `*.ps1` here is dot-sourced by the profile |
| `Themes/` | Oh My Posh themes (`wild_west_powerline.omp.json` is the active one; `wild_west_vibes` is the old plain prompt) |
| `fastfetch/` | fastfetch config and the ASCII logos (`options/2_triskele.txt` is active; the others are spares) |
| `wezterm/wezterm.lua` | WezTerm config, my main terminal (see [WezTerm](#wezterm)); `~/.wezterm.lua` loads it |
| `btop/wild_west_vibes.theme` | btop theme; symlinked into `%LOCALAPPDATA%\Microsoft\WinGet\Links\themes` (btop reads themes next to the `btop.exe` link) |
| `mmv/` | GPU viewer for large Mermaid diagrams (see [mmv](#mmv)), a Rust project of its own |
| `settings.json` | Windows Terminal settings, kept as a fallback (the live copy is under `%LOCALAPPDATA%\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState`) |

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

## WezTerm

WezTerm is my main terminal. It feels much faster than Windows Terminal and
runs the same profile, Oh My Posh theme and fastfetch. Windows Terminal stays
installed as a fallback and for elevated (admin) shells.

`wezterm/wezterm.lua` holds the config. WezTerm only reads `~/.wezterm.lua`, so
that file is a small stub that loads the one in this repo. Edit the repo copy
and WezTerm reloads it on save.

What the config sets:

- **Shell:** `pwsh.exe -NoLogo`, found through PATH. PowerShell is the Store
  install, so there is no `C:\Program Files\PowerShell\7\pwsh.exe`.
- **Font:** CaskaydiaCove Nerd Font, size 10, `DemiBold`. WezTerm draws text
  thinner than Windows Terminal; switch to `Regular` if it looks too heavy.
- **Colors:** Wild West Vibes, the scheme the powerline prompt is built on.
- **Rendering:** WebGPU on the discrete GPU, capped at 240 fps (the main
  monitor is 240 Hz), steady cursor and a solid window. The Acrylic blur was dropped
  because it added input lag; `window_background_opacity = 0.9` with the
  backdrop disabled is the cheap way to get some transparency back.
- **Window frame:** no title bar; the minimize/maximize/close buttons sit in
  the tab bar, which is always visible.
- **Not ported:** always on top (use PowerToys) and the fixed `Sean@Cerri` tab
  title.

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy when text is selected, otherwise interrupt |
| Ctrl+V | Paste |
| Ctrl+Shift+F | Search |
| Alt+Shift+D | Split along the longer side |
| Alt+Shift+Arrows | Open a new pane above, below, left or right |
| Alt+Shift+- / Alt+Shift++ | Split below / split right |
| Alt+Arrows | Move between panes |
| Alt+Shift+R | Rename the current tab (empty name resets it) |
| Ctrl+Shift+W | Close the current pane (no confirmation) |
| Alt+Shift+M | Ask for a file and open it in [mmv](#mmv) (use a full path) |

Setup on a new machine: install WezTerm and the CaskaydiaCove Nerd Font
(from nerdfonts.com), then create the stub:

```powershell
winget install wez.wezterm

@'
local wezterm = require 'wezterm'
local path = 'S:\\Personal\\wezterm\\wezterm.lua'
wezterm.add_to_config_reload_watch_list(path)
return dofile(path)
'@ | Set-Content "$HOME\.wezterm.lua"
```

WezTerm updates through winget, so its own update check is off in the config.

## mmv

`mmv` opens Mermaid diagrams too big for normal previews in a GPU-rendered,
keyboard-driven window. Mermaid still does the layout, so the first open of a
large file takes a few seconds; after that the layout comes from a cache until
the file changes.

There are three ways to open a file:

- **Terminal:** `mmv diagram.mmd`. The command holds the terminal until the
  window closes; `Start-Process mmv diagram.mmd` gives the prompt back.
- **Yazi:** press Enter on a `.mmd` or `.mermaid` file. Shift+O still offers the
  editor.
- **WezTerm:** press Alt+Shift+M and type a full path.

| Key | Action |
|-----|--------|
| `h` `j` `k` `l` or arrows | Pan; hold Shift to go faster |
| `u` `d` | Half a screen up or down |
| `+` `-` or `i` `o` | Zoom in or out |
| `f` or `0` | Fit the whole diagram |
| `1` | 100% zoom at the cursor |
| `/` | Search node and table names |
| `n` `N` | Next or previous match |
| Enter | Keep the matches and leave search |
| Esc | Cancel search and clear the matches |
| `r` | Re-render, ignoring the cache |
| `?` | Show all the keys |
| `q` | Quit |

The mouse wheel zooms at the cursor and dragging pans. Saving the file reloads
the diagram and keeps the current view.

Colors set in a diagram (`classDef`, `style`, a named Mermaid theme) always win.
Everything else uses the WezTerm palette, which `wezterm/wezterm.lua` writes out
for `mmv` when it loads. The window background stays dark, except that a diagram
naming a light Mermaid theme gets white. Text switches to dark or light when it
would not read on the fill behind it.

Install it, or rebuild it after changing the source:

```powershell
cargo install --path S:\Personal\mmv --locked
```

[`mmv/README.md`](mmv/README.md) has the prerequisites, how it works and its
limits.

## Setup on a new machine

Install the tools:

```powershell
winget install JanDeDobbeleer.OhMyPosh
winget install ajeetdsouza.zoxide
winget install junegunn.fzf
```

For the terminal itself, follow [WezTerm](#wezterm) above. For the diagram
viewer, follow [Install](mmv/README.md#install) in its README.

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
