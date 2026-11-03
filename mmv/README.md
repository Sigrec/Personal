# mmv

A keyboard-driven, GPU-rendered viewer for large Mermaid diagrams on Windows.
Run `mmv file.mmd` from the terminal and it opens a native window that draws
the diagram as vector geometry through Direct3D 12.

It exists for diagrams that normal previews refuse, such as a 104 KB ER diagram
over Mermaid's default `maxTextSize`.

## How it works

1. **Layout:** runs real Mermaid through `mmdc` (mermaid-cli) with `maxTextSize`
   and `maxEdges` raised, and caches the SVG by content hash in
   `%LOCALAPPDATA%\mmv\cache`.
2. **Geometry:** parses the SVG with `usvg` and tessellates it with `lyon`.
3. **Drawing:** `wgpu` draws the triangles with 4x MSAA. Off-screen tiles are
   skipped, and the status line and labels use `glyphon`.

There is no Mermaid parser or layout engine in here.

## Install

You need Rust, Node.js and either Edge or Chrome.

```powershell
winget install Rustlang.Rustup
```

mermaid-cli normally downloads its own Chromium. mmv points it at the installed
Edge or Chrome, so that download can be skipped:

```powershell
$env:PUPPETEER_SKIP_DOWNLOAD = 'true'
npm install -g @mermaid-js/mermaid-cli
```

Then build and install `mmv` into `~\.cargo\bin`:

```powershell
cargo install --path S:\Personal\mmv
```

## Usage

```powershell
mmv Platform_ER_Diagram.mmd
```

The input can be any Mermaid source file, or an `.svg` to skip mermaid-cli. The
command stays attached to the terminal until the window closes. To get the
prompt back straight away, use `Start-Process mmv file.mmd`.

## Keys

| Key | Action |
|-----|--------|
| `h` `j` `k` `l` or arrows | Pan; hold Shift to go faster |
| `u` `d` | Half a screen up or down |
| `+` `-` or `i` `o` | Zoom in or out |
| Wheel, left drag | Zoom at the cursor, pan |
| `f` or `0` | Fit the whole diagram |
| `1` | 100% zoom |
| `/` | Search node and table names |
| `n` `N` | Next or previous match |
| Enter | Keep the matches and leave search |
| Esc | Cancel search and clear the matches |
| `r` | Re-render, ignoring the cache |
| `?` | Toggle help |
| `q` | Quit |

Search jumps as you type: exact names first, then prefixes, then any name
containing the text. The current match is outlined in amber and the others in
blue.

When zoomed far out, the diagram's own text is too small to read and is not
drawn. Node names appear as labels instead, and any that would overlap are
hidden, largest tables first.

Saving the file re-renders it and keeps the current view. If the new version
has a syntax error, the old diagram stays up with the error in the status line.

## Yazi

Add an opener and a rule to `yazi/yazi.toml`, keeping the rules already in
`prepend_rules`. This uses the opener syntax of Yazi 26.9.

```toml
[opener]
mmv = [
  { run = "mmv %s", orphan = true, desc = "mmv", for = "windows" },
]

[open]
prepend_rules = [
  { url = "*.mmd", use = ["mmv", "edit", "reveal"] },
  { url = "*.mermaid", use = ["mmv", "edit", "reveal"] },
]
```

## WezTerm

This binding asks for a path and opens it without tying up a pane. Add it to
`config.keys` in `wezterm/wezterm.lua`. Alt+Shift+M is free in the current
config. Give it a full path, because the process does not start in the pane's
folder.

```lua
{
  key = 'm',
  mods = 'ALT|SHIFT',
  action = act.PromptInputLine {
    description = 'Mermaid file to open in mmv:',
    action = wezterm.action_callback(function(_, _, line)
      if line and line ~= '' then
        wezterm.background_child_process { 'mmv', line }
      end
    end),
  },
},
```

WezTerm has to have `~\.cargo\bin` on its PATH, so restart it after installing
Rust.

## Colours and limits

- The palette is Wild West Vibes, set in `MERMAID_CONFIG` in `src/source.rs`.
  Changing it changes the cache key, so diagrams re-render.
- Gradients draw as their first colour, translucent paints are mixed against
  the background, and drop shadows and clip paths are ignored. Mermaid diagrams
  rarely depend on these, but a hand-made `.svg` might.
- Embedded images are not drawn. Labels must be SVG text, so `htmlLabels` is
  turned off in the Mermaid config.
- Search covers node and table names, not column names or edge labels.
- Windows only.

## Environment variables

| Variable | Purpose |
|----------|---------|
| `MMV_MMDC` | Path to `mmdc.cmd`, if it is not on PATH |
| `MMV_BROWSER` | Browser for mermaid-cli, instead of Edge or Chrome |
| `MMV_PRESENT` | `immediate` turns off vsync, for measuring |

## Scripted control

`mmv --control FILE` also reads commands from stdin, one per line, and prints
results to stdout. It is there for testing and benchmarks.

| Command | Effect |
|---------|--------|
| `search TEXT` | Search and jump |
| `zoom S`, `center X Y` | Move the camera |
| `bench FRAMES` | Pan for that many frames and print frame times |
| `settled` | Print `settled` once the view stops moving |
| `stats` | Print the camera, matches and draw counts |
| `f`, `n`, `?`, `r`, `q`... | Same as the key |

## Dependencies and licences

All direct dependencies are permissively licensed, and none has a revenue or
usage condition.

| Crate | Purpose | Licence |
|-------|---------|---------|
| `wgpu` | GPU | MIT or Apache-2.0 |
| `winit` | Window and input | Apache-2.0 |
| `usvg` | SVG parsing | Apache-2.0 or MIT |
| `lyon_tessellation` | Tessellation | MIT or Apache-2.0 |
| `glyphon` | Status line and label text | MIT, Apache-2.0 or Zlib |
| `rayon` | Parallel tessellation | MIT or Apache-2.0 |
| `sha2` | Cache key | MIT or Apache-2.0 |
| `bytemuck` | Vertex data | Zlib, Apache-2.0 or MIT |
| `pollster` | Blocking on GPU setup | Apache-2.0 or MIT |

At runtime it calls mermaid-cli, which is MIT.
