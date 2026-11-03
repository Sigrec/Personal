-- WezTerm config, ported from the live Windows Terminal settings
-- (Wild West Vibes scheme, CaskaydiaCove Nerd Font, same shortcuts).
-- ~/.wezterm.lua loads this file, so edit it here and the repo stays the source.
local wezterm = require 'wezterm'
local act = wezterm.action
local config = wezterm.config_builder()

-- Shell: PowerShell 7 without the banner. PowerShell is the Store (MSIX) install,
-- so there is no pwsh.exe under Program Files; resolve it through PATH instead.
config.default_prog = { 'pwsh.exe', '-NoLogo' }

-- Font
-- The installed family has no Medium file, and WezTerm's renderer draws thinner
-- than Windows Terminal's, so Regular looks light. DemiBold maps to the SemiBold
-- file; drop back to 'Regular' if it feels too heavy.
config.font = wezterm.font('CaskaydiaCove Nerd Font', { weight = 'DemiBold' })
config.font_size = 10.0

-- Colors: Wild West Vibes (the same scheme the powerline prompt is built on)
local background = '#1D1A23'
local foreground = '#E1DFE5'

config.colors = {
  foreground = foreground,
  background = background,
  cursor_bg = '#ABA6B6',
  cursor_fg = background,
  cursor_border = '#ABA6B6',
  selection_bg = '#564F64',
  selection_fg = foreground,
  scrollbar_thumb = '#564F64',
  ansi = {
    '#1D1A23', -- black
    '#A25A63', -- red
    '#708B91', -- green
    '#BFA481', -- yellow
    '#7790A8', -- blue
    '#8D88B6', -- purple
    '#2C8585', -- cyan
    '#918A9F', -- white
  },
  brights = {
    '#898098',
    '#E16F6F',
    '#6BC46D',
    '#F0BF5E',
    '#6C9EDD',
    '#A48ADE',
    '#4EB1B1',
    '#E1DFE5',
  },
  -- Tab bar uses the powerline prompt palette
  tab_bar = {
    background = '#15121a',
    active_tab = { bg_color = '#bfa481', fg_color = '#1e1b24' },
    inactive_tab = { bg_color = background, fg_color = '#918A9F' },
    inactive_tab_hover = { bg_color = '#2a2533', fg_color = foreground },
    new_tab = { bg_color = background, fg_color = '#918A9F' },
    new_tab_hover = { bg_color = '#2a2533', fg_color = foreground },
  },
}

-- Hand the palette to mmv, which uses it for diagrams that set no colors of
-- their own. mmv creates the folder the first time it runs.
local mmv_palette = io.open(os.getenv('LOCALAPPDATA') .. '\\mmv\\terminal-palette.txt', 'w')
if mmv_palette then
  mmv_palette:write('foreground=', foreground, '\n')
  for i, color in ipairs(config.colors.ansi) do
    mmv_palette:write('ansi', i - 1, '=', color, '\n')
  end
  for i, color in ipairs(config.colors.brights) do
    mmv_palette:write('bright', i - 1, '=', color, '\n')
  end
  mmv_palette:close()
end

-- Window
-- Solid window. The Acrylic blur (win32_system_backdrop = 'Acrylic' with an opacity
-- below 1.0) made Windows recomposite every frame and added input lag. A plain
-- window_background_opacity of about 0.9 with the backdrop disabled is the cheap
-- way to get some transparency back.
config.win32_system_backdrop = 'Disable'
config.window_background_opacity = 1.0

-- Rendering: WebGPU on the discrete GPU. If anything glitches after a driver
-- update, remove front_end to fall back to the default OpenGL renderer.
-- The main monitor is 240 Hz (the second one is 60 Hz), so allow up to 240 fps.
config.front_end = 'WebGpu'
config.webgpu_power_preference = 'HighPerformance'
config.max_fps = 240
config.animation_fps = 60
-- No title bar: the minimize/maximize/close buttons move into the tab bar, which
-- therefore has to stay visible even with a single tab.
config.window_decorations = 'INTEGRATED_BUTTONS|RESIZE'
config.window_padding = { left = 8, right = 8, top = 4, bottom = 4 }
config.enable_scroll_bar = true
config.scrollback_lines = 10000
config.default_cursor_style = 'SteadyBar' -- no blink, so an idle terminal never redraws

-- Tabs: always shown, because the window buttons live in the tab bar
config.use_fancy_tab_bar = false
config.hide_tab_bar_if_only_one_tab = false
config.tab_bar_at_bottom = false

-- winget updates WezTerm (see upall), so skip its own update check
config.check_for_updates = false

-- Split along the longer side, like Windows Terminal's "auto" split.
-- A character cell is about twice as tall as it is wide.
local function split_auto(window, pane)
  local dims = pane:get_dimensions()
  local direction = dims.cols > dims.viewport_rows * 2 and 'Right' or 'Down'
  window:perform_action(act.SplitPane { direction = direction }, pane)
end

-- Ctrl+C copies when text is selected, otherwise it is the interrupt key
local function copy_or_interrupt(window, pane)
  if window:get_selection_text_for_pane(pane) ~= '' then
    window:perform_action(act.CopyTo 'ClipboardAndPrimarySelection', pane)
    window:perform_action(act.ClearSelection, pane)
  else
    window:perform_action(act.SendKey { key = 'c', mods = 'CTRL' }, pane)
  end
end

config.keys = {
  { key = 'c', mods = 'CTRL', action = wezterm.action_callback(copy_or_interrupt) },
  { key = 'v', mods = 'CTRL', action = act.PasteFrom 'Clipboard' },
  { key = 'f', mods = 'CTRL|SHIFT', action = act.Search 'CurrentSelectionOrEmptyString' },

  -- Rename the current tab; an empty name goes back to the automatic title
  {
    key = 'r',
    mods = 'ALT|SHIFT',
    action = act.PromptInputLine {
      description = 'Tab name (empty to reset)',
      action = wezterm.action_callback(function(window, _, line)
        if line then
          window:active_tab():set_title(line)
        end
      end),
    },
  },

  -- Open a Mermaid file in mmv; give it a full path
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

  -- Panes (same shortcuts as Windows Terminal)
  { key = 'd', mods = 'ALT|SHIFT', action = wezterm.action_callback(split_auto) },
  { key = 'phys:Minus', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Down' } },
  { key = 'phys:Equal', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Right' } },
  -- Alt+Shift+arrows open a new pane on that side of the current one
  { key = 'UpArrow', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Up' } },
  { key = 'DownArrow', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Down' } },
  { key = 'LeftArrow', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Left' } },
  { key = 'RightArrow', mods = 'ALT|SHIFT', action = act.SplitPane { direction = 'Right' } },
  { key = 'w', mods = 'CTRL|SHIFT', action = act.CloseCurrentPane { confirm = false } },
  { key = 'LeftArrow', mods = 'ALT', action = act.ActivatePaneDirection 'Left' },
  { key = 'RightArrow', mods = 'ALT', action = act.ActivatePaneDirection 'Right' },
  { key = 'UpArrow', mods = 'ALT', action = act.ActivatePaneDirection 'Up' },
  { key = 'DownArrow', mods = 'ALT', action = act.ActivatePaneDirection 'Down' },
}

return config
