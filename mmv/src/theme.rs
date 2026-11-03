//! Colours for diagrams that set none of their own, and for mmv's own UI.
//!
//! They come from the terminal's palette, read from a small file the terminal
//! config writes. Without that file mmv uses Wild West Vibes. The palette is
//! drawn on the fixed `scene::BACKGROUND`, so one with text too dark to read on
//! it is ignored.

use crate::scene::BACKGROUND;
use glyphon::Color;
use std::path::PathBuf;

pub type Rgb = [u8; 3];

const LIGHT_BACKGROUND: Rgb = [0xff, 0xff, 0xff];
const MIN_TEXT_CONTRAST: f32 = 3.0;

#[derive(Clone)]
pub struct Theme {
    foreground: Rgb,
    ansi: [Rgb; 8],
    brights: [Rgb; 8],
}

#[derive(Clone, Copy)]
pub struct TextContrast {
    pub light: Rgb,
    pub dark: Rgb,
}

impl TextContrast {
    /// Keeps `text` if it reads on `backdrop`, else the better of light and dark.
    /// A diagram can set a light fill without a text colour, which leaves the
    /// palette's light text on it.
    pub fn resolve(&self, text: Rgb, backdrop: Rgb) -> Rgb {
        if contrast(text, backdrop) >= MIN_TEXT_CONTRAST {
            text
        } else if contrast(self.dark, backdrop) > contrast(self.light, backdrop) {
            self.dark
        } else {
            self.light
        }
    }
}

pub struct UiColors {
    pub bar: [u8; 4],
    pub pill: [u8; 4],
    pub panel: [u8; 4],
    pub panel_edge: [u8; 4],
    pub current_match: [u8; 4],
    pub other_match: [u8; 4],
    pub text: Color,
    pub muted_text: Color,
    pub current_text: Color,
    pub match_text: Color,
    pub error_text: Color,
}

impl Theme {
    fn builtin() -> Self {
        Self {
            foreground: [0xe1, 0xdf, 0xe5],
            ansi: [
                [0x1d, 0x1a, 0x23],
                [0xa2, 0x5a, 0x63],
                [0x70, 0x8b, 0x91],
                [0xbf, 0xa4, 0x81],
                [0x77, 0x90, 0xa8],
                [0x8d, 0x88, 0xb6],
                [0x2c, 0x85, 0x85],
                [0x91, 0x8a, 0x9f],
            ],
            brights: [
                [0x89, 0x80, 0x98],
                [0xe1, 0x6f, 0x6f],
                [0x6b, 0xc4, 0x6d],
                [0xf0, 0xbf, 0x5e],
                [0x6c, 0x9e, 0xdd],
                [0xa4, 0x8a, 0xde],
                [0x4e, 0xb1, 0xb1],
                [0xe1, 0xdf, 0xe5],
            ],
        }
    }

    /// The terminal's palette if it has written one, else the built-in palette.
    pub fn load() -> Self {
        let builtin = Self::builtin();
        let Ok(text) = std::fs::read_to_string(palette_path()) else {
            return builtin;
        };
        let mut theme = builtin.clone();
        for line in text.lines() {
            let Some((key, value)) = line.split_once('=') else {
                continue;
            };
            let Some(color) = parse_hex(value.trim()) else {
                continue;
            };
            let key = key.trim();
            if key == "foreground" {
                theme.foreground = color;
            } else if let Some(slot) = palette_slot(&mut theme, key) {
                *slot = color;
            }
        }
        if contrast(theme.foreground, BACKGROUND) < MIN_TEXT_CONTRAST {
            return builtin;
        }
        theme
    }

    /// Mermaid config. `named_theme` means the diagram picks a Mermaid theme
    /// itself, so only the settings mmv needs are passed. Otherwise Mermaid's
    /// `base` theme gets this palette, and anything the diagram sets still wins.
    pub fn mermaid_config(&self, named_theme: bool) -> String {
        const COMMON: &str = r#""maxTextSize": 50000000,
  "maxEdges": 200000,
  "look": "classic",
  "htmlLabels": false,
  "flowchart": { "htmlLabels": false },
  "er": { "useMaxWidth": false },
  "class": { "htmlLabels": false }"#;
        // Edge labels get a solid, outlined box so they stand out from the lines.
        if named_theme {
            return format!(
                "{{\n  {COMMON},\n  \"themeCSS\": \".edgeLabel rect {{ opacity: 1; stroke: #888888; stroke-width: 1px; }}\"\n}}"
            );
        }

        let fg = hex(self.foreground);
        let bg = hex(BACKGROUND);
        let surface = hex(mix(BACKGROUND, self.foreground, 0.065));
        let surface_low = hex(mix(BACKGROUND, self.foreground, 0.03));
        let surface_lower = hex(mix(BACKGROUND, self.foreground, 0.012));
        let sunken = hex(scale(BACKGROUND, 0.72));
        let blue = hex(self.ansi[4]);
        let purple = hex(self.ansi[5]);
        let grey = hex(self.ansi[7]);
        let title = hex(self.brights[3]);
        let label_box = hex(mix(BACKGROUND, self.foreground, 0.12));
        format!(
            r#"{{
  {COMMON},
  "theme": "base",
  "themeCSS": ".edgeLabel rect {{ opacity: 1; stroke: {grey}; stroke-width: 1px; }}",
  "themeVariables": {{
    "darkMode": true,
    "background": "{bg}",
    "fontFamily": "Segoe UI, Arial, sans-serif",
    "primaryColor": "{surface}",
    "primaryBorderColor": "{blue}",
    "primaryTextColor": "{fg}",
    "secondaryColor": "{surface_low}",
    "secondaryBorderColor": "{purple}",
    "secondaryTextColor": "{fg}",
    "tertiaryColor": "{sunken}",
    "tertiaryBorderColor": "{grey}",
    "tertiaryTextColor": "{fg}",
    "lineColor": "{grey}",
    "textColor": "{fg}",
    "titleColor": "{title}",
    "mainBkg": "{surface}",
    "nodeBorder": "{blue}",
    "edgeLabelBackground": "{label_box}",
    "attributeBackgroundColorOdd": "{surface_low}",
    "attributeBackgroundColorEven": "{surface_lower}"
  }}
}}"#
        )
    }

    /// Text colours to switch to when text would not read on what is behind it.
    pub fn text_contrast(&self) -> TextContrast {
        TextContrast {
            light: self.foreground,
            dark: BACKGROUND,
        }
    }

    pub fn ui(&self) -> UiColors {
        let sunken = scale(BACKGROUND, 0.72);
        UiColors {
            bar: with_alpha(sunken, 0xff),
            pill: with_alpha(sunken, 0xd8),
            panel: with_alpha(sunken, 0xf4),
            panel_edge: with_alpha(mix(BACKGROUND, self.ansi[7], 0.75), 0xff),
            current_match: with_alpha(self.brights[3], 0xff),
            other_match: with_alpha(self.brights[4], 0xff),
            text: text_color(self.foreground),
            muted_text: text_color(self.ansi[7]),
            current_text: text_color(self.brights[3]),
            match_text: text_color(self.brights[4]),
            error_text: text_color(self.brights[1]),
        }
    }
}

/// Mermaid hard-codes white for the hollow circle on ER "zero" markers. Under
/// the terminal palette it takes the background so it reads as hollow.
pub fn recolor_svg(svg: &str) -> String {
    svg.replace(
        "<circle fill=\"white\"",
        &format!("<circle fill=\"{}\"", hex(BACKGROUND)),
    )
}

/// The Mermaid theme a diagram's front matter or `%%{...}%%` directive names.
/// `themeVariables` alone does not count: those layer on the palette.
pub fn mermaid_theme_name(source: &str) -> Option<String> {
    let source = source.trim_start_matches('\u{feff}');
    let mut lines = source.lines();
    if lines.next().is_some_and(|first| first.trim() == "---") {
        for line in lines {
            let line = line.trim();
            if line == "---" {
                break;
            }
            if line.starts_with("theme:")
                && let Some(name) = theme_value(line)
            {
                return Some(name);
            }
        }
    }

    let mut rest = source;
    while let Some(start) = rest.find("%%{") {
        let after = &rest[start + 3..];
        let end = after.find("}%%").unwrap_or(after.len());
        if let Some(name) = theme_value(&after[..end]) {
            return Some(name);
        }
        rest = &after[end..];
    }
    None
}

/// The window background for a diagram. Mermaid's themes other than `dark` are
/// designed for a light page; everything else gets the fixed dark background.
pub fn background_for(theme_name: Option<&str>) -> Rgb {
    match theme_name {
        Some(name) if !name.eq_ignore_ascii_case("dark") => LIGHT_BACKGROUND,
        _ => BACKGROUND,
    }
}

/// Finds a `theme` key, quoted or not, and returns its value.
fn theme_value(text: &str) -> Option<String> {
    let is_quote = |c: char| c == '"' || c == '\'';
    let mut rest = text;
    while let Some(at) = rest.find("theme") {
        rest = &rest[at + "theme".len()..];
        let Some(value) = rest
            .trim_start_matches(is_quote)
            .trim_start()
            .strip_prefix(':')
        else {
            continue;
        };
        let name: String = value
            .trim_start()
            .trim_start_matches(is_quote)
            .chars()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == '-' || *c == '_')
            .collect();
        if !name.is_empty() {
            return Some(name);
        }
    }
    None
}

fn palette_path() -> PathBuf {
    std::env::var_os("MMV_PALETTE").map_or_else(
        || crate::source::data_dir().join("terminal-palette.txt"),
        PathBuf::from,
    )
}

fn palette_slot<'a>(theme: &'a mut Theme, key: &str) -> Option<&'a mut Rgb> {
    let (group, index) = if let Some(index) = key.strip_prefix("ansi") {
        (&mut theme.ansi, index)
    } else {
        (&mut theme.brights, key.strip_prefix("bright")?)
    };
    group.get_mut(index.parse::<usize>().ok()?)
}

fn parse_hex(value: &str) -> Option<Rgb> {
    let digits = value.strip_prefix('#')?;
    if digits.len() != 6 || !digits.is_ascii() {
        return None;
    }
    let channel = |at: usize| u8::from_str_radix(&digits[at..at + 2], 16).ok();
    Some([channel(0)?, channel(2)?, channel(4)?])
}

fn hex(color: Rgb) -> String {
    format!("#{:02x}{:02x}{:02x}", color[0], color[1], color[2])
}

fn mix(from: Rgb, to: Rgb, amount: f32) -> Rgb {
    let channel = |i: usize| {
        (f32::from(from[i]) + (f32::from(to[i]) - f32::from(from[i])) * amount).round() as u8
    };
    [channel(0), channel(1), channel(2)]
}

fn scale(color: Rgb, amount: f32) -> Rgb {
    mix([0, 0, 0], color, amount)
}

fn with_alpha(color: Rgb, alpha: u8) -> [u8; 4] {
    [color[0], color[1], color[2], alpha]
}

fn text_color(color: Rgb) -> Color {
    Color::rgb(color[0], color[1], color[2])
}

/// WCAG contrast ratio between two sRGB colours.
fn contrast(a: Rgb, b: Rgb) -> f32 {
    let luminance = |c: Rgb| {
        let linear = |v: u8| {
            let v = f32::from(v) / 255.0;
            if v <= 0.04045 {
                v / 12.92
            } else {
                ((v + 0.055) / 1.055).powf(2.4)
            }
        };
        0.2126 * linear(c[0]) + 0.7152 * linear(c[1]) + 0.0722 * linear(c[2])
    };
    let (la, lb) = (luminance(a), luminance(b));
    (la.max(lb) + 0.05) / (la.min(lb) + 0.05)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn name(source: &str) -> Option<String> {
        mermaid_theme_name(source)
    }

    #[test]
    fn reads_theme_from_front_matter() {
        let source = "---\nconfig:\n  theme: forest\n---\nflowchart LR\n";
        assert_eq!(name(source).as_deref(), Some("forest"));
        let quoted = "---\nconfig:\n  theme: \"neutral\"\n---\nflowchart LR\n";
        assert_eq!(name(quoted).as_deref(), Some("neutral"));
    }

    #[test]
    fn reads_theme_from_directive() {
        let double = "%%{init: {\"theme\": \"neutral\"}}%%\nerDiagram\n";
        assert_eq!(name(double).as_deref(), Some("neutral"));
        let single = "%%{ init: { 'theme': 'forest' } }%%\nflowchart LR\n";
        assert_eq!(name(single).as_deref(), Some("forest"));
        let bare = "%%{init: {theme: dark}}%%\nflowchart LR\n";
        assert_eq!(name(bare).as_deref(), Some("dark"));
    }

    #[test]
    fn theme_variables_are_not_a_named_theme() {
        let front = "---\nconfig:\n  themeVariables:\n    lineColor: \"#00ff00\"\n---\nflowchart\n";
        assert_eq!(name(front), None);
        let directive = "%%{init: {\"themeVariables\": {\"lineColor\": \"#0f0\"}}}%%\nflowchart\n";
        assert_eq!(name(directive), None);
    }

    #[test]
    fn title_and_body_text_are_not_a_named_theme() {
        let title = "---\ntitle: theme: forest\n---\nflowchart LR\n  A[theme: forest]\n";
        assert_eq!(name(title), None);
        assert_eq!(name("flowchart LR\n  A[theme: forest]\n"), None);
    }

    #[test]
    fn light_text_on_a_light_fill_turns_dark() {
        let pair = Theme::builtin().text_contrast();
        let on_light = pair.resolve(pair.light, [0xff, 0xe0, 0xb2]);
        assert_eq!(on_light, pair.dark);
    }

    #[test]
    fn readable_text_keeps_its_colour() {
        let pair = Theme::builtin().text_contrast();
        let yellow = [0xff, 0xff, 0x00];
        assert_eq!(pair.resolve(yellow, [0xaa, 0x00, 0x00]), yellow);
        assert_eq!(pair.resolve(pair.light, BACKGROUND), pair.light);
    }

    #[test]
    fn only_dark_and_unnamed_stay_dark() {
        assert_eq!(background_for(None), BACKGROUND);
        assert_eq!(background_for(Some("dark")), BACKGROUND);
        for light in ["default", "forest", "neutral", "base"] {
            assert_eq!(background_for(Some(light)), LIGHT_BACKGROUND);
        }
    }
}
