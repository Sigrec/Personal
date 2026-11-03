//! Turns an input file into an SVG on disk: runs mermaid-cli once per content
//! hash and keeps the result in the cache folder.

use crate::theme::{self, Rgb, TextContrast, Theme};
use sha2::{Digest, Sha256};
use std::fmt::Write as _;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;

const CREATE_NO_WINDOW: u32 = 0x0800_0000;

const BROWSERS: [&str; 3] = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
];

pub struct SvgFile {
    pub path: PathBuf,
    /// True when this load ran mermaid-cli rather than using the cache.
    pub rendered: bool,
    pub background: Rgb,
    /// Set under the terminal palette. A theme or SVG the diagram chose is left as it is.
    pub text_contrast: Option<TextContrast>,
}

pub fn data_dir() -> PathBuf {
    let base = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir);
    base.join("mmv")
}

/// Returns the SVG for `input`. `force` ignores any cached copy.
pub fn svg_for(input: &Path, force: bool, status: &dyn Fn(String)) -> Result<SvgFile, String> {
    let is_svg = input
        .extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("svg"));
    if is_svg {
        return Ok(SvgFile {
            path: input.to_path_buf(),
            rendered: false,
            background: theme::background_for(None),
            text_contrast: None,
        });
    }

    let bytes = std::fs::read(input).map_err(|e| format!("read {}: {e}", input.display()))?;
    let dir = data_dir().join("cache");
    std::fs::create_dir_all(&dir).map_err(|e| format!("create {}: {e}", dir.display()))?;

    let theme_name = theme::mermaid_theme_name(&String::from_utf8_lossy(&bytes));
    let background = theme::background_for(theme_name.as_deref());
    let named_theme = theme_name.is_some();
    let palette = Theme::load();
    let config = palette.mermaid_config(named_theme);
    let text_contrast = (!named_theme).then(|| palette.text_contrast());

    let mut hasher = Sha256::new();
    hasher.update(config.as_bytes());
    hasher.update(&bytes);
    let mut name = String::new();
    for byte in hasher.finalize().iter().take(16) {
        let _ = write!(name, "{byte:02x}");
    }

    let svg = dir.join(format!("{name}.svg"));
    if svg.exists() && !force {
        return Ok(SvgFile {
            path: svg,
            rendered: false,
            background,
            text_contrast,
        });
    }

    status("laying out with mermaid-cli...".to_string());
    run_mmdc(input, &svg, &dir, &name, &config, named_theme)?;
    Ok(SvgFile {
        path: svg,
        rendered: true,
        background,
        text_contrast,
    })
}

fn run_mmdc(
    input: &Path,
    svg: &Path,
    dir: &Path,
    name: &str,
    config_text: &str,
    named_theme: bool,
) -> Result<(), String> {
    let config = dir.join("mermaid-config.json");
    std::fs::write(&config, config_text).map_err(|e| format!("write config: {e}"))?;

    // mmdc picks the renderer from the extension, so the temp name ends in .svg.
    let partial = dir.join(format!("{name}.partial.svg"));
    let mmdc = find_mmdc()
        .ok_or("mermaid-cli not found. Install it with: npm install -g @mermaid-js/mermaid-cli")?;

    let mut command = Command::new(&mmdc);
    command
        .arg("-q")
        .arg("-i")
        .arg(input)
        .arg("-o")
        .arg(&partial)
        .arg("-c")
        .arg(&config)
        .arg("-b")
        .arg("transparent")
        .creation_flags(CREATE_NO_WINDOW);

    if let Some(browser) = browser_path() {
        let puppeteer = dir.join("puppeteer-config.json");
        let body = format!(
            "{{\"executablePath\":\"{}\",\"headless\":\"shell\"}}",
            browser.replace('\\', "/")
        );
        std::fs::write(&puppeteer, body).map_err(|e| format!("write config: {e}"))?;
        command.arg("-p").arg(&puppeteer);
    }

    let output = command
        .output()
        .map_err(|e| format!("cannot run {}: {e}", mmdc.display()))?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        eprintln!("{stderr}");
        let reason = stderr
            .lines()
            .find(|l| l.contains("Error"))
            .or_else(|| stderr.lines().rfind(|l| !l.trim().is_empty()))
            .unwrap_or("no error text");
        return Err(format!("mermaid-cli failed: {}", reason.trim()));
    }
    if !named_theme {
        let text = std::fs::read_to_string(&partial).map_err(|e| format!("read svg: {e}"))?;
        std::fs::write(&partial, theme::recolor_svg(&text))
            .map_err(|e| format!("write svg: {e}"))?;
    }
    std::fs::rename(&partial, svg).map_err(|e| format!("store svg: {e}"))
}

/// `MMV_MMDC` wins; otherwise `mmdc.cmd` from PATH, where npm puts it.
fn find_mmdc() -> Option<PathBuf> {
    if let Some(path) = std::env::var_os("MMV_MMDC") {
        return Some(PathBuf::from(path));
    }
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .map(|dir| dir.join("mmdc.cmd"))
        .find(|candidate| candidate.is_file())
}

fn browser_path() -> Option<String> {
    if let Ok(browser) = std::env::var("MMV_BROWSER") {
        return Some(browser);
    }
    BROWSERS
        .iter()
        .find(|p| Path::new(p).exists())
        .map(|p| (*p).to_string())
}
