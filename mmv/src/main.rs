//! mmv: a keyboard-driven, GPU-rendered viewer for large Mermaid diagrams.

mod app;
mod camera;
mod gpu;
mod scene;
mod source;
mod theme;

use std::ffi::{OsStr, c_void};
use std::io::IsTerminal;
use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;
use winit::event_loop::EventLoop;

pub const EXIT_FAILURE: i32 = 1;
const EXIT_USAGE: i32 = 2;
const MB_ICONERROR: u32 = 0x10;

#[link(name = "user32")]
unsafe extern "system" {
    fn MessageBoxW(window: *mut c_void, text: *const u16, caption: *const u16, kind: u32) -> i32;
}

const USAGE: &str = "\
mmv - GPU viewer for large Mermaid diagrams

usage: mmv [--control] FILE

FILE is a Mermaid source file, or an .svg to skip mermaid-cli.
Press ? in the window for the keys.

  --control   read scripted commands from stdin (see the README)
  --help      show this text
  --version   show the version";

fn main() {
    let mut input = None;
    let mut control = false;
    for arg in std::env::args_os().skip(1) {
        match arg.to_str() {
            Some("--help" | "-h") => {
                println!("{USAGE}");
                return;
            }
            Some("--version" | "-V") => {
                println!("mmv {}", env!("CARGO_PKG_VERSION"));
                return;
            }
            Some("--control") => control = true,
            _ if input.is_none() => input = Some(PathBuf::from(arg)),
            _ => {
                eprintln!("{USAGE}");
                std::process::exit(EXIT_USAGE);
            }
        }
    }

    let Some(input) = input else {
        eprintln!("{USAGE}");
        std::process::exit(EXIT_USAGE);
    };
    if let Err(problem) = check_input(&input) {
        fatal(&problem, EXIT_USAGE, control);
    }

    let event_loop = match EventLoop::<app::UserEvent>::with_user_event().build() {
        Ok(event_loop) => event_loop,
        Err(e) => fatal(
            &format!("cannot create the event loop: {e}"),
            EXIT_FAILURE,
            control,
        ),
    };
    let mut app = app::App::new(event_loop.create_proxy(), input, control);
    if let Err(e) = event_loop.run_app(&mut app) {
        fatal(&e.to_string(), EXIT_FAILURE, control);
    }
}

fn check_input(input: &std::path::Path) -> Result<(), String> {
    let shown = input.display();
    match std::fs::metadata(input) {
        Ok(meta) if meta.is_dir() => Err(format!("{shown} is a folder, not a file")),
        Ok(_) => std::fs::File::open(input)
            .map(|_| ())
            .map_err(|e| format!("cannot read {shown}: {e}")),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Err(format!("no such file: {shown}")),
        Err(e) => Err(format!("cannot read {shown}: {e}")),
    }
}

/// Reports an error and exits. yazi and WezTerm start mmv with no terminal to
/// print to, so there the error goes in a message box instead.
pub fn fatal(message: &str, code: i32, control: bool) -> ! {
    let has_terminal = std::io::stderr().is_terminal();
    eprintln!("mmv: {message}");
    if !has_terminal && !control {
        let wide = |text: &str| -> Vec<u16> {
            OsStr::new(text)
                .encode_wide()
                .chain(std::iter::once(0))
                .collect()
        };
        let (text, caption) = (wide(message), wide("mmv"));
        // SAFETY: both strings are NUL-terminated and outlive the call.
        unsafe {
            MessageBoxW(
                std::ptr::null_mut(),
                text.as_ptr(),
                caption.as_ptr(),
                MB_ICONERROR,
            );
        }
    }
    std::process::exit(code);
}
