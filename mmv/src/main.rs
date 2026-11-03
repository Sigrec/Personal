//! mmv: a keyboard-driven, GPU-rendered viewer for large Mermaid diagrams.

mod app;
mod camera;
mod gpu;
mod scene;
mod source;

use std::path::PathBuf;
use winit::event_loop::EventLoop;

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
                std::process::exit(2);
            }
        }
    }

    let Some(input) = input else {
        eprintln!("{USAGE}");
        std::process::exit(2);
    };
    if !input.is_file() {
        eprintln!("mmv: no such file: {}", input.display());
        std::process::exit(2);
    }

    let event_loop = match EventLoop::<app::UserEvent>::with_user_event().build() {
        Ok(event_loop) => event_loop,
        Err(e) => {
            eprintln!("mmv: cannot create the event loop: {e}");
            std::process::exit(1);
        }
    };
    let mut app = app::App::new(event_loop.create_proxy(), input, control);
    if let Err(e) = event_loop.run_app(&mut app) {
        eprintln!("mmv: {e}");
        std::process::exit(1);
    }
}
