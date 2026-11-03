//! The window: input handling, search, file watching and per-frame UI.

use crate::camera::Camera;
use crate::gpu::{DrawStats, Renderer, UiVertex, ViewState};
use crate::scene::{self, Scene};
use crate::source;
use glyphon::{Attrs, Buffer, Color, Family, Metrics, Shaping, TextArea, TextBounds};
use std::collections::HashSet;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::{Duration, Instant, SystemTime};
use winit::application::ApplicationHandler;
use winit::dpi::LogicalSize;
use winit::event::{ElementState, KeyEvent, MouseButton, MouseScrollDelta, WindowEvent};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoopProxy};
use winit::keyboard::{Key, KeyCode, ModifiersState, NamedKey, PhysicalKey};
use winit::window::{Theme, Window, WindowId};

const UI_FONT: &str = "CaskaydiaCove Nerd Font";
const MAX_SCALE: f64 = 8.0;
/// Below this zoom the diagram's own text is too small to read, so it is not drawn.
const TEXT_MIN_SCALE: f64 = 0.28;
/// Below this zoom node names also appear as screen-sized labels.
const LABEL_MAX_SCALE: f64 = 0.45;
const PAN_PX_PER_SEC: f64 = 1100.0;
const ZOOM_PER_SEC: f64 = 2.0;
const WHEEL_STEP: f64 = 1.18;
const POLL_EVERY: Duration = Duration::from_millis(500);
const STATUS_PX: f32 = 26.0;
const LABEL_PX: f32 = 13.0;
const STATUS_FONT_PX: f32 = 13.5;

const BAR: [u8; 4] = [0x15, 0x12, 0x1a, 0xff];
const PILL: [u8; 4] = [0x15, 0x12, 0x1a, 0xd8];
const PANEL: [u8; 4] = [0x15, 0x12, 0x1a, 0xf4];
const BLUE: [u8; 4] = [0x6c, 0x9e, 0xdd, 0xff];
const AMBER: [u8; 4] = [0xf0, 0xbf, 0x5e, 0xff];
const FG: Color = Color::rgb(0xe1, 0xdf, 0xe5);
const MUTED: Color = Color::rgb(0x91, 0x8a, 0x9f);
const AMBER_TEXT: Color = Color::rgb(0xf0, 0xbf, 0x5e);
const RED_TEXT: Color = Color::rgb(0xe1, 0x6f, 0x6f);

const HELP: &str = "\
h j k l / arrows   pan (Shift = faster)
u d                half a screen up / down
+ - / i o          zoom in / out
wheel, drag        zoom at cursor, pan
f or 0             fit the whole diagram
1                  100% zoom
/                  search node names
n N                next / previous match
Enter              keep the matches, leave search
Esc                cancel search, clear matches
r                  re-render, ignoring the cache
?                  toggle this help
q                  quit";

pub enum UserEvent {
    Status(u64, String),
    Loaded(u64, Result<Box<Loaded>, String>),
    Control(String),
}

pub struct Loaded {
    scene: Scene,
    svg_seconds: f64,
    rendered: bool,
    build_seconds: f64,
}

pub struct App {
    proxy: EventLoopProxy<UserEvent>,
    input: PathBuf,
    control: bool,
    viewer: Option<Viewer>,
}

impl App {
    pub fn new(proxy: EventLoopProxy<UserEvent>, input: PathBuf, control: bool) -> Self {
        Self {
            proxy,
            input,
            control,
            viewer: None,
        }
    }
}

impl ApplicationHandler<UserEvent> for App {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        if self.viewer.is_some() {
            return;
        }
        let file_name = self
            .input
            .file_name()
            .map_or_else(String::new, |n| n.to_string_lossy().into_owned());
        let attributes = Window::default_attributes()
            .with_title(format!("mmv - {file_name}"))
            .with_inner_size(LogicalSize::new(1400.0, 900.0))
            .with_theme(Some(Theme::Dark));
        let window = match event_loop.create_window(attributes) {
            Ok(window) => Arc::new(window),
            Err(e) => {
                eprintln!("mmv: cannot create the window: {e}");
                event_loop.exit();
                return;
            }
        };
        window.set_maximized(true);
        let renderer = match Renderer::new(window.clone(), event_loop) {
            Ok(renderer) => renderer,
            Err(e) => {
                eprintln!("mmv: {e}");
                event_loop.exit();
                return;
            }
        };
        if self.control {
            spawn_control_reader(self.proxy.clone());
        }
        let mut viewer = Viewer::new(
            window,
            renderer,
            self.proxy.clone(),
            self.input.clone(),
            file_name,
            self.control,
        );
        viewer.start_load(false);
        self.viewer = Some(viewer);
    }

    fn user_event(&mut self, event_loop: &ActiveEventLoop, event: UserEvent) {
        if let Some(viewer) = &mut self.viewer {
            viewer.user_event(event_loop, event);
        }
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, _id: WindowId, event: WindowEvent) {
        if let Some(viewer) = &mut self.viewer {
            viewer.window_event(event_loop, event);
        }
    }

    fn about_to_wait(&mut self, event_loop: &ActiveEventLoop) {
        if let Some(viewer) = &mut self.viewer {
            viewer.poll_file();
            event_loop.set_control_flow(ControlFlow::WaitUntil(viewer.next_poll));
        }
    }
}

fn spawn_control_reader(proxy: EventLoopProxy<UserEvent>) {
    std::thread::spawn(move || {
        for line in std::io::stdin().lines().map_while(Result::ok) {
            if proxy.send_event(UserEvent::Control(line)).is_err() {
                break;
            }
        }
    });
}

struct Label {
    buffer: Buffer,
    width: f32,
    height: f32,
}

enum Slot {
    Label(usize),
    StatusLeft,
    StatusRight,
    Help,
}

struct Placement {
    slot: Slot,
    x: f32,
    y: f32,
    color: Color,
}

struct Bench {
    frames_left: u32,
    samples: Vec<f64>,
    origin: [f64; 2],
    clock: f64,
    last_stats: DrawStats,
}

struct Viewer {
    window: Arc<Window>,
    renderer: Renderer,
    proxy: EventLoopProxy<UserEvent>,
    input: PathBuf,
    file_name: String,
    control: bool,

    camera: Camera,
    bounds: Option<([f32; 2], [f32; 2])>,
    needs_fit: bool,
    nodes: Vec<scene::SearchNode>,
    labels: Vec<Label>,
    /// Node indices, largest first: big tables win when labels would overlap.
    label_order: Vec<usize>,

    held: HashSet<KeyCode>,
    modifiers: ModifiersState,
    cursor: [f64; 2],
    dragging: bool,

    searching: bool,
    query: String,
    matches: Vec<usize>,
    current: usize,
    show_help: bool,

    status_left: Buffer,
    status_right: Buffer,
    help: Buffer,
    status_left_text: String,
    status_right_text: String,
    info: String,
    error: Option<String>,

    load_generation: u64,
    loading: bool,
    watched_mtime: Option<SystemTime>,
    next_poll: Instant,

    last_frame: Instant,
    frame_ms: f64,
    last_stats: DrawStats,
    bench: Option<Bench>,
    report_when_settled: bool,
}

impl Viewer {
    fn new(
        window: Arc<Window>,
        mut renderer: Renderer,
        proxy: EventLoopProxy<UserEvent>,
        input: PathBuf,
        file_name: String,
        control: bool,
    ) -> Self {
        let dpi = window.scale_factor() as f32;
        let fonts = &mut renderer.font_system;
        let status_left = make_buffer(fonts, "", STATUS_FONT_PX * dpi);
        let status_right = make_buffer(fonts, "", STATUS_FONT_PX * dpi);
        let help = make_buffer(fonts, HELP, 14.5 * dpi);
        let watched_mtime = modified(&input);
        window.request_redraw();
        Self {
            window,
            renderer,
            proxy,
            input,
            file_name,
            control,
            camera: Camera::new(),
            bounds: None,
            needs_fit: true,
            nodes: Vec::new(),
            labels: Vec::new(),
            label_order: Vec::new(),
            held: HashSet::new(),
            modifiers: ModifiersState::empty(),
            cursor: [0.0, 0.0],
            dragging: false,
            searching: false,
            query: String::new(),
            matches: Vec::new(),
            current: 0,
            show_help: false,
            status_left,
            status_right,
            help,
            status_left_text: String::new(),
            status_right_text: String::new(),
            info: String::new(),
            error: None,
            load_generation: 0,
            loading: false,
            watched_mtime,
            next_poll: Instant::now() + POLL_EVERY,
            last_frame: Instant::now(),
            frame_ms: 0.0,
            last_stats: DrawStats::default(),
            bench: None,
            report_when_settled: false,
        }
    }

    // ---- loading ----------------------------------------------------------

    /// Lays out and tessellates on a worker. The old diagram stays on screen
    /// until the new one is ready, and a newer load wins over an older one.
    fn start_load(&mut self, force: bool) {
        self.load_generation += 1;
        self.loading = true;
        self.info = "loading...".to_string();
        let generation = self.load_generation;
        let proxy = self.proxy.clone();
        let input = self.input.clone();
        std::thread::spawn(move || {
            let status_proxy = proxy.clone();
            let status = move |text: String| {
                let _ = status_proxy.send_event(UserEvent::Status(generation, text));
            };
            let result = load(&input, force, &status).map(Box::new);
            let _ = proxy.send_event(UserEvent::Loaded(generation, result));
        });
        self.window.request_redraw();
    }

    fn poll_file(&mut self) {
        let now = Instant::now();
        if now < self.next_poll {
            return;
        }
        self.next_poll = now + POLL_EVERY;
        // A missing file is usually an editor's atomic save in progress.
        let Some(mtime) = modified(&self.input) else {
            return;
        };
        if self.watched_mtime != Some(mtime) {
            self.watched_mtime = Some(mtime);
            self.start_load(false);
        }
    }

    fn accept_scene(&mut self, loaded: Loaded) {
        let upload_start = Instant::now();
        let scene = loaded.scene;
        if let Err(e) = self.renderer.upload(&scene) {
            self.set_error(e);
            return;
        }
        let upload_seconds = upload_start.elapsed().as_secs_f64();

        self.bounds = Some((scene.min, scene.max));
        let dpi = self.window.scale_factor() as f32;
        self.labels = scene
            .nodes
            .iter()
            .map(|node| {
                let buffer =
                    make_buffer(&mut self.renderer.font_system, &node.name, LABEL_PX * dpi);
                let (width, height) = measure(&buffer);
                Label {
                    buffer,
                    width,
                    height,
                }
            })
            .collect();
        self.nodes = scene.nodes;
        self.label_order = (0..self.nodes.len()).collect();
        let area = |n: &scene::SearchNode| (n.max[0] - n.min[0]) * (n.max[1] - n.min[1]);
        self.label_order
            .sort_by(|a, b| area(&self.nodes[*b]).total_cmp(&area(&self.nodes[*a])));

        self.error = None;
        let layout = if loaded.rendered { "mmdc" } else { "cached" };
        self.info = format!(
            "{layout} {:.1}s, mesh {:.1}s, upload {:.2}s",
            loaded.svg_seconds, loaded.build_seconds, upload_seconds
        );
        if self.control {
            println!(
                "loaded nodes={} triangles={} shape_vertices={} text_vertices={} skipped_paths={} {}",
                self.nodes.len(),
                self.renderer.total_triangles(),
                scene.shape_vertices.len(),
                scene.text_vertices.len(),
                scene.skipped_paths,
                self.info
            );
        }

        if self.needs_fit {
            self.needs_fit = false;
            self.fit();
            self.camera.snap();
        }
        self.refresh_matches(false);
    }

    fn set_error(&mut self, message: String) {
        eprintln!("mmv: {message}");
        if self.control {
            println!("error {message}");
        }
        self.error = Some(message);
    }

    // ---- events -----------------------------------------------------------

    fn user_event(&mut self, event_loop: &ActiveEventLoop, event: UserEvent) {
        match event {
            UserEvent::Status(generation, text) => {
                if generation == self.load_generation {
                    self.info = text;
                }
            }
            UserEvent::Loaded(generation, result) => {
                if generation != self.load_generation {
                    return;
                }
                self.loading = false;
                match result {
                    Ok(loaded) => self.accept_scene(*loaded),
                    Err(e) => {
                        self.info.clear();
                        self.set_error(e);
                    }
                }
            }
            UserEvent::Control(line) => self.control_command(event_loop, &line),
        }
        self.window.request_redraw();
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => event_loop.exit(),
            WindowEvent::Resized(size) => {
                self.renderer.resize(size.width, size.height);
                self.window.request_redraw();
            }
            WindowEvent::RedrawRequested => self.redraw(),
            WindowEvent::Focused(false) => {
                self.held.clear();
                self.dragging = false;
            }
            WindowEvent::ModifiersChanged(modifiers) => self.modifiers = modifiers.state(),
            WindowEvent::KeyboardInput { event, .. } => self.key_event(event_loop, &event),
            WindowEvent::CursorMoved { position, .. } => {
                let next = [position.x, position.y];
                if self.dragging {
                    let scale = self.camera.scale;
                    self.camera.target_center[0] -= (next[0] - self.cursor[0]) / scale;
                    self.camera.target_center[1] -= (next[1] - self.cursor[1]) / scale;
                    self.camera.center = self.camera.target_center;
                    self.window.request_redraw();
                }
                self.cursor = next;
            }
            WindowEvent::MouseInput {
                state,
                button: MouseButton::Left,
                ..
            } => self.dragging = state == ElementState::Pressed,
            WindowEvent::MouseWheel { delta, .. } => {
                let steps = match delta {
                    MouseScrollDelta::LineDelta(_, y) => f64::from(y),
                    MouseScrollDelta::PixelDelta(p) => p.y / 60.0,
                };
                let (width, height) = self.renderer.size();
                let anchor = [
                    self.cursor[0] - f64::from(width) * 0.5,
                    self.cursor[1] - f64::from(height) * 0.5,
                ];
                let scale = self.clamp_scale(self.camera.target_scale * WHEEL_STEP.powf(steps));
                self.camera.zoom_at(anchor, scale);
                self.window.request_redraw();
            }
            _ => {}
        }
    }

    fn key_event(&mut self, event_loop: &ActiveEventLoop, event: &KeyEvent) {
        let code = match event.physical_key {
            PhysicalKey::Code(code) => Some(code),
            PhysicalKey::Unidentified(_) => None,
        };
        if event.state == ElementState::Released {
            if let Some(code) = code {
                self.held.remove(&code);
            }
            return;
        }

        if self.searching {
            self.search_key(event);
            self.window.request_redraw();
            return;
        }

        if let Some(code) = code.filter(|c| is_motion_key(*c)) {
            self.held.insert(code);
            self.last_frame = Instant::now();
        }
        if !event.repeat {
            match &event.logical_key {
                Key::Named(NamedKey::Escape) => self.command(event_loop, "esc"),
                Key::Character(text) => {
                    let text = text.to_string();
                    self.command(event_loop, &text);
                }
                _ => {}
            }
        }
        self.window.request_redraw();
    }

    fn search_key(&mut self, event: &KeyEvent) {
        match &event.logical_key {
            Key::Named(NamedKey::Escape) => {
                self.searching = false;
                self.query.clear();
                self.matches.clear();
            }
            Key::Named(NamedKey::Enter) => self.searching = false,
            Key::Named(NamedKey::Backspace) => {
                self.query.pop();
                self.refresh_matches(true);
            }
            _ => {
                if let Some(text) = event
                    .text
                    .as_ref()
                    .filter(|t| t.chars().all(|c| !c.is_control()))
                {
                    self.query.push_str(text);
                    self.refresh_matches(true);
                }
            }
        }
    }

    /// One-shot commands. Shared by the keyboard and `--control` mode.
    fn command(&mut self, event_loop: &ActiveEventLoop, name: &str) {
        let (_, height) = self.renderer.size();
        match name {
            "q" => event_loop.exit(),
            "esc" => {
                self.show_help = false;
                self.query.clear();
                self.matches.clear();
            }
            "/" => {
                self.searching = true;
                self.show_help = false;
                self.query.clear();
                self.matches.clear();
            }
            "n" => self.step_match(1),
            "N" => self.step_match(-1),
            "f" | "F" | "0" => self.fit(),
            "1" => self.camera.target_scale = self.clamp_scale(1.0),
            "u" | "d" => {
                let half_screen = f64::from(height) * 0.5 / self.camera.target_scale;
                let sign = if name == "u" { -1.0 } else { 1.0 };
                self.camera.target_center[1] += sign * half_screen;
            }
            "r" => self.start_load(true),
            "?" => self.show_help = !self.show_help,
            _ => {}
        }
    }

    /// `--control` mode lines: any command name, or `search TEXT`, `zoom S`,
    /// `center X Y`, `bench FRAMES`, `settled`, `stats`.
    fn control_command(&mut self, event_loop: &ActiveEventLoop, line: &str) {
        let line = line.trim();
        let (verb, rest) = line.split_once(' ').unwrap_or((line, ""));
        match verb {
            "search" => {
                self.query = rest.to_string();
                self.refresh_matches(true);
            }
            "zoom" => {
                if let Ok(scale) = rest.parse::<f64>() {
                    self.camera.target_scale = self.clamp_scale(scale);
                }
            }
            "center" => {
                let mut parts = rest.split_ascii_whitespace().map(str::parse::<f64>);
                if let (Some(Ok(x)), Some(Ok(y))) = (parts.next(), parts.next()) {
                    self.camera.target_center = [x, y];
                }
            }
            "bench" => {
                self.camera.snap();
                self.bench = Some(Bench {
                    frames_left: rest.parse().unwrap_or(600),
                    samples: Vec::new(),
                    origin: self.camera.center,
                    clock: 0.0,
                    last_stats: DrawStats::default(),
                });
                self.last_frame = Instant::now();
            }
            "settled" => self.report_when_settled = true,
            "stats" => self.print_stats(),
            other => self.command(event_loop, other),
        }
    }

    fn print_stats(&self) {
        let current = self.matches.get(self.current).map(|&i| &self.nodes[i]);
        println!(
            "stats adapter={:?} center={:.1},{:.1} scale={:.4} matches={} current={} current_center={} drawn_triangles={} draws={} loading={}",
            self.renderer.adapter_name,
            self.camera.center[0],
            self.camera.center[1],
            self.camera.scale,
            self.matches.len(),
            current.map_or("-", |n| n.name.as_str()),
            current.map_or_else(
                || "-".to_string(),
                |n| format!(
                    "{:.1},{:.1}",
                    (n.min[0] + n.max[0]) * 0.5,
                    (n.min[1] + n.max[1]) * 0.5
                )
            ),
            self.last_stats.triangles,
            self.last_stats.draws,
            self.loading,
        );
    }

    // ---- search -----------------------------------------------------------

    /// Exact names first, then prefixes, then anything containing the query.
    fn refresh_matches(&mut self, jump: bool) {
        let needle = self.query.to_lowercase();
        self.matches.clear();
        self.current = 0;
        if needle.is_empty() {
            return;
        }
        let rank = |lower: &str| {
            if lower == needle {
                0
            } else if lower.starts_with(&needle) {
                1
            } else {
                2
            }
        };
        self.matches = (0..self.nodes.len())
            .filter(|&i| self.nodes[i].lower.contains(&needle))
            .collect();
        let nodes = &self.nodes;
        self.matches.sort_by(|&a, &b| {
            (rank(&nodes[a].lower), &nodes[a].lower).cmp(&(rank(&nodes[b].lower), &nodes[b].lower))
        });
        if jump {
            self.jump_to_current();
        }
    }

    fn step_match(&mut self, by: isize) {
        if self.matches.is_empty() {
            return;
        }
        let len = self.matches.len() as isize;
        self.current = (self.current as isize + by).rem_euclid(len) as usize;
        self.jump_to_current();
    }

    fn jump_to_current(&mut self) {
        let Some(&index) = self.matches.get(self.current) else {
            return;
        };
        let node = &self.nodes[index];
        let (min, max) = (node.min, node.max);
        let (width, height) = self.renderer.size();
        let usable = (f64::from(height) - f64::from(STATUS_PX)) * 0.8;
        let node_w = f64::from(max[0] - min[0]).max(1.0);
        let node_h = f64::from(max[1] - min[1]).max(1.0);
        let fit = (f64::from(width) * 0.8 / node_w).min(usable / node_h);
        let scale = self.clamp_scale(fit.clamp(TEXT_MIN_SCALE * 1.5, 1.0));
        self.camera.target_scale = scale;
        // A table taller than the window lands with its header in view.
        let center_y = if node_h * scale > usable {
            f64::from(min[1]) + usable * 0.5 / scale
        } else {
            f64::from(min[1] + max[1]) * 0.5
        };
        self.camera.target_center = [f64::from(min[0] + max[0]) * 0.5, center_y];
    }

    // ---- camera -----------------------------------------------------------

    fn fit_scale(&self) -> f64 {
        let Some((min, max)) = self.bounds else {
            return 1.0;
        };
        let (width, height) = self.renderer.size();
        let usable_h = f64::from(height) - f64::from(STATUS_PX) * self.dpi();
        let scale_x = f64::from(width) * 0.97 / f64::from((max[0] - min[0]).max(1.0));
        let scale_y = usable_h * 0.97 / f64::from((max[1] - min[1]).max(1.0));
        scale_x.min(scale_y)
    }

    fn fit(&mut self) {
        let Some((min, max)) = self.bounds else {
            return;
        };
        let scale = self.fit_scale();
        self.camera.target_scale = scale;
        let bar = f64::from(STATUS_PX) * self.dpi() * 0.5 / scale;
        self.camera.target_center = [
            f64::from(min[0] + max[0]) * 0.5,
            f64::from(min[1] + max[1]) * 0.5 + bar,
        ];
    }

    fn clamp_scale(&self, scale: f64) -> f64 {
        scale.clamp(self.fit_scale() * 0.5, MAX_SCALE)
    }

    fn dpi(&self) -> f64 {
        self.window.scale_factor()
    }

    fn integrate_held_keys(&mut self, dt: f64) -> bool {
        let down = |codes: &[KeyCode]| codes.iter().any(|c| self.held.contains(c));
        let x = f64::from(i8::from(down(&[KeyCode::KeyL, KeyCode::ArrowRight])))
            - f64::from(i8::from(down(&[KeyCode::KeyH, KeyCode::ArrowLeft])));
        let y = f64::from(i8::from(down(&[KeyCode::KeyJ, KeyCode::ArrowDown])))
            - f64::from(i8::from(down(&[KeyCode::KeyK, KeyCode::ArrowUp])));
        let zoom = f64::from(i8::from(down(&[
            KeyCode::Equal,
            KeyCode::KeyI,
            KeyCode::NumpadAdd,
        ]))) - f64::from(i8::from(down(&[
            KeyCode::Minus,
            KeyCode::KeyO,
            KeyCode::NumpadSubtract,
        ])));

        let boost = if self.modifiers.shift_key() { 3.0 } else { 1.0 };
        let step = PAN_PX_PER_SEC * self.dpi() * boost * dt / self.camera.target_scale;
        self.camera.target_center[0] += x * step;
        self.camera.target_center[1] += y * step;
        if zoom != 0.0 {
            let scale = self.camera.target_scale * (zoom * ZOOM_PER_SEC * dt).exp();
            self.camera.target_scale = self.clamp_scale(scale);
        }
        x != 0.0 || y != 0.0 || zoom != 0.0
    }

    fn clamp_center(&mut self) {
        if let Some((min, max)) = self.bounds {
            self.camera.target_center[0] =
                self.camera.target_center[0].clamp(f64::from(min[0]), f64::from(max[0]));
            self.camera.target_center[1] =
                self.camera.target_center[1].clamp(f64::from(min[1]), f64::from(max[1]));
        }
    }

    // ---- frame ------------------------------------------------------------

    fn redraw(&mut self) {
        let now = Instant::now();
        let dt = now.duration_since(self.last_frame).as_secs_f64().min(0.05);
        // A long gap means the viewer was idle, not that a frame was slow.
        let gap_ms = now.duration_since(self.last_frame).as_secs_f64() * 1000.0;
        if gap_ms < 100.0 {
            self.frame_ms += (gap_ms - self.frame_ms) * 0.08;
        }
        self.last_frame = now;

        let held = self.integrate_held_keys(dt);
        self.clamp_center();
        let moving = self.camera.step(dt);
        let benching = self.bench_step();

        self.refresh_status();
        let (quads, placements) = self.build_ui();

        let (width, height) = self.renderer.size();
        let bottom = height as i32;
        let areas: Vec<TextArea> = placements
            .iter()
            .map(|p| TextArea {
                buffer: match p.slot {
                    Slot::Label(i) => &self.labels[i].buffer,
                    Slot::StatusLeft => &self.status_left,
                    Slot::StatusRight => &self.status_right,
                    Slot::Help => &self.help,
                },
                left: p.x,
                top: p.y,
                scale: 1.0,
                bounds: TextBounds {
                    left: 0,
                    top: 0,
                    right: width as i32,
                    bottom,
                },
                default_color: p.color,
                custom_glyphs: &[],
            })
            .collect();

        let state = ViewState {
            center: [self.camera.center[0] as f32, self.camera.center[1] as f32],
            scale: self.camera.scale as f32,
            show_text: self.camera.scale >= TEXT_MIN_SCALE,
        };
        let presented = self.renderer.render(&state, &quads, areas);
        if let Some(stats) = presented {
            self.last_stats = stats;
            if let Some(bench) = &mut self.bench {
                bench.last_stats = stats;
            }
        }

        if self.report_when_settled && !moving && !held && !self.loading && !benching {
            self.report_when_settled = false;
            println!("settled");
        }
        if moving || held || benching || self.report_when_settled || presented.is_none() {
            self.window.request_redraw();
        }
    }

    /// Scripted panning for measuring frame times. Returns true while running.
    fn bench_step(&mut self) -> bool {
        let (width, _) = self.renderer.size();
        let Some(bench) = &mut self.bench else {
            return false;
        };
        bench.frames_left = bench.frames_left.saturating_sub(1);
        bench.clock += 1.0;
        let sway = (bench.clock * 0.02).sin() * f64::from(width) * 0.25 / self.camera.scale;
        self.camera.center = [bench.origin[0] + sway, bench.origin[1]];
        self.camera.target_center = self.camera.center;
        bench.samples.push(now_ms());
        if bench.frames_left > 0 {
            return true;
        }

        let mut deltas: Vec<f64> = bench.samples.windows(2).map(|w| w[1] - w[0]).collect();
        deltas.sort_by(f64::total_cmp);
        if !deltas.is_empty() {
            let mean = deltas.iter().sum::<f64>() / deltas.len() as f64;
            let at = |q: f64| deltas[((deltas.len() - 1) as f64 * q) as usize];
            println!(
                "bench scale={:.3} frames={} mean_ms={:.2} p50_ms={:.2} p99_ms={:.2} max_ms={:.2} fps={:.0} triangles={} draws={}",
                self.camera.scale,
                deltas.len(),
                mean,
                at(0.5),
                at(0.99),
                deltas[deltas.len() - 1],
                1000.0 / mean,
                bench.last_stats.triangles,
                bench.last_stats.draws,
            );
        }
        self.bench = None;
        false
    }

    fn refresh_status(&mut self) {
        let zoom = self.camera.scale * 100.0;
        let left = if self.searching {
            let count = if self.query.is_empty() {
                String::new()
            } else if self.matches.is_empty() {
                "  no match".to_string()
            } else {
                let name = &self.nodes[self.matches[self.current]].name;
                format!("  {}/{}  {name}", self.current + 1, self.matches.len())
            };
            format!("/{}_{count}", self.query)
        } else if let Some(error) = &self.error {
            format!("{}  {error}", self.file_name)
        } else {
            let matches = if self.matches.is_empty() {
                String::new()
            } else {
                let name = &self.nodes[self.matches[self.current]].name;
                format!(
                    "  |  /{}  {}/{}  {name}",
                    self.query,
                    self.current + 1,
                    self.matches.len()
                )
            };
            format!(
                "{}  |  {zoom:.1}%  |  {} nodes{matches}",
                self.file_name,
                self.nodes.len()
            )
        };
        let triangles = self.renderer.total_triangles();
        let right = format!(
            "{}  |  {:.1}M tris  |  {:.1} ms  |  ? help",
            self.info,
            triangles as f64 / 1e6,
            self.frame_ms
        );

        let fonts = &mut self.renderer.font_system;
        if left != self.status_left_text {
            set_text(fonts, &mut self.status_left, &left);
            self.status_left_text = left;
        }
        if right != self.status_right_text {
            set_text(fonts, &mut self.status_right, &right);
            self.status_right_text = right;
        }
    }

    fn build_ui(&self) -> (Vec<UiVertex>, Vec<Placement>) {
        let (width, height) = self.renderer.size();
        let (width, height) = (width as f32, height as f32);
        let dpi = self.dpi() as f32;
        let bar_h = STATUS_PX * dpi;
        let scale = self.camera.scale as f32;
        let center = [self.camera.center[0] as f32, self.camera.center[1] as f32];
        let to_screen = |p: [f32; 2]| {
            [
                (p[0] - center[0]) * scale + width * 0.5,
                (p[1] - center[1]) * scale + height * 0.5,
            ]
        };

        let mut quads = Vec::with_capacity(1024);
        let mut places = Vec::with_capacity(64);

        // Match outlines
        for (position, &index) in self.matches.iter().enumerate() {
            let node = &self.nodes[index];
            let a = to_screen(node.min);
            let b = to_screen(node.max);
            let is_current = position == self.current;
            let thick = if is_current { 3.0 * dpi } else { 2.0 * dpi };
            let color = if is_current { AMBER } else { BLUE };
            let pad = thick + 2.0 * dpi;
            outline(
                &mut quads,
                [a[0] - pad, a[1] - pad],
                [b[0] + pad, b[1] + pad],
                thick,
                color,
            );
        }

        // Labels stand in for the diagram's own text when zoomed out
        if self.camera.scale < LABEL_MAX_SCALE && !self.show_help {
            let current = self.matches.get(self.current).copied();
            let ordered = current
                .into_iter()
                .chain(self.matches.iter().copied().filter(|m| Some(*m) != current))
                .chain(self.label_order.iter().copied());
            let mut seen = HashSet::new();
            let mut taken: Vec<[f32; 4]> = Vec::new();
            let pad_x = 4.0 * dpi;
            let pad_y = 1.5 * dpi;
            for index in ordered {
                if !seen.insert(index) {
                    continue;
                }
                let node = &self.nodes[index];
                let label = &self.labels[index];
                let a = to_screen(node.min);
                let b = to_screen(node.max);
                let is_match = self.matches.contains(&index);
                if !is_match && b[0] - a[0] < 14.0 * dpi {
                    continue;
                }
                let x = (a[0] + b[0]) * 0.5 - label.width * 0.5;
                // A node with room for it gets the label over its header row.
                let y = if b[1] - a[1] > label.height * 4.0 {
                    a[1] + label.height * 0.5
                } else {
                    (a[1] + b[1]) * 0.5 - label.height * 0.5
                };
                let rect = [
                    x - pad_x,
                    y - pad_y,
                    x + label.width + pad_x,
                    y + label.height + pad_y,
                ];
                let clipped =
                    rect[0] < 0.0 || rect[2] > width || rect[1] < 0.0 || rect[3] > height - bar_h;
                if clipped && !is_match {
                    continue;
                }
                if rect[2] < 0.0 || rect[0] > width || rect[3] < 0.0 || rect[1] > height - bar_h {
                    continue;
                }
                let overlaps = taken
                    .iter()
                    .any(|t| rect[0] < t[2] && rect[2] > t[0] && rect[1] < t[3] && rect[3] > t[1]);
                if overlaps {
                    continue;
                }
                taken.push(rect);
                quad(&mut quads, [rect[0], rect[1]], [rect[2], rect[3]], PILL);
                let color = if Some(index) == current {
                    AMBER_TEXT
                } else if is_match {
                    Color::rgb(0x6c, 0x9e, 0xdd)
                } else {
                    FG
                };
                places.push(Placement {
                    slot: Slot::Label(index),
                    x,
                    y,
                    color,
                });
            }
        }

        // Status line
        quad(&mut quads, [0.0, height - bar_h], [width, height], BAR);
        let (_, left_h) = measure(&self.status_left);
        let text_y = height - bar_h + (bar_h - left_h) * 0.5;
        let left_color = if self.error.is_some() && !self.searching {
            RED_TEXT
        } else {
            FG
        };
        places.push(Placement {
            slot: Slot::StatusLeft,
            x: 8.0 * dpi,
            y: text_y,
            color: left_color,
        });
        let (right_w, _) = measure(&self.status_right);
        places.push(Placement {
            slot: Slot::StatusRight,
            x: width - right_w - 8.0 * dpi,
            y: text_y,
            color: MUTED,
        });

        if self.show_help {
            let (help_w, help_h) = measure(&self.help);
            let pad = 16.0 * dpi;
            let x = (width - help_w) * 0.5;
            let y = (height - help_h) * 0.5;
            quad(
                &mut quads,
                [x - pad, y - pad],
                [x + help_w + pad, y + help_h + pad],
                PANEL,
            );
            outline(
                &mut quads,
                [x - pad, y - pad],
                [x + help_w + pad, y + help_h + pad],
                1.0 * dpi,
                [0x7a, 0x70, 0x8a, 0xff],
            );
            places.push(Placement {
                slot: Slot::Help,
                x,
                y,
                color: FG,
            });
        }
        (quads, places)
    }
}

fn load(input: &std::path::Path, force: bool, status: &dyn Fn(String)) -> Result<Loaded, String> {
    let svg_start = Instant::now();
    let svg = source::svg_for(input, force, status)?;
    let svg_seconds = svg_start.elapsed().as_secs_f64();

    status("meshing...".to_string());
    let build_start = Instant::now();
    let text = std::fs::read_to_string(&svg.path)
        .map_err(|e| format!("read {}: {e}", svg.path.display()))?;
    let scene = scene::build(&text)?;
    Ok(Loaded {
        scene,
        svg_seconds,
        rendered: svg.rendered,
        build_seconds: build_start.elapsed().as_secs_f64(),
    })
}

fn modified(path: &std::path::Path) -> Option<SystemTime> {
    std::fs::metadata(path).and_then(|m| m.modified()).ok()
}

fn is_motion_key(code: KeyCode) -> bool {
    matches!(
        code,
        KeyCode::KeyH
            | KeyCode::KeyJ
            | KeyCode::KeyK
            | KeyCode::KeyL
            | KeyCode::ArrowLeft
            | KeyCode::ArrowRight
            | KeyCode::ArrowUp
            | KeyCode::ArrowDown
            | KeyCode::Equal
            | KeyCode::Minus
            | KeyCode::KeyI
            | KeyCode::KeyO
            | KeyCode::NumpadAdd
            | KeyCode::NumpadSubtract
    )
}

fn now_ms() -> f64 {
    static START: std::sync::OnceLock<Instant> = std::sync::OnceLock::new();
    START.get_or_init(Instant::now).elapsed().as_secs_f64() * 1000.0
}

fn make_buffer(fonts: &mut glyphon::FontSystem, text: &str, size: f32) -> Buffer {
    let mut buffer = Buffer::new(fonts, Metrics::new(size, (size * 1.3).ceil()));
    buffer.set_size(None, None);
    set_text(fonts, &mut buffer, text);
    buffer
}

fn set_text(fonts: &mut glyphon::FontSystem, buffer: &mut Buffer, text: &str) {
    let attrs = Attrs::new().family(Family::Name(UI_FONT));
    buffer.set_text(text, &attrs, Shaping::Advanced, None);
    buffer.shape_until_scroll(fonts, false);
}

fn measure(buffer: &Buffer) -> (f32, f32) {
    let mut width = 0.0f32;
    let mut lines = 0;
    for run in buffer.layout_runs() {
        width = width.max(run.line_w);
        lines += 1;
    }
    (width, lines.max(1) as f32 * buffer.metrics().line_height)
}

fn quad(out: &mut Vec<UiVertex>, min: [f32; 2], max: [f32; 2], color: [u8; 4]) {
    let v = |x, y| UiVertex { pos: [x, y], color };
    out.extend_from_slice(&[
        v(min[0], min[1]),
        v(max[0], min[1]),
        v(max[0], max[1]),
        v(min[0], min[1]),
        v(max[0], max[1]),
        v(min[0], max[1]),
    ]);
}

fn outline(out: &mut Vec<UiVertex>, min: [f32; 2], max: [f32; 2], thick: f32, color: [u8; 4]) {
    quad(out, min, [max[0], min[1] + thick], color);
    quad(out, [min[0], max[1] - thick], max, color);
    quad(
        out,
        [min[0], min[1] + thick],
        [min[0] + thick, max[1] - thick],
        color,
    );
    quad(
        out,
        [max[0] - thick, min[1] + thick],
        [max[0], max[1] - thick],
        color,
    );
}
