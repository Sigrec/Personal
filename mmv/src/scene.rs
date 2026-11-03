//! Parses an SVG with usvg, tessellates it with lyon and buckets the triangles
//! into grid tiles so the renderer can skip what is off screen.

use crate::theme::TextContrast;
use bytemuck::{Pod, Zeroable};
use lt::path::iterator::PathIterator;
use lt::path::{Path as LyonPath, PathEvent};
use lt::{
    BuffersBuilder, FillOptions, FillTessellator, FillVertex, StrokeOptions, StrokeTessellator,
    StrokeVertex, VertexBuffers,
};
use lyon_tessellation as lt;
use rayon::prelude::*;
use usvg::tiny_skia_path::PathSegment;

/// The window background, unless the diagram names a light Mermaid theme.
pub const BACKGROUND: [u8; 3] = [0x1d, 0x1a, 0x23];

const SHAPE_TOLERANCE: f32 = 0.08;
const TEXT_TOLERANCE: f32 = 0.03;
const GRID_X: usize = 64;
const GRID_Y: usize = 32;

#[repr(C)]
#[derive(Clone, Copy, Pod, Zeroable)]
pub struct ShapeVertex {
    pub pos: [f32; 2],
    pub normal: [f32; 2],
    pub half_width: f32,
    pub depth: f32,
    pub color: [u8; 4],
}

#[repr(C)]
#[derive(Clone, Copy, Pod, Zeroable)]
pub struct TextVertex {
    pub pos: [f32; 2],
    pub depth: f32,
    pub color: [u8; 4],
}

#[derive(Clone, Copy)]
pub struct Tile {
    pub min: [f32; 2],
    pub max: [f32; 2],
    pub start: u32,
    pub count: u32,
}

pub struct SearchNode {
    pub name: String,
    pub lower: String,
    pub min: [f32; 2],
    pub max: [f32; 2],
}

pub struct Scene {
    pub shape_vertices: Vec<ShapeVertex>,
    pub shape_indices: Vec<u32>,
    pub shape_tiles: Vec<Tile>,
    pub text_vertices: Vec<TextVertex>,
    pub text_indices: Vec<u32>,
    pub text_tiles: Vec<Tile>,
    pub nodes: Vec<SearchNode>,
    pub min: [f32; 2],
    pub max: [f32; 2],
    pub background: [u8; 3],
    pub skipped_paths: usize,
}

enum Job {
    Fill {
        path: LyonPath,
        rule: lt::FillRule,
        color: [u8; 4],
        is_text: bool,
    },
    Stroke {
        path: LyonPath,
        options: StrokeOptions,
        color: [u8; 4],
    },
}

enum Mesh {
    Shape(VertexBuffers<ShapeVertex, u32>),
    Text(VertexBuffers<TextVertex, u32>),
}

struct Placed {
    mesh: Mesh,
    min: [f32; 2],
    max: [f32; 2],
}

pub fn build(
    svg_text: &str,
    background: [u8; 3],
    text_contrast: Option<TextContrast>,
) -> Result<Scene, String> {
    let mut options = usvg::Options::default();
    options.fontdb_mut().load_system_fonts();
    let tree = usvg::Tree::from_str(svg_text, &options).map_err(|e| format!("parse svg: {e}"))?;

    let mut jobs = Vec::new();
    walk(tree.root(), 1.0, false, background, &mut jobs);
    if let Some(pair) = text_contrast {
        fix_text_contrast(&mut jobs, background, pair);
    }

    let total = jobs.len().max(1) as f32 + 2.0;
    let meshes: Vec<Option<Placed>> = jobs
        .par_iter()
        .enumerate()
        .map(|(i, job)| tessellate(job, 1.0 - (i as f32 + 1.0) / total))
        .collect();
    let skipped_paths = meshes.iter().filter(|m| m.is_none()).count();
    let placed: Vec<Placed> = meshes.into_iter().flatten().collect();
    if placed.is_empty() {
        return Err("the diagram has no drawable geometry".to_string());
    }

    let mut min = [f32::MAX; 2];
    let mut max = [f32::MIN; 2];
    for mesh in &placed {
        min = [min[0].min(mesh.min[0]), min[1].min(mesh.min[1])];
        max = [max[0].max(mesh.max[0]), max[1].max(mesh.max[1])];
    }

    let mut scene = Scene {
        shape_vertices: Vec::new(),
        shape_indices: Vec::new(),
        shape_tiles: Vec::new(),
        text_vertices: Vec::new(),
        text_indices: Vec::new(),
        text_tiles: Vec::new(),
        nodes: search_nodes(svg_text, &tree),
        min,
        max,
        background,
        skipped_paths,
    };
    bucket(&placed, &mut scene);
    Ok(scene)
}

/// Recolours text that would not read on the topmost fill painted behind it.
fn fix_text_contrast(jobs: &mut [Job], background: [u8; 3], pair: TextContrast) {
    let mut fills: Vec<([f32; 2], [f32; 2], [u8; 3])> = Vec::new();
    for job in jobs {
        let Job::Fill {
            path,
            color,
            is_text,
            ..
        } = job
        else {
            continue;
        };
        let Some((min, max)) = path_extent(path) else {
            continue;
        };
        let rgb = [color[0], color[1], color[2]];
        if !*is_text {
            fills.push((min, max, rgb));
            continue;
        }
        let center = [(min[0] + max[0]) * 0.5, (min[1] + max[1]) * 0.5];
        let backdrop = fills
            .iter()
            .rev()
            .find(|(lo, hi, _)| {
                center[0] >= lo[0] && center[0] <= hi[0] && center[1] >= lo[1] && center[1] <= hi[1]
            })
            .map_or(background, |fill| fill.2);
        let readable = pair.resolve(rgb, backdrop);
        *color = [readable[0], readable[1], readable[2], color[3]];
    }
}

fn path_extent(path: &LyonPath) -> Option<([f32; 2], [f32; 2])> {
    extent(path.iter().flat_map(|event| {
        let (a, b, c) = match event {
            PathEvent::Begin { at } => (Some(at), None, None),
            PathEvent::Line { to, .. } => (Some(to), None, None),
            PathEvent::Quadratic { ctrl, to, .. } => (Some(ctrl), Some(to), None),
            PathEvent::Cubic {
                ctrl1, ctrl2, to, ..
            } => (Some(ctrl1), Some(ctrl2), Some(to)),
            PathEvent::End { .. } => (None, None, None),
        };
        [a, b, c].into_iter().flatten().map(|p| p.to_array())
    }))
}

fn walk(group: &usvg::Group, opacity: f32, is_text: bool, bg: [u8; 3], jobs: &mut Vec<Job>) {
    for child in group.children() {
        match child {
            usvg::Node::Group(g) => walk(g, opacity * g.opacity().get(), is_text, bg, jobs),
            usvg::Node::Path(p) => add_path(p, opacity, is_text, bg, jobs),
            usvg::Node::Text(t) => walk(t.flattened(), opacity, true, bg, jobs),
            usvg::Node::Image(_) => {}
        }
    }
}

fn add_path(p: &usvg::Path, opacity: f32, is_text: bool, bg: [u8; 3], jobs: &mut Vec<Job>) {
    if !p.is_visible() {
        return;
    }
    let tf = p.abs_transform();
    let path = to_lyon(p, tf);

    let fill = p.fill().and_then(|fill| {
        let color = solid_color(fill.paint(), fill.opacity().get() * opacity, bg)?;
        let rule = match fill.rule() {
            usvg::FillRule::NonZero => lt::FillRule::NonZero,
            usvg::FillRule::EvenOdd => lt::FillRule::EvenOdd,
        };
        Some(Job::Fill {
            path: path.clone(),
            rule,
            color,
            is_text,
        })
    });

    let stroke = if is_text {
        None
    } else {
        p.stroke().and_then(|stroke| {
            let color = solid_color(stroke.paint(), stroke.opacity().get() * opacity, bg)?;
            let scale = (tf.sx * tf.sy - tf.kx * tf.ky).abs().sqrt();
            let dashed = stroke.dasharray().and_then(|pattern| {
                let scaled: Vec<f32> = pattern.iter().map(|d| d * scale).collect();
                dash(&path, &scaled, stroke.dashoffset() * scale)
            });
            Some(Job::Stroke {
                path: dashed.unwrap_or_else(|| path.clone()),
                options: stroke_options(stroke, scale),
                color,
            })
        })
    };

    let (first, second) = match p.paint_order() {
        usvg::PaintOrder::FillAndStroke => (fill, stroke),
        usvg::PaintOrder::StrokeAndFill => (stroke, fill),
    };
    jobs.extend(first);
    jobs.extend(second);
}

fn to_lyon(p: &usvg::Path, tf: usvg::Transform) -> LyonPath {
    let map = |pt: usvg::tiny_skia_path::Point| {
        lt::math::point(
            tf.sx * pt.x + tf.kx * pt.y + tf.tx,
            tf.ky * pt.x + tf.sy * pt.y + tf.ty,
        )
    };
    let mut builder = LyonPath::builder();
    let mut open = false;
    for segment in p.data().segments() {
        match segment {
            PathSegment::MoveTo(to) => {
                if open {
                    builder.end(false);
                }
                builder.begin(map(to));
                open = true;
            }
            PathSegment::LineTo(to) => {
                builder.line_to(map(to));
            }
            PathSegment::QuadTo(c, to) => {
                builder.quadratic_bezier_to(map(c), map(to));
            }
            PathSegment::CubicTo(c1, c2, to) => {
                builder.cubic_bezier_to(map(c1), map(c2), map(to));
            }
            PathSegment::Close => {
                builder.end(true);
                open = false;
            }
        }
    }
    if open {
        builder.end(false);
    }
    builder.build()
}

/// Depth testing replaces blending, so translucent paints are mixed against the
/// background here. Gradients take their first stop and patterns a flat grey.
fn solid_color(paint: &usvg::Paint, alpha: f32, bg: [u8; 3]) -> Option<[u8; 4]> {
    if alpha < 0.02 {
        return None;
    }
    let (r, g, b) = match paint {
        usvg::Paint::Color(c) => (c.red, c.green, c.blue),
        usvg::Paint::LinearGradient(g) => first_stop(g.stops())?,
        usvg::Paint::RadialGradient(g) => first_stop(g.stops())?,
        usvg::Paint::Pattern(_) => (0x91, 0x8a, 0x9f),
    };
    let mix =
        |c: u8, under: u8| (f32::from(c) * alpha + f32::from(under) * (1.0 - alpha)).round() as u8;
    Some([mix(r, bg[0]), mix(g, bg[1]), mix(b, bg[2]), 255])
}

fn first_stop(stops: &[usvg::Stop]) -> Option<(u8, u8, u8)> {
    let color = stops.first()?.color();
    Some((color.red, color.green, color.blue))
}

fn stroke_options(stroke: &usvg::Stroke, scale: f32) -> StrokeOptions {
    let cap = match stroke.linecap() {
        usvg::LineCap::Butt => lt::LineCap::Butt,
        usvg::LineCap::Round => lt::LineCap::Round,
        usvg::LineCap::Square => lt::LineCap::Square,
    };
    let join = match stroke.linejoin() {
        usvg::LineJoin::Miter => lt::LineJoin::Miter,
        usvg::LineJoin::MiterClip => lt::LineJoin::MiterClip,
        usvg::LineJoin::Round => lt::LineJoin::Round,
        usvg::LineJoin::Bevel => lt::LineJoin::Bevel,
    };
    StrokeOptions::tolerance(SHAPE_TOLERANCE)
        .with_line_width(stroke.width().get() * scale)
        .with_line_cap(cap)
        .with_line_join(join)
        .with_miter_limit(stroke.miterlimit().get().max(1.0))
}

/// Cuts a path into its dashes. Returns None for a pattern that draws solid.
fn dash(path: &LyonPath, pattern: &[f32], offset: f32) -> Option<LyonPath> {
    let period: f32 = pattern.iter().sum();
    if period <= 0.001 || pattern.iter().any(|d| *d < 0.0) {
        return None;
    }

    let mut builder = LyonPath::builder();
    let mut index = 0;
    let mut left = pattern[0];
    let mut skip = offset.rem_euclid(period);
    while skip >= left {
        skip -= left;
        index = (index + 1) % pattern.len();
        left = pattern[index];
    }
    left -= skip;

    let mut pen_down = false;
    for event in path.iter().flattened(SHAPE_TOLERANCE) {
        let (from, to, closing) = match event {
            PathEvent::Begin { at } => {
                if index % 2 == 0 {
                    builder.begin(at);
                    pen_down = true;
                }
                continue;
            }
            PathEvent::Line { from, to } => (from, to, false),
            PathEvent::End { last, first, close } => (last, first, close),
            _ => continue,
        };
        let ending = matches!(event, PathEvent::End { .. });
        if !ending || closing {
            let length = (to - from).length();
            let mut walked = 0.0;
            while length - walked > left {
                walked += left;
                let at = from.lerp(to, walked / length);
                if index % 2 == 0 {
                    if pen_down {
                        builder.line_to(at);
                        builder.end(false);
                        pen_down = false;
                    }
                } else {
                    builder.begin(at);
                    pen_down = true;
                }
                index = (index + 1) % pattern.len();
                left = pattern[index].max(0.001);
            }
            left -= length - walked;
            if pen_down {
                builder.line_to(to);
            }
        }
        if ending && pen_down {
            builder.end(false);
            pen_down = false;
        }
    }
    Some(builder.build())
}

fn tessellate(job: &Job, depth: f32) -> Option<Placed> {
    match job {
        Job::Fill {
            path,
            rule,
            color,
            is_text: true,
        } => {
            let options = FillOptions::tolerance(TEXT_TOLERANCE).with_fill_rule(*rule);
            let mut out: VertexBuffers<TextVertex, u32> = VertexBuffers::new();
            let ctor = |v: FillVertex| TextVertex {
                pos: v.position().to_array(),
                depth,
                color: *color,
            };
            FillTessellator::new()
                .tessellate_path(path, &options, &mut BuffersBuilder::new(&mut out, ctor))
                .ok()?;
            let (min, max) = extent(out.vertices.iter().map(|v| v.pos))?;
            Some(Placed {
                mesh: Mesh::Text(out),
                min,
                max,
            })
        }
        Job::Fill {
            path, rule, color, ..
        } => {
            let options = FillOptions::tolerance(SHAPE_TOLERANCE).with_fill_rule(*rule);
            let mut out: VertexBuffers<ShapeVertex, u32> = VertexBuffers::new();
            let ctor = |v: FillVertex| ShapeVertex {
                pos: v.position().to_array(),
                normal: [0.0, 0.0],
                half_width: 0.0,
                depth,
                color: *color,
            };
            FillTessellator::new()
                .tessellate_path(path, &options, &mut BuffersBuilder::new(&mut out, ctor))
                .ok()?;
            let (min, max) = extent(out.vertices.iter().map(|v| v.pos))?;
            Some(Placed {
                mesh: Mesh::Shape(out),
                min,
                max,
            })
        }
        Job::Stroke {
            path,
            options,
            color,
        } => {
            let half_width = options.line_width * 0.5;
            let mut out: VertexBuffers<ShapeVertex, u32> = VertexBuffers::new();
            let ctor = |v: StrokeVertex| ShapeVertex {
                pos: v.position_on_path().to_array(),
                normal: v.normal().to_array(),
                half_width,
                depth,
                color: *color,
            };
            StrokeTessellator::new()
                .tessellate_path(path, options, &mut BuffersBuilder::new(&mut out, ctor))
                .ok()?;
            let (min, max) = extent(out.vertices.iter().map(|v| v.pos))?;
            Some(Placed {
                mesh: Mesh::Shape(out),
                min,
                max,
            })
        }
    }
}

fn extent(points: impl Iterator<Item = [f32; 2]>) -> Option<([f32; 2], [f32; 2])> {
    let mut min = [f32::MAX; 2];
    let mut max = [f32::MIN; 2];
    let mut any = false;
    for p in points {
        min = [min[0].min(p[0]), min[1].min(p[1])];
        max = [max[0].max(p[0]), max[1].max(p[1])];
        any = true;
    }
    any.then_some((min, max))
}

/// Groups meshes by the tile holding their centre. A mesh wider than a tile
/// (a long edge, say) goes into one extra tile covering its full extent.
fn bucket(placed: &[Placed], scene: &mut Scene) {
    let cell_w = ((scene.max[0] - scene.min[0]) / GRID_X as f32).max(1.0);
    let cell_h = ((scene.max[1] - scene.min[1]) / GRID_Y as f32).max(1.0);
    let wide = GRID_X * GRID_Y;
    let mut cells: Vec<Vec<usize>> = vec![Vec::new(); wide + 1];

    for (i, mesh) in placed.iter().enumerate() {
        let w = mesh.max[0] - mesh.min[0];
        let h = mesh.max[1] - mesh.min[1];
        let cell = if w > cell_w || h > cell_h {
            wide
        } else {
            let cx = (mesh.min[0] + w * 0.5 - scene.min[0]) / cell_w;
            let cy = (mesh.min[1] + h * 0.5 - scene.min[1]) / cell_h;
            let x = (cx as usize).min(GRID_X - 1);
            let y = (cy as usize).min(GRID_Y - 1);
            y * GRID_X + x
        };
        cells[cell].push(i);
    }

    for ids in &cells {
        let mut shape = TileBuilder::new(scene.shape_indices.len());
        let mut text = TileBuilder::new(scene.text_indices.len());
        for &i in ids {
            let mesh = &placed[i];
            match &mesh.mesh {
                Mesh::Shape(out) => {
                    append(out, &mut scene.shape_vertices, &mut scene.shape_indices);
                    shape.grow(mesh);
                }
                Mesh::Text(out) => {
                    append(out, &mut scene.text_vertices, &mut scene.text_indices);
                    text.grow(mesh);
                }
            }
        }
        shape.finish(scene.shape_indices.len(), &mut scene.shape_tiles);
        text.finish(scene.text_indices.len(), &mut scene.text_tiles);
    }
}

fn append<V: Copy>(out: &VertexBuffers<V, u32>, vertices: &mut Vec<V>, indices: &mut Vec<u32>) {
    let base = vertices.len() as u32;
    vertices.extend_from_slice(&out.vertices);
    indices.extend(out.indices.iter().map(|i| i + base));
}

struct TileBuilder {
    start: usize,
    min: [f32; 2],
    max: [f32; 2],
}

impl TileBuilder {
    fn new(start: usize) -> Self {
        Self {
            start,
            min: [f32::MAX; 2],
            max: [f32::MIN; 2],
        }
    }

    fn grow(&mut self, mesh: &Placed) {
        self.min = [self.min[0].min(mesh.min[0]), self.min[1].min(mesh.min[1])];
        self.max = [self.max[0].max(mesh.max[0]), self.max[1].max(mesh.max[1])];
    }

    fn finish(self, end: usize, tiles: &mut Vec<Tile>) {
        if end > self.start {
            tiles.push(Tile {
                min: self.min,
                max: self.max,
                start: self.start as u32,
                count: (end - self.start) as u32,
            });
        }
    }
}

/// Finds the searchable nodes: every `<g class="node ...">` with an id, named
/// from the id for ER entities and from its label text for everything else.
fn search_nodes(svg_text: &str, tree: &usvg::Tree) -> Vec<SearchNode> {
    let Ok(doc) = usvg::roxmltree::Document::parse(svg_text) else {
        return Vec::new();
    };

    let mut nodes = Vec::new();
    for element in doc.descendants().filter(|n| n.tag_name().name() == "g") {
        let is_node = element
            .attribute("class")
            .is_some_and(|c| c.split_ascii_whitespace().any(|t| t == "node"));
        let Some(id) = element.attribute("id").filter(|_| is_node) else {
            continue;
        };
        let name = entity_name(id)
            .map(str::to_string)
            .or_else(|| label_text(element));
        let (Some(name), Some(found)) = (name, tree.node_by_id(id)) else {
            continue;
        };
        let mut min = [f32::MAX; 2];
        let mut max = [f32::MIN; 2];
        grow_bounds(found, &mut min, &mut max);
        if min[0] <= max[0] {
            nodes.push(SearchNode {
                lower: name.to_lowercase(),
                name,
                min,
                max,
            });
        }
    }
    nodes
}

/// Unions the absolute boxes of the paths and text under a node. A group's own
/// box comes out misplaced for Mermaid's nested transforms, so it is not used.
fn grow_bounds(node: &usvg::Node, min: &mut [f32; 2], max: &mut [f32; 2]) {
    let bounds = match node {
        usvg::Node::Group(group) => {
            for child in group.children() {
                grow_bounds(child, min, max);
            }
            return;
        }
        usvg::Node::Path(path) => path.abs_bounding_box(),
        usvg::Node::Text(text) => text.abs_bounding_box(),
        usvg::Node::Image(_) => return,
    };
    *min = [min[0].min(bounds.left()), min[1].min(bounds.top())];
    *max = [max[0].max(bounds.right()), max[1].max(bounds.bottom())];
}

/// A label's words, with its wrapped lines joined by spaces. Mermaid puts each
/// word in its own tspan inside one tspan per line, keeping the spaces.
fn label_text(element: usvg::roxmltree::Node) -> Option<String> {
    let text_of = |node: usvg::roxmltree::Node| -> String {
        node.descendants()
            .filter(|n| n.is_text())
            .filter_map(|n| n.text())
            .collect()
    };
    let mut lines: Vec<String> = element
        .descendants()
        .filter(|n| {
            n.attribute("class")
                .is_some_and(|c| c.split_ascii_whitespace().any(|t| t == "text-outer-tspan"))
        })
        .map(text_of)
        .collect();
    if lines.is_empty() {
        lines.push(text_of(element));
    }
    let name = join_wrapped_lines(&lines);
    (!name.is_empty()).then_some(name)
}

/// Joins a label's lines with spaces, except where Mermaid cut a long word to
/// fit: a full-width line of one word followed by a line starting in lowercase.
fn join_wrapped_lines(lines: &[String]) -> String {
    let longest = lines
        .iter()
        .map(|l| l.trim().chars().count())
        .max()
        .unwrap_or(0);
    let mut name = String::new();
    let mut previous_cut = false;
    for line in lines {
        let line = line.split_whitespace().collect::<Vec<_>>().join(" ");
        if line.is_empty() {
            continue;
        }
        let starts_lowercase = line.chars().next().is_some_and(char::is_lowercase);
        if !name.is_empty() && !(previous_cut && starts_lowercase) {
            name.push(' ');
        }
        name.push_str(&line);
        previous_cut = !line.contains(' ') && line.chars().count() + 2 >= longest;
    }
    name
}

/// `my-svg-entity-app_user-0` is the entity `app_user`.
fn entity_name(id: &str) -> Option<&str> {
    let (_, rest) = id.split_once("-entity-")?;
    let (name, suffix) = rest.rsplit_once('-')?;
    suffix.chars().all(|c| c.is_ascii_digit()).then_some(name)
}

#[cfg(test)]
mod tests {
    use super::join_wrapped_lines;

    fn join(lines: &[&str]) -> String {
        let lines: Vec<String> = lines.iter().map(|l| (*l).to_string()).collect();
        join_wrapped_lines(&lines)
    }

    #[test]
    fn rejoins_a_word_mermaid_cut() {
        assert_eq!(
            join(&["DatabaseCleanu", "p Lambda", "(Scheduled)"]),
            "DatabaseCleanup Lambda (Scheduled)"
        );
        assert_eq!(
            join(&["CloudFront CDN", "(BrandAssetsDist", "ribution)"]),
            "CloudFront CDN (BrandAssetsDistribution)"
        );
    }

    #[test]
    fn keeps_spaces_between_whole_words() {
        assert_eq!(join(&["Cloud", "Map"]), "Cloud Map");
        assert_eq!(
            join(&[
                "AWS AppConfig",
                "(shared",
                "non-secret",
                "config,",
                "hot-reload via"
            ]),
            "AWS AppConfig (shared non-secret config, hot-reload via"
        );
        assert_eq!(join(&["Notification", "Lambda"]), "Notification Lambda");
    }
}
