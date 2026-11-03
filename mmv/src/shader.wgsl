struct View {
    center: vec2<f32>,
    half_size: vec2<f32>,
    scale: f32,
    min_stroke_px: f32,
    pad: vec2<f32>,
};

@group(0) @binding(0) var<uniform> view: View;

struct Out {
    @builtin(position) position: vec4<f32>,
    @location(0) color: vec4<f32>,
};

fn to_clip(world: vec2<f32>, depth: f32) -> vec4<f32> {
    let ndc = (world - view.center) * view.scale / view.half_size;
    return vec4<f32>(ndc.x, -ndc.y, depth, 1.0);
}

// Strokes keep a minimum width on screen, so edges and borders stay visible
// when the whole diagram is in view.
@vertex
fn vs_shape(
    @location(0) pos: vec2<f32>,
    @location(1) normal: vec2<f32>,
    @location(2) half_width: f32,
    @location(3) depth: f32,
    @location(4) color: vec4<f32>,
) -> Out {
    var width = 0.0;
    if half_width > 0.0 {
        width = max(half_width, 0.5 * view.min_stroke_px / view.scale);
    }
    var out: Out;
    out.position = to_clip(pos + normal * width, depth);
    out.color = color;
    return out;
}

@vertex
fn vs_text(
    @location(0) pos: vec2<f32>,
    @location(1) depth: f32,
    @location(2) color: vec4<f32>,
) -> Out {
    var out: Out;
    out.position = to_clip(pos, depth);
    out.color = color;
    return out;
}

@vertex
fn vs_ui(@location(0) pos: vec2<f32>, @location(1) color: vec4<f32>) -> Out {
    let ndc = pos / view.half_size - vec2<f32>(1.0, 1.0);
    var out: Out;
    out.position = vec4<f32>(ndc.x, -ndc.y, 0.0, 1.0);
    out.color = color;
    return out;
}

@fragment
fn fs_main(in: Out) -> @location(0) vec4<f32> {
    return in.color;
}
