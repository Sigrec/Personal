//! wgpu renderer on Direct3D 12: the diagram in a 4x MSAA pass, then the status
//! line, highlights and labels in a second pass over the resolved image.

use crate::scene::{BACKGROUND, Scene, ShapeVertex, TextVertex, Tile};
use bytemuck::{Pod, Zeroable};
use glyphon::{
    Cache, ColorMode, FontSystem, Resolution, SwashCache, TextArea, TextAtlas, TextRenderer,
};
use std::sync::Arc;
use wgpu::util::{BufferInitDescriptor, DeviceExt};
use winit::event_loop::ActiveEventLoop;
use winit::window::Window;

const MSAA: u32 = 4;
const COLOR_FORMAT: wgpu::TextureFormat = wgpu::TextureFormat::Bgra8Unorm;
const DEPTH_FORMAT: wgpu::TextureFormat = wgpu::TextureFormat::Depth32Float;
const MIN_STROKE_PX: f32 = 1.0;

#[repr(C)]
#[derive(Clone, Copy, Pod, Zeroable)]
struct ViewUniform {
    center: [f32; 2],
    half_size: [f32; 2],
    scale: f32,
    min_stroke_px: f32,
    pad: [f32; 2],
}

#[repr(C)]
#[derive(Clone, Copy, Pod, Zeroable)]
pub struct UiVertex {
    pub pos: [f32; 2],
    pub color: [u8; 4],
}

pub struct ViewState {
    pub center: [f32; 2],
    pub scale: f32,
    pub show_text: bool,
}

#[derive(Clone, Copy, Default)]
pub struct DrawStats {
    pub draws: u32,
    pub triangles: u64,
}

struct GpuScene {
    shape_vb: wgpu::Buffer,
    shape_ib: wgpu::Buffer,
    shape_tiles: Vec<Tile>,
    text_vb: wgpu::Buffer,
    text_ib: wgpu::Buffer,
    text_tiles: Vec<Tile>,
    triangles: u64,
}

pub struct Renderer {
    instance: wgpu::Instance,
    window: Arc<Window>,
    device: wgpu::Device,
    queue: wgpu::Queue,
    surface: wgpu::Surface<'static>,
    config: wgpu::SurfaceConfiguration,
    msaa_view: wgpu::TextureView,
    depth_view: wgpu::TextureView,
    view_buffer: wgpu::Buffer,
    view_group: wgpu::BindGroup,
    shape_pipeline: wgpu::RenderPipeline,
    text_pipeline: wgpu::RenderPipeline,
    ui_pipeline: wgpu::RenderPipeline,
    ui_buffer: wgpu::Buffer,
    scene: Option<GpuScene>,
    pub font_system: FontSystem,
    swash_cache: SwashCache,
    viewport: glyphon::Viewport,
    atlas: TextAtlas,
    text_renderer: TextRenderer,
    pub adapter_name: String,
}

impl Renderer {
    pub fn new(window: Arc<Window>, event_loop: &ActiveEventLoop) -> Result<Self, String> {
        pollster::block_on(Self::create(window, event_loop))
    }

    async fn create(window: Arc<Window>, event_loop: &ActiveEventLoop) -> Result<Self, String> {
        let mut descriptor = wgpu::InstanceDescriptor::new_with_display_handle(Box::new(
            event_loop.owned_display_handle(),
        ));
        descriptor.backends = wgpu::Backends::DX12;
        let instance = wgpu::Instance::new(descriptor);
        let surface = instance
            .create_surface(window.clone())
            .map_err(|e| format!("create surface: {e}"))?;
        let adapter = instance
            .request_adapter(&wgpu::RequestAdapterOptions {
                power_preference: wgpu::PowerPreference::HighPerformance,
                force_fallback_adapter: false,
                compatible_surface: Some(&surface),
                apply_limit_buckets: false,
            })
            .await
            .map_err(|e| format!("no Direct3D 12 adapter: {e}"))?;
        let adapter_name = adapter.get_info().name;
        let (device, queue) = adapter
            .request_device(&wgpu::DeviceDescriptor {
                required_limits: adapter.limits(),
                ..Default::default()
            })
            .await
            .map_err(|e| format!("create device: {e}"))?;

        let size = window.inner_size();
        let present_mode = match std::env::var("MMV_PRESENT").as_deref() {
            Ok("immediate") => wgpu::PresentMode::Immediate,
            _ => wgpu::PresentMode::Fifo,
        };
        let config = wgpu::SurfaceConfiguration {
            usage: wgpu::TextureUsages::RENDER_ATTACHMENT,
            format: COLOR_FORMAT,
            width: size.width.max(1),
            height: size.height.max(1),
            present_mode,
            alpha_mode: wgpu::CompositeAlphaMode::Opaque,
            view_formats: vec![],
            desired_maximum_frame_latency: 2,
            color_space: wgpu::SurfaceColorSpace::Auto,
        };
        surface.configure(&device, &config);

        let view_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("view"),
            size: size_of::<ViewUniform>() as u64,
            usage: wgpu::BufferUsages::UNIFORM | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });
        let layout = device.create_bind_group_layout(&wgpu::BindGroupLayoutDescriptor {
            label: Some("view"),
            entries: &[wgpu::BindGroupLayoutEntry {
                binding: 0,
                visibility: wgpu::ShaderStages::VERTEX,
                ty: wgpu::BindingType::Buffer {
                    ty: wgpu::BufferBindingType::Uniform,
                    has_dynamic_offset: false,
                    min_binding_size: None,
                },
                count: None,
            }],
        });
        let view_group = device.create_bind_group(&wgpu::BindGroupDescriptor {
            label: Some("view"),
            layout: &layout,
            entries: &[wgpu::BindGroupEntry {
                binding: 0,
                resource: view_buffer.as_entire_binding(),
            }],
        });
        let pipeline_layout = device.create_pipeline_layout(&wgpu::PipelineLayoutDescriptor {
            label: Some("mmv"),
            bind_group_layouts: &[Some(&layout)],
            immediate_size: 0,
        });
        let shader = device.create_shader_module(wgpu::include_wgsl!("shader.wgsl"));

        let shape_attrs = wgpu::vertex_attr_array![
            0 => Float32x2, 1 => Float32x2, 2 => Float32, 3 => Float32, 4 => Unorm8x4
        ];
        let text_attrs = wgpu::vertex_attr_array![0 => Float32x2, 1 => Float32, 2 => Unorm8x4];
        let ui_attrs = wgpu::vertex_attr_array![0 => Float32x2, 1 => Unorm8x4];

        let shape_pipeline = world_pipeline(
            &device,
            &pipeline_layout,
            &shader,
            "vs_shape",
            wgpu::VertexBufferLayout {
                array_stride: size_of::<ShapeVertex>() as u64,
                step_mode: wgpu::VertexStepMode::Vertex,
                attributes: &shape_attrs,
            },
        );
        let text_pipeline = world_pipeline(
            &device,
            &pipeline_layout,
            &shader,
            "vs_text",
            wgpu::VertexBufferLayout {
                array_stride: size_of::<TextVertex>() as u64,
                step_mode: wgpu::VertexStepMode::Vertex,
                attributes: &text_attrs,
            },
        );
        let ui_pipeline = device.create_render_pipeline(&wgpu::RenderPipelineDescriptor {
            label: Some("ui"),
            layout: Some(&pipeline_layout),
            vertex: wgpu::VertexState {
                module: &shader,
                entry_point: Some("vs_ui"),
                compilation_options: Default::default(),
                buffers: &[Some(wgpu::VertexBufferLayout {
                    array_stride: size_of::<UiVertex>() as u64,
                    step_mode: wgpu::VertexStepMode::Vertex,
                    attributes: &ui_attrs,
                })],
            },
            primitive: wgpu::PrimitiveState::default(),
            depth_stencil: None,
            multisample: wgpu::MultisampleState::default(),
            fragment: Some(wgpu::FragmentState {
                module: &shader,
                entry_point: Some("fs_main"),
                compilation_options: Default::default(),
                targets: &[Some(wgpu::ColorTargetState {
                    format: COLOR_FORMAT,
                    blend: Some(wgpu::BlendState::ALPHA_BLENDING),
                    write_mask: wgpu::ColorWrites::ALL,
                })],
            }),
            multiview_mask: None,
            cache: None,
        });
        let ui_buffer = ui_vertex_buffer(&device, 4096);
        let (msaa_view, depth_view) = attachments(&device, &config);

        let font_system = FontSystem::new();
        let swash_cache = SwashCache::new();
        let cache = Cache::new(&device);
        let viewport = glyphon::Viewport::new(&device, &cache);
        let mut atlas =
            TextAtlas::with_color_mode(&device, &queue, &cache, COLOR_FORMAT, ColorMode::Web);
        let text_renderer =
            TextRenderer::new(&mut atlas, &device, wgpu::MultisampleState::default(), None);

        Ok(Self {
            instance,
            window,
            device,
            queue,
            surface,
            config,
            msaa_view,
            depth_view,
            view_buffer,
            view_group,
            shape_pipeline,
            text_pipeline,
            ui_pipeline,
            ui_buffer,
            scene: None,
            font_system,
            swash_cache,
            viewport,
            atlas,
            text_renderer,
            adapter_name,
        })
    }

    pub fn size(&self) -> (u32, u32) {
        (self.config.width, self.config.height)
    }

    pub fn resize(&mut self, width: u32, height: u32) {
        if width == 0 || height == 0 {
            return;
        }
        self.config.width = width;
        self.config.height = height;
        self.surface.configure(&self.device, &self.config);
        (self.msaa_view, self.depth_view) = attachments(&self.device, &self.config);
    }

    /// Moves the scene's geometry to the GPU, replacing whatever was there.
    pub fn upload(&mut self, scene: &Scene) -> Result<(), String> {
        let limit = self.device.limits().max_buffer_size;
        let largest = (scene.text_vertices.len() * size_of::<TextVertex>())
            .max(scene.shape_vertices.len() * size_of::<ShapeVertex>())
            .max(scene.text_indices.len() * 4)
            .max(scene.shape_indices.len() * 4) as u64;
        if largest > limit {
            return Err(format!(
                "diagram needs a {largest} byte buffer, the GPU allows {limit}"
            ));
        }

        let usage_v = wgpu::BufferUsages::VERTEX;
        let usage_i = wgpu::BufferUsages::INDEX;
        self.scene = Some(GpuScene {
            shape_vb: init_buffer(&self.device, &scene.shape_vertices, usage_v),
            shape_ib: init_buffer(&self.device, &scene.shape_indices, usage_i),
            shape_tiles: scene.shape_tiles.clone(),
            text_vb: init_buffer(&self.device, &scene.text_vertices, usage_v),
            text_ib: init_buffer(&self.device, &scene.text_indices, usage_i),
            text_tiles: scene.text_tiles.clone(),
            triangles: (scene.shape_indices.len() + scene.text_indices.len()) as u64 / 3,
        });
        Ok(())
    }

    pub fn total_triangles(&self) -> u64 {
        self.scene.as_ref().map_or(0, |s| s.triangles)
    }

    /// Draws a frame. Returns None when no frame was presented, because the
    /// surface was rebuilt or is occluded, and the caller should try again.
    pub fn render<'a>(
        &mut self,
        state: &ViewState,
        ui: &[UiVertex],
        areas: impl IntoIterator<Item = TextArea<'a>>,
    ) -> Option<DrawStats> {
        let (width, height) = self.size();
        let half = [width as f32 * 0.5, height as f32 * 0.5];
        let uniform = ViewUniform {
            center: state.center,
            half_size: half,
            scale: state.scale,
            min_stroke_px: MIN_STROKE_PX,
            pad: [0.0, 0.0],
        };
        self.queue
            .write_buffer(&self.view_buffer, 0, bytemuck::bytes_of(&uniform));

        let needed = size_of_val(ui) as u64;
        if needed > self.ui_buffer.size() {
            self.ui_buffer = ui_vertex_buffer(&self.device, ui.len() * 2);
        }
        if !ui.is_empty() {
            self.queue
                .write_buffer(&self.ui_buffer, 0, bytemuck::cast_slice(ui));
        }

        self.viewport
            .update(&self.queue, Resolution { width, height });
        if let Err(e) = self.text_renderer.prepare(
            &self.device,
            &self.queue,
            &mut self.font_system,
            &mut self.atlas,
            &self.viewport,
            areas,
            &mut self.swash_cache,
        ) {
            eprintln!("text prepare failed: {e}");
        }

        let frame = match self.surface.get_current_texture() {
            wgpu::CurrentSurfaceTexture::Success(frame) => frame,
            wgpu::CurrentSurfaceTexture::Timeout | wgpu::CurrentSurfaceTexture::Occluded => {
                return None;
            }
            wgpu::CurrentSurfaceTexture::Outdated | wgpu::CurrentSurfaceTexture::Suboptimal(_) => {
                self.surface.configure(&self.device, &self.config);
                return None;
            }
            wgpu::CurrentSurfaceTexture::Lost => {
                self.surface = self.instance.create_surface(self.window.clone()).ok()?;
                self.surface.configure(&self.device, &self.config);
                return None;
            }
            wgpu::CurrentSurfaceTexture::Validation => {
                eprintln!("surface validation error");
                return None;
            }
        };
        let view = frame
            .texture
            .create_view(&wgpu::TextureViewDescriptor::default());
        let mut encoder = self
            .device
            .create_command_encoder(&wgpu::CommandEncoderDescriptor { label: None });

        let mut stats = DrawStats::default();
        {
            let clear = wgpu::Color {
                r: f64::from(BACKGROUND[0]) / 255.0,
                g: f64::from(BACKGROUND[1]) / 255.0,
                b: f64::from(BACKGROUND[2]) / 255.0,
                a: 1.0,
            };
            let mut pass = encoder.begin_render_pass(&wgpu::RenderPassDescriptor {
                label: Some("world"),
                color_attachments: &[Some(wgpu::RenderPassColorAttachment {
                    view: &self.msaa_view,
                    depth_slice: None,
                    resolve_target: Some(&view),
                    ops: wgpu::Operations {
                        load: wgpu::LoadOp::Clear(clear),
                        store: wgpu::StoreOp::Discard,
                    },
                })],
                depth_stencil_attachment: Some(wgpu::RenderPassDepthStencilAttachment {
                    view: &self.depth_view,
                    depth_ops: Some(wgpu::Operations {
                        load: wgpu::LoadOp::Clear(1.0),
                        store: wgpu::StoreOp::Discard,
                    }),
                    stencil_ops: None,
                }),
                timestamp_writes: None,
                occlusion_query_set: None,
                multiview_mask: None,
            });
            if let Some(scene) = &self.scene {
                let min = [
                    state.center[0] - half[0] / state.scale,
                    state.center[1] - half[1] / state.scale,
                ];
                let max = [
                    state.center[0] + half[0] / state.scale,
                    state.center[1] + half[1] / state.scale,
                ];
                let margin = MIN_STROKE_PX / state.scale;
                pass.set_bind_group(0, &self.view_group, &[]);

                pass.set_pipeline(&self.shape_pipeline);
                pass.set_vertex_buffer(0, scene.shape_vb.slice(..));
                pass.set_index_buffer(scene.shape_ib.slice(..), wgpu::IndexFormat::Uint32);
                draw_tiles(&mut pass, &scene.shape_tiles, min, max, margin, &mut stats);

                if state.show_text {
                    pass.set_pipeline(&self.text_pipeline);
                    pass.set_vertex_buffer(0, scene.text_vb.slice(..));
                    pass.set_index_buffer(scene.text_ib.slice(..), wgpu::IndexFormat::Uint32);
                    draw_tiles(&mut pass, &scene.text_tiles, min, max, 0.0, &mut stats);
                }
            }
        }
        {
            let mut pass = encoder.begin_render_pass(&wgpu::RenderPassDescriptor {
                label: Some("ui"),
                color_attachments: &[Some(wgpu::RenderPassColorAttachment {
                    view: &view,
                    depth_slice: None,
                    resolve_target: None,
                    ops: wgpu::Operations {
                        load: wgpu::LoadOp::Load,
                        store: wgpu::StoreOp::Store,
                    },
                })],
                depth_stencil_attachment: None,
                timestamp_writes: None,
                occlusion_query_set: None,
                multiview_mask: None,
            });
            if !ui.is_empty() {
                pass.set_pipeline(&self.ui_pipeline);
                pass.set_bind_group(0, &self.view_group, &[]);
                pass.set_vertex_buffer(0, self.ui_buffer.slice(..needed));
                pass.draw(0..ui.len() as u32, 0..1);
            }
            if let Err(e) = self
                .text_renderer
                .render(&self.atlas, &self.viewport, &mut pass)
            {
                eprintln!("text render failed: {e}");
            }
        }

        self.queue.submit(Some(encoder.finish()));
        self.window.pre_present_notify();
        self.queue.present(frame);
        self.atlas.trim();
        Some(stats)
    }
}

fn draw_tiles(
    pass: &mut wgpu::RenderPass<'_>,
    tiles: &[Tile],
    min: [f32; 2],
    max: [f32; 2],
    margin: f32,
    stats: &mut DrawStats,
) {
    for tile in tiles {
        let off_screen = tile.max[0] + margin < min[0]
            || tile.min[0] - margin > max[0]
            || tile.max[1] + margin < min[1]
            || tile.min[1] - margin > max[1];
        if off_screen {
            continue;
        }
        pass.draw_indexed(tile.start..tile.start + tile.count, 0, 0..1);
        stats.draws += 1;
        stats.triangles += u64::from(tile.count) / 3;
    }
}

fn world_pipeline(
    device: &wgpu::Device,
    layout: &wgpu::PipelineLayout,
    shader: &wgpu::ShaderModule,
    entry: &str,
    buffer: wgpu::VertexBufferLayout<'_>,
) -> wgpu::RenderPipeline {
    device.create_render_pipeline(&wgpu::RenderPipelineDescriptor {
        label: Some(entry),
        layout: Some(layout),
        vertex: wgpu::VertexState {
            module: shader,
            entry_point: Some(entry),
            compilation_options: Default::default(),
            buffers: &[Some(buffer)],
        },
        primitive: wgpu::PrimitiveState::default(),
        depth_stencil: Some(wgpu::DepthStencilState {
            format: DEPTH_FORMAT,
            depth_write_enabled: Some(true),
            depth_compare: Some(wgpu::CompareFunction::Less),
            stencil: wgpu::StencilState::default(),
            bias: wgpu::DepthBiasState::default(),
        }),
        multisample: wgpu::MultisampleState {
            count: MSAA,
            mask: !0,
            alpha_to_coverage_enabled: false,
        },
        fragment: Some(wgpu::FragmentState {
            module: shader,
            entry_point: Some("fs_main"),
            compilation_options: Default::default(),
            targets: &[Some(wgpu::ColorTargetState {
                format: COLOR_FORMAT,
                blend: None,
                write_mask: wgpu::ColorWrites::ALL,
            })],
        }),
        multiview_mask: None,
        cache: None,
    })
}

fn attachments(
    device: &wgpu::Device,
    config: &wgpu::SurfaceConfiguration,
) -> (wgpu::TextureView, wgpu::TextureView) {
    let make = |label, format| {
        device
            .create_texture(&wgpu::TextureDescriptor {
                label: Some(label),
                size: wgpu::Extent3d {
                    width: config.width,
                    height: config.height,
                    depth_or_array_layers: 1,
                },
                mip_level_count: 1,
                sample_count: MSAA,
                dimension: wgpu::TextureDimension::D2,
                format,
                usage: wgpu::TextureUsages::RENDER_ATTACHMENT,
                view_formats: &[],
            })
            .create_view(&wgpu::TextureViewDescriptor::default())
    };
    (make("msaa", COLOR_FORMAT), make("depth", DEPTH_FORMAT))
}

fn ui_vertex_buffer(device: &wgpu::Device, vertices: usize) -> wgpu::Buffer {
    device.create_buffer(&wgpu::BufferDescriptor {
        label: Some("ui"),
        size: (vertices * size_of::<UiVertex>()) as u64,
        usage: wgpu::BufferUsages::VERTEX | wgpu::BufferUsages::COPY_DST,
        mapped_at_creation: false,
    })
}

fn init_buffer<T: Pod>(
    device: &wgpu::Device,
    data: &[T],
    usage: wgpu::BufferUsages,
) -> wgpu::Buffer {
    let contents: &[u8] = if data.is_empty() {
        &[0; 16]
    } else {
        bytemuck::cast_slice(data)
    };
    device.create_buffer_init(&BufferInitDescriptor {
        label: None,
        contents,
        usage,
    })
}
