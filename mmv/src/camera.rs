//! Smoothed 2D camera. The keys move a target and the view eases towards it.

const EASE_PER_SEC: f64 = 16.0;

pub struct Camera {
    pub center: [f64; 2],
    pub scale: f64,
    pub target_center: [f64; 2],
    pub target_scale: f64,
}

impl Camera {
    pub fn new() -> Self {
        Self {
            center: [0.0, 0.0],
            scale: 1.0,
            target_center: [0.0, 0.0],
            target_scale: 1.0,
        }
    }

    /// Advances the easing and returns true while the view is still moving.
    pub fn step(&mut self, dt: f64) -> bool {
        let t = 1.0 - (-EASE_PER_SEC * dt).exp();
        let log_scale = self.scale.ln() + (self.target_scale.ln() - self.scale.ln()) * t;
        self.scale = log_scale.exp();
        self.center[0] += (self.target_center[0] - self.center[0]) * t;
        self.center[1] += (self.target_center[1] - self.center[1]) * t;

        let px_x = (self.target_center[0] - self.center[0]).abs() * self.scale;
        let px_y = (self.target_center[1] - self.center[1]).abs() * self.scale;
        let zoom_gap = (self.target_scale / self.scale - 1.0).abs();
        if px_x < 0.05 && px_y < 0.05 && zoom_gap < 0.0005 {
            self.snap();
            return false;
        }
        true
    }

    pub fn snap(&mut self) {
        self.center = self.target_center;
        self.scale = self.target_scale;
    }

    /// Zooms the target while keeping the world point under `anchor` fixed.
    /// `anchor` is in pixels relative to the window centre.
    pub fn zoom_at(&mut self, anchor: [f64; 2], new_scale: f64) {
        let world_x = self.target_center[0] + anchor[0] / self.target_scale;
        let world_y = self.target_center[1] + anchor[1] / self.target_scale;
        self.target_scale = new_scale;
        self.target_center = [
            world_x - anchor[0] / new_scale,
            world_y - anchor[1] / new_scale,
        ];
    }
}
