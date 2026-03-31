"""Centralized configuration for the AI Gesture Control System."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Tuple


@dataclass(slots=True)
class CameraConfig:
    device_index: int = 0
    width: int = 1280
    height: int = 720
    target_fps: int = 30


@dataclass(slots=True)
class TrackingConfig:
    max_num_hands: int = 2
    detection_confidence: float = 0.65
    tracking_confidence: float = 0.6
    model_complexity: int = 1


@dataclass(slots=True)
class GestureConfig:
    move_smoothing_window: int = 6
    pinch_click_distance_px: float = 30.0
    click_cooldown_sec: float = 0.45
    scroll_speed_factor: float = 18.0
    volume_min_distance_px: float = 40.0
    volume_max_distance_px: float = 220.0
    interaction_area_margin: int = 120
    sensitivity: float = 1.15


@dataclass(slots=True)
class DrawingConfig:
    default_brush_size: int = 8
    eraser_size: int = 28
    canvas_alpha: float = 0.78
    shape_stabilization_points: int = 20
    auto_fill_shapes: bool = False


@dataclass(slots=True)
class UiConfig:
    toolbar_height: int = 80
    status_font_scale: float = 0.7
    colors: Dict[str, Tuple[int, int, int]] = field(
        default_factory=lambda: {
            "blue": (255, 145, 80),
            "green": (80, 220, 100),
            "pink": (230, 100, 220),
            "yellow": (80, 220, 230),
            "white": (240, 240, 240),
            "eraser": (0, 0, 0),
        }
    )


@dataclass(slots=True)
class AppConfig:
    camera: CameraConfig = field(default_factory=CameraConfig)
    tracking: TrackingConfig = field(default_factory=TrackingConfig)
    gesture: GestureConfig = field(default_factory=GestureConfig)
    drawing: DrawingConfig = field(default_factory=DrawingConfig)
    ui: UiConfig = field(default_factory=UiConfig)
    output_dir: Path = field(default_factory=lambda: Path("outputs"))
    log_file: Path = field(default_factory=lambda: Path("logs/gesture_ai.log"))


CONFIG = AppConfig()
