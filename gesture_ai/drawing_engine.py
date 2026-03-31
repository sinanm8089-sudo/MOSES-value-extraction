"""Air drawing engine and toolbar UI."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

import cv2
import numpy as np

from gesture_ai.config import DrawingConfig, UiConfig
from gesture_ai.shape_recognition import ShapeRecognizer


@dataclass
class DrawingState:
    active_color_name: str = "blue"
    brush_size: int = 8
    last_point: Optional[Tuple[int, int]] = None
    current_stroke: List[Tuple[int, int]] = field(default_factory=list)


class DrawingEngine:
    def __init__(
        self,
        frame_shape: Tuple[int, int, int],
        drawing_config: DrawingConfig,
        ui_config: UiConfig,
        output_dir: Path,
    ) -> None:
        h, w, _ = frame_shape
        self.canvas = np.zeros((h, w, 3), dtype=np.uint8)
        self.overlay = np.zeros((h, w, 3), dtype=np.uint8)
        self.state = DrawingState(brush_size=drawing_config.default_brush_size)
        self.drawing_config = drawing_config
        self.ui_config = ui_config
        self.shape_recognizer = ShapeRecognizer()
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.palette_regions = self._init_palette_regions()

    def _init_palette_regions(self):
        x = 20
        regions = {}
        for name in self.ui_config.colors:
            regions[name] = (x, 18, x + 76, 60)
            x += 90
        return regions

    def handle_selection(self, cursor: Tuple[int, int]) -> None:
        x, y = cursor
        for name, (x1, y1, x2, y2) in self.palette_regions.items():
            if x1 <= x <= x2 and y1 <= y <= y2:
                self.state.active_color_name = name

    def draw(self, point: Tuple[int, int], selecting: bool = False) -> None:
        if selecting:
            self.state.last_point = None
            if self.state.current_stroke:
                self.state.current_stroke.clear()
            return

        color = self.ui_config.colors[self.state.active_color_name]
        thickness = (
            self.drawing_config.eraser_size
            if self.state.active_color_name == "eraser"
            else self.state.brush_size
        )

        if self.state.last_point is None:
            self.state.last_point = point
        cv2.line(self.canvas, self.state.last_point, point, color, thickness)
        self.state.last_point = point
        self.state.current_stroke.append(point)

    def lift_pen(self) -> None:
        if len(self.state.current_stroke) >= self.drawing_config.shape_stabilization_points:
            shape_match = self.shape_recognizer.recognize(self.state.current_stroke)
            if shape_match:
                color = self.ui_config.colors[self.state.active_color_name]
                self.shape_recognizer.draw_refined_shape(
                    self.canvas,
                    shape_match,
                    color=color,
                    thickness=self.state.brush_size,
                    fill=self.drawing_config.auto_fill_shapes,
                )
        self.state.current_stroke.clear()
        self.state.last_point = None

    def clear(self) -> None:
        self.canvas[:] = 0

    def save(self) -> Path:
        ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        path = self.output_dir / f"air_drawing_{ts}.png"
        cv2.imwrite(str(path), self.canvas)
        return path

    def compose(self, frame):
        cv2.addWeighted(frame, 1.0, self.canvas, self.drawing_config.canvas_alpha, 0, frame)
        self._draw_toolbar(frame)

    def _draw_toolbar(self, frame) -> None:
        cv2.rectangle(frame, (0, 0), (frame.shape[1], self.ui_config.toolbar_height), (20, 24, 35), -1)
        for name, (x1, y1, x2, y2) in self.palette_regions.items():
            color = self.ui_config.colors[name]
            thickness = 3 if self.state.active_color_name == name else 1
            cv2.rectangle(frame, (x1, y1), (x2, y2), color, -1)
            cv2.rectangle(frame, (x1, y1), (x2, y2), (255, 255, 255), thickness)
            cv2.putText(frame, name[:3].upper(), (x1 + 8, y2 + 14), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (230, 230, 230), 1)
