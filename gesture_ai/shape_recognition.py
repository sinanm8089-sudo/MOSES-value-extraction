"""OpenCV contour-based shape recognition utilities."""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Tuple

import cv2
import numpy as np


@dataclass
class ShapeMatch:
    shape_name: str
    contour: np.ndarray
    approx: np.ndarray


class ShapeRecognizer:
    """Detect geometric primitives from a path trace."""

    def recognize(self, points: List[Tuple[int, int]]) -> Optional[ShapeMatch]:
        if len(points) < 12:
            return None

        canvas = np.zeros((720, 1280), dtype=np.uint8)
        contour = np.array(points, dtype=np.int32).reshape(-1, 1, 2)
        cv2.polylines(canvas, [contour], isClosed=True, color=255, thickness=3)
        contours, _ = cv2.findContours(canvas, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        if not contours:
            return None

        cnt = max(contours, key=cv2.contourArea)
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.03 * peri, True)
        vertices = len(approx)

        shape = "circle"
        if vertices == 3:
            shape = "triangle"
        elif vertices == 4:
            x, y, w, h = cv2.boundingRect(approx)
            ratio = w / max(h, 1)
            shape = "square" if 0.9 <= ratio <= 1.1 else "rectangle"
        elif vertices >= 5:
            shape = "circle"

        return ShapeMatch(shape_name=shape, contour=cnt, approx=approx)

    def draw_refined_shape(
        self,
        frame,
        shape: ShapeMatch,
        color: Tuple[int, int, int],
        thickness: int,
        fill: bool = False,
    ) -> None:
        if shape.shape_name == "circle":
            (x, y), radius = cv2.minEnclosingCircle(shape.contour)
            cv2.circle(frame, (int(x), int(y)), int(radius), color, -1 if fill else thickness)
        elif shape.shape_name in {"square", "rectangle", "triangle"}:
            cv2.drawContours(frame, [shape.approx], -1, color, -1 if fill else thickness)
