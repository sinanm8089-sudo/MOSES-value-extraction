"""Signal smoothing helpers."""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from typing import Deque, Iterable, Tuple

import numpy as np


@dataclass
class MovingAverageSmoother:
    """Moving-average point smoother for reducing cursor jitter."""

    window_size: int = 5

    def __post_init__(self) -> None:
        self._history_x: Deque[float] = deque(maxlen=self.window_size)
        self._history_y: Deque[float] = deque(maxlen=self.window_size)

    def update(self, point: Tuple[float, float]) -> Tuple[int, int]:
        x, y = point
        self._history_x.append(x)
        self._history_y.append(y)
        return int(np.mean(self._history_x)), int(np.mean(self._history_y))

    def reset(self) -> None:
        self._history_x.clear()
        self._history_y.clear()


@dataclass
class OptionalKalmanFilter:
    """Minimal 2D Kalman-like filter using exponential blend fallback.

    Intended as an optional smoother where full Kalman setup is unnecessary.
    """

    alpha: float = 0.35
    _state: Tuple[float, float] | None = None

    def update(self, point: Tuple[float, float]) -> Tuple[int, int]:
        if self._state is None:
            self._state = point
            return int(point[0]), int(point[1])
        x = self.alpha * point[0] + (1 - self.alpha) * self._state[0]
        y = self.alpha * point[1] + (1 - self.alpha) * self._state[1]
        self._state = (x, y)
        return int(x), int(y)

    def reset(self) -> None:
        self._state = None


def point_distance(a: Tuple[int, int], b: Tuple[int, int]) -> float:
    return float(np.linalg.norm(np.array(a) - np.array(b)))


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(value, high))


def normalize(value: float, from_range: Iterable[float], to_range: Iterable[float]) -> float:
    start, end = from_range
    t_start, t_end = to_range
    if end == start:
        return t_start
    ratio = (value - start) / (end - start)
    return t_start + ratio * (t_end - t_start)
