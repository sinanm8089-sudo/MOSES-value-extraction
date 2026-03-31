"""System action controller for gesture-driven commands."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Tuple

import pyautogui

from gesture_ai.gesture_engine import GestureState


@dataclass
class ControllerConfig:
    enable_mouse: bool = True
    enable_scroll: bool = True


class SystemController:
    def __init__(self, logger: logging.Logger, config: ControllerConfig | None = None) -> None:
        self.logger = logger
        self.config = config or ControllerConfig()
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.0

    def apply(self, gesture: GestureState) -> None:
        if self.config.enable_mouse and gesture.mouse_pos:
            self._move_mouse(gesture.mouse_pos)
        if self.config.enable_mouse and gesture.click:
            pyautogui.click()
        if self.config.enable_scroll and gesture.scroll_delta:
            pyautogui.scroll(-gesture.scroll_delta)

    def _move_mouse(self, point: Tuple[int, int]) -> None:
        x, y = point
        pyautogui.moveTo(x, y)

    def update_volume(self, level: float) -> None:
        # Cross-platform volume control usually requires OS-specific APIs.
        # Keep non-destructive by logging target level for extension/plugin layers.
        self.logger.debug("Target volume level: %.2f", level)
