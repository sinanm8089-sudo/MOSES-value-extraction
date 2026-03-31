"""Gesture parsing and action generation."""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from gesture_ai.config import GestureConfig
from gesture_ai.hand_tracking import HandData
from gesture_ai.utils.filters import MovingAverageSmoother, clamp, normalize, point_distance


@dataclass
class GestureState:
    mouse_pos: Tuple[int, int] | None = None
    click: bool = False
    scroll_delta: int = 0
    volume_level: float | None = None
    draw_point: Tuple[int, int] | None = None
    select_mode: bool = False
    fist: bool = False
    thumbs_up: bool = False
    open_hand: bool = False


class GestureEngine:
    def __init__(self, config: GestureConfig, screen_size: Tuple[int, int]) -> None:
        self.config = config
        self.screen_w, self.screen_h = screen_size
        self.smoother = MovingAverageSmoother(config.move_smoothing_window)
        self._last_click_time = 0.0

    def parse(self, hand: Optional[HandData], frame_size: Tuple[int, int]) -> GestureState:
        state = GestureState()
        if hand is None:
            self.smoother.reset()
            return state

        lm = hand.landmarks_px
        finger_up = self._finger_states(lm)
        idx_tip = lm[8]
        thumb_tip = lm[4]
        middle_tip = lm[12]

        state.fist = not any(finger_up.values())
        state.thumbs_up = finger_up["thumb"] and not any(
            finger_up[k] for k in ("index", "middle", "ring", "pinky")
        )
        state.open_hand = all(finger_up.values())

        # mouse move / draw pointer
        frame_w, frame_h = frame_size
        margin = self.config.interaction_area_margin
        ix = clamp(idx_tip[0], margin, frame_w - margin)
        iy = clamp(idx_tip[1], margin, frame_h - margin)
        mapped_x = int(normalize(ix, (margin, frame_w - margin), (0, self.screen_w)) * self.config.sensitivity)
        mapped_y = int(normalize(iy, (margin, frame_h - margin), (0, self.screen_h)) * self.config.sensitivity)
        state.mouse_pos = self.smoother.update((mapped_x, mapped_y))
        state.draw_point = idx_tip

        # pinch click
        pinch_distance = point_distance(idx_tip, thumb_tip)
        now = time.time()
        if pinch_distance < self.config.pinch_click_distance_px and now - self._last_click_time > self.config.click_cooldown_sec:
            state.click = True
            self._last_click_time = now

        # two fingers for selection / scroll
        if finger_up["index"] and finger_up["middle"] and not finger_up["ring"] and not finger_up["pinky"]:
            state.select_mode = True
            state.scroll_delta = int((middle_tip[1] - idx_tip[1]) / self.config.scroll_speed_factor)

        # volume distance map
        thumb_index_distance = point_distance(thumb_tip, idx_tip)
        state.volume_level = clamp(
            normalize(
                thumb_index_distance,
                (self.config.volume_min_distance_px, self.config.volume_max_distance_px),
                (0.0, 1.0),
            ),
            0.0,
            1.0,
        )

        return state

    @staticmethod
    def _finger_states(lm: Dict[int, Tuple[int, int]] | list[Tuple[int, int]]) -> Dict[str, bool]:
        tip_ids = {"thumb": 4, "index": 8, "middle": 12, "ring": 16, "pinky": 20}
        pip_ids = {"thumb": 3, "index": 6, "middle": 10, "ring": 14, "pinky": 18}
        states = {}
        for name in tip_ids:
            tip = lm[tip_ids[name]]
            pip = lm[pip_ids[name]]
            states[name] = tip[1] < pip[1] if name != "thumb" else tip[0] > pip[0]
        return states
