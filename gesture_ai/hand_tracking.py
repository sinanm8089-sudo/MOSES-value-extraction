"""MediaPipe hand tracking engine."""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Tuple

import cv2
import mediapipe as mp

from gesture_ai.config import TrackingConfig


@dataclass
class HandData:
    hand_index: int
    handedness: str
    landmarks_px: List[Tuple[int, int]]
    landmarks_norm: List[Tuple[float, float, float]]


class HandTrackingEngine:
    def __init__(self, config: TrackingConfig) -> None:
        self._mp_hands = mp.solutions.hands
        self._hands = self._mp_hands.Hands(
            static_image_mode=False,
            model_complexity=config.model_complexity,
            max_num_hands=config.max_num_hands,
            min_detection_confidence=config.detection_confidence,
            min_tracking_confidence=config.tracking_confidence,
        )
        self._drawer = mp.solutions.drawing_utils

    def process(self, frame_bgr) -> List[HandData]:
        frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
        results = self._hands.process(frame_rgb)
        hands: List[HandData] = []
        if not results.multi_hand_landmarks:
            return hands

        frame_h, frame_w = frame_bgr.shape[:2]

        for i, (lm, handedness) in enumerate(
            zip(results.multi_hand_landmarks, results.multi_handedness)
        ):
            px_landmarks = []
            norm_landmarks = []
            for p in lm.landmark:
                px_landmarks.append((int(p.x * frame_w), int(p.y * frame_h)))
                norm_landmarks.append((p.x, p.y, p.z))
            hands.append(
                HandData(
                    hand_index=i,
                    handedness=handedness.classification[0].label,
                    landmarks_px=px_landmarks,
                    landmarks_norm=norm_landmarks,
                )
            )

        return hands

    def draw(self, frame_bgr, hands: List[HandData]) -> None:
        for hand in hands:
            landmarks = hand.landmarks_px
            for connection in self._mp_hands.HAND_CONNECTIONS:
                pt1, pt2 = landmarks[connection[0]], landmarks[connection[1]]
                cv2.line(frame_bgr, pt1, pt2, (120, 170, 255), 2)
            for idx, (x, y) in enumerate(landmarks):
                color = (255, 255, 255) if idx in (4, 8, 12, 16, 20) else (80, 180, 255)
                cv2.circle(frame_bgr, (x, y), 4, color, -1)

    def close(self) -> None:
        self._hands.close()
