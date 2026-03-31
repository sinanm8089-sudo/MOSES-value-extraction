"""Entry point for AI Gesture Control System."""
from __future__ import annotations

import argparse
import time
from enum import Enum

import cv2
import pyautogui

from gesture_ai.config import CONFIG
from gesture_ai.controller import SystemController
from gesture_ai.drawing_engine import DrawingEngine
from gesture_ai.gesture_engine import GestureEngine
from gesture_ai.hand_tracking import HandTrackingEngine
from gesture_ai.utils.logger import setup_logger


class AppMode(str, Enum):
    IDLE = "IDLE"
    DRAW = "DRAW"
    CONTROL = "CONTROL"


class GestureApp:
    def __init__(self) -> None:
        self.logger = setup_logger(CONFIG.log_file)
        self.mode = AppMode.DRAW
        self.paused = False
        self._init_camera()
        self.tracker = HandTrackingEngine(CONFIG.tracking)
        self.controller = SystemController(self.logger)
        screen_size = pyautogui.size()
        self.gesture_engine = GestureEngine(CONFIG.gesture, screen_size=screen_size)
        ok, frame = self.cap.read()
        if not ok:
            raise RuntimeError("Unable to read from camera on startup.")
        self.drawer = DrawingEngine(frame.shape, CONFIG.drawing, CONFIG.ui, CONFIG.output_dir)
        self.prev_time = time.time()

    def _init_camera(self) -> None:
        self.cap = cv2.VideoCapture(CONFIG.camera.device_index)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, CONFIG.camera.width)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CONFIG.camera.height)
        self.cap.set(cv2.CAP_PROP_FPS, CONFIG.camera.target_fps)
        if not self.cap.isOpened():
            raise RuntimeError("Camera could not be opened.")

    def run(self) -> None:
        self.logger.info("Gesture AI started")
        while True:
            ok, frame = self.cap.read()
            if not ok:
                self.logger.warning("Frame capture failed")
                continue

            frame = cv2.flip(frame, 1)
            hands = self.tracker.process(frame)
            primary_hand = hands[0] if hands else None
            gesture = self.gesture_engine.parse(primary_hand, frame_size=(frame.shape[1], frame.shape[0]))

            self._handle_mode_switching(gesture)
            if not self.paused:
                self._handle_actions(gesture)

            self.tracker.draw(frame, hands)
            self.drawer.compose(frame)
            self._draw_hud(frame)

            cv2.imshow("AI Gesture Control System", frame)
            key = cv2.waitKey(1) & 0xFF
            if key == ord("q"):
                break
            if key == ord("m"):
                self.mode = AppMode.CONTROL if self.mode != AppMode.CONTROL else AppMode.DRAW
            if key == ord("+"):
                self.drawer.state.brush_size = min(40, self.drawer.state.brush_size + 1)
            if key == ord("-"):
                self.drawer.state.brush_size = max(1, self.drawer.state.brush_size - 1)

        self.shutdown()

    def _handle_mode_switching(self, gesture) -> None:
        if gesture.open_hand:
            self.paused = True
            return
        self.paused = False

        if gesture.thumbs_up:
            output = self.drawer.save()
            self.logger.info("Drawing saved to %s", output)

        if gesture.fist:
            self.drawer.clear()

    def _handle_actions(self, gesture) -> None:
        if self.mode == AppMode.CONTROL:
            self.controller.apply(gesture)
            if gesture.volume_level is not None:
                self.controller.update_volume(gesture.volume_level)
            self.drawer.lift_pen()
            return

        # draw mode
        if gesture.draw_point is None:
            self.drawer.lift_pen()
            return

        if gesture.select_mode:
            self.drawer.handle_selection(gesture.draw_point)
            self.drawer.draw(gesture.draw_point, selecting=True)
        else:
            self.drawer.draw(gesture.draw_point, selecting=False)

    def _draw_hud(self, frame) -> None:
        now = time.time()
        fps = 1.0 / max(now - self.prev_time, 1e-6)
        self.prev_time = now
        cv2.putText(frame, f"FPS: {fps:.1f}", (20, frame.shape[0] - 20), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (80, 255, 150), 2)
        cv2.putText(frame, f"Mode: {self.mode.value}", (frame.shape[1] - 220, 36), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
        cv2.putText(
            frame,
            f"Paused: {'YES' if self.paused else 'NO'}",
            (frame.shape[1] - 220, 66),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.6,
            (255, 180, 120),
            2,
        )
        margin = CONFIG.gesture.interaction_area_margin
        cv2.rectangle(frame, (margin, margin), (frame.shape[1] - margin, frame.shape[0] - margin), (80, 120, 180), 1)

    def shutdown(self) -> None:
        self.logger.info("Shutting down")
        self.tracker.close()
        self.cap.release()
        cv2.destroyAllWindows()


def parse_args():
    parser = argparse.ArgumentParser(description="AI Gesture Control System")
    parser.add_argument("--mode", choices=[m.value for m in AppMode], default=AppMode.DRAW.value)
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    app = GestureApp()
    app.mode = AppMode(args.mode)
    app.run()
