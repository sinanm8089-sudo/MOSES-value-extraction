# AI Gesture Control System with Air Drawing and Shape Recognition

Production-ready real-time hand-gesture system built with **MediaPipe + OpenCV** for:
- Desktop gesture control (mouse, click, scroll, volume intent)
- Air drawing with toolbar and brush controls
- AI-assisted shape recognition (triangle, square, rectangle, circle)
- Gesture-driven UI actions (clear, save, pause)

## Project Structure

```text
gesture_ai/
  __init__.py
  config.py
  hand_tracking.py
  gesture_engine.py
  drawing_engine.py
  shape_recognition.py
  controller.py
  utils/
    __init__.py
    filters.py
    logger.py
main.py
requirements.txt
README.md
```

## Features

### 1) Hand Tracking Engine
- MediaPipe Hands with 21 landmarks per hand
- Up to 2 hands supported
- Optimized camera settings and efficient per-frame parsing

### 2) Gesture Control
- Index finger: cursor mapping with moving-average smoothing
- Thumb + index pinch: click with cooldown
- Two-finger gesture: selection mode + scroll signal
- Finger distance mapped to volume level intent (plugin-ready)

### 3) Air Drawing
- Draw with index fingertip
- Two-finger selection mode for toolbar interactions
- Color palette + eraser
- Adjustable brush thickness (`+`/`-`)
- Transparent canvas overlay
- Save drawing to `outputs/`

### 4) AI Shape Recognition
- Stroke path tracking and contour approximation
- Detects circle, square, rectangle, triangle
- Replaces rough strokes with clean geometric outlines
- Optional shape auto-fill in config

### 5) UI & Mode Controls
- Minimal HUD and top toolbar
- FPS counter
- Status indicators: mode + pause state
- Bounded interaction area for stable tracking
- Mode switching: `DRAW` / `CONTROL` / `IDLE` foundation in app enum

## Controls Guide

### Keyboard
- `q`: Quit
- `m`: Toggle DRAW/CONTROL mode
- `+`: Increase brush size
- `-`: Decrease brush size

### Gestures
- **Index move**: Move pointer / drawing pointer
- **Pinch (thumb+index)**: Click
- **Index+middle up**: Selection mode and scroll intent
- **Fist**: Clear canvas
- **Thumb up**: Save drawing image
- **Open hand**: Pause live actions

## Setup

### 1) Create environment
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows
```

### 2) Install dependencies
```bash
pip install -r requirements.txt
```

### 3) Run app
```bash
python main.py
```

Optional startup mode:
```bash
python main.py --mode CONTROL
```

## Demo Instructions

1. Ensure webcam is connected and unobstructed.
2. Launch with `python main.py`.
3. Keep hand inside the on-screen interaction rectangle.
4. Start in `DRAW` mode, test palette selection, draw shapes.
5. Show thumbs up to save image.
6. Press `m` to switch to `CONTROL` and test pointer/click/scroll.

## Production Notes

- Logging: `logs/gesture_ai.log`
- Output images: `outputs/`
- All tunables are centralized in `gesture_ai/config.py`
- `SystemController.update_volume()` is intentionally non-destructive and extension-friendly for OS-specific audio plugins

## Error Handling

- Camera initialization and startup frame-read checks
- Runtime frame-capture warnings
- Safe shutdown of MediaPipe resources, camera, and OpenCV windows

## Bonus Architecture Notes

### Web Version (WebRTC + JS)
- Browser: MediaPipe Hands via JS, canvas overlay in WebGL/2D
- Transport: WebRTC data channel for gesture events
- Optional backend: Python microservice for model-heavy analytics

### Mobile (Flutter)
- Camera feed + platform channels for native CV
- Gesture event bus mapped to UI commands
- Optional on-device MediaPipe task graph

### Plugin System Direction
- Add gesture plugins implementing `on_gesture(state)` and registering with controller
- Add OS-specific plugins for volume/shortcuts/window actions
- Keep core recognition independent of side effects

## Cross-Platform Considerations

- Mouse and scroll are implemented through PyAutoGUI.
- Volume control uses a plugin hook pattern for OS-specific integration.
- The codebase is modular for Windows/macOS/Linux extension.
