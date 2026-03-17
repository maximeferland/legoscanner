#!/usr/bin/env python3
"""
calibrate-station.py — Webcam crop calibration (per-camera profiles)
======================================================================
Shows your webcam feed live. Draw a rectangle around the white board.
Saves crop coordinates to station.cfg under a per-camera profile.
Switching cameras automatically loads the correct saved crop.

Usage:
    py calibrate-station.py              (camera 0)
    py calibrate-station.py --camera 1   (iPhone, DroidCam, etc.)

Controls:
    Click + drag  → draw crop rectangle
    ENTER         → save and quit
    R             → reset / redraw
    ESC           → quit without saving

station.cfg structure (one profile per camera):
    {
      "active_camera": 0,
      "cameras": {
        "0": { "crop": [x1,y1,x2,y2], "crop_width": W, "crop_height": H,
                "camera_width": CW, "camera_height": CH },
        "1": { "crop": [...], ... }
      }
    }
"""

import sys
import json
import argparse
from pathlib import Path

try:
    import cv2
except ImportError:
    print("Missing: pip install opencv-python --break-system-packages")
    sys.exit(1)

CFG_PATH = Path("station.cfg")

# ── Mouse state ───────────────────────────────────────────────────────────────
drawing   = False
start_x   = start_y = 0
rect      = None
temp_rect = None


def mouse_cb(event, x, y, flags, param):
    global drawing, start_x, start_y, rect, temp_rect

    if event == cv2.EVENT_LBUTTONDOWN:
        drawing = True
        start_x, start_y = x, y
        rect = None
        temp_rect = None

    elif event == cv2.EVENT_MOUSEMOVE and drawing:
        temp_rect = (min(start_x, x), min(start_y, y),
                     max(start_x, x), max(start_y, y))

    elif event == cv2.EVENT_LBUTTONUP:
        drawing = False
        x1, y1 = min(start_x, x), min(start_y, y)
        x2, y2 = max(start_x, x), max(start_y, y)
        temp_rect = None
        if (x2 - x1) > 20 and (y2 - y1) > 20:
            rect = (x1, y1, x2, y2)


def load_cfg():
    if CFG_PATH.exists():
        try:
            return json.loads(CFG_PATH.read_text())
        except Exception:
            pass
    return {"active_camera": 0, "cameras": {}}


def save_cfg(cfg):
    CFG_PATH.write_text(json.dumps(cfg, indent=2))


def main():
    global rect, temp_rect

    parser = argparse.ArgumentParser()
    parser.add_argument("--camera", type=int, default=0,
                        help="Camera index to calibrate (default 0)")
    parser.add_argument("--stream", type=str, default="", help="HTTP MJPEG stream URL")
    parser.add_argument("--image",  type=str, default="", help="Static image file (no camera needed)")
    args = parser.parse_args()
    cam_key = str(args.camera) if not args.stream else "stream"

    # ── Static image mode — scan-gui passes its last frame so we never fight over the device
    _static_frame = None
    _http_stream  = None
    _http_buf     = b""
    cap           = None
    actual_w, actual_h = 1920, 1080

    if args.image and Path(args.image).exists():
        import numpy as np
        _static_frame = cv2.imread(args.image)
        if _static_frame is not None:
            actual_h, actual_w = _static_frame.shape[:2]
            print(f"Using static frame: {Path(args.image).name}  ({actual_w}x{actual_h}px)")
        else:
            print(f"Warning: Could not load {args.image} — falling back to camera")

    if _static_frame is None:
        if args.stream:
            import urllib.request
            try:
                _http_stream = urllib.request.urlopen(args.stream, timeout=5)
                print(f"Connected to stream: {args.stream}")
            except Exception as e:
                print(f"Error: Cannot connect to stream: {e}")
                return
            actual_w, actual_h = 1280, 720
            print("Detecting stream resolution from first frame...")
            _detect_buf = b""
            for _ in range(500):
                _detect_buf += _http_stream.read(4096)
                a  = _detect_buf.find(b'\xff\xd8')
                b_ = _detect_buf.find(b'\xff\xd9')
                if a != -1 and b_ != -1 and b_ > a:
                    import numpy as np
                    jpg = _detect_buf[a:b_+2]
                    _http_buf = _detect_buf[b_+2:]
                    arr = np.frombuffer(jpg, dtype=np.uint8)
                    _first = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                    if _first is not None:
                        actual_h, actual_w = _first.shape[:2]
                    break
            print(f"Stream resolution: {actual_w}x{actual_h}px")
        else:
            import time as _time
            print(f"Opening camera {args.camera}...")
            # Retry for up to 5 seconds — scan-gui may still be releasing the device
            deadline = _time.time() + 5.0
            while _time.time() < deadline and cap is None:
                for backend in [cv2.CAP_MSMF, cv2.CAP_DSHOW]:
                    c = cv2.VideoCapture(args.camera, backend)
                    if c.isOpened():
                        for _ in range(20):
                            ret, _ = c.read()
                            if ret:
                                cap = c
                                bname = 'MSMF' if backend == cv2.CAP_MSMF else 'DSHOW'
                                print(f"Camera {args.camera} ready ({bname})")
                                break
                        if cap:
                            break
                        c.release()
                if cap is None:
                    print(f"Waiting for camera {args.camera}...")
                    _time.sleep(0.5)
            if cap is None:
                print(f"Error: Could not open camera {args.camera} after 5 seconds")
                sys.exit(1)
            cap.set(cv2.CAP_PROP_FRAME_WIDTH,  1920)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
            actual_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            actual_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            print(f"Camera {args.camera}: {actual_w}x{actual_h}px")

    # Load existing profile for this camera
    cfg = load_cfg()
    cam_profiles = cfg.get("cameras", {})
    if cam_key in cam_profiles:
        saved = cam_profiles[cam_key].get("crop")
        if saved:
            # Scale saved crop if resolution has changed since last calibration
            saved_w = cam_profiles[cam_key].get("camera_width",  actual_w)
            saved_h = cam_profiles[cam_key].get("camera_height", actual_h)
            if saved_w != actual_w or saved_h != actual_h:
                sx = actual_w / saved_w
                sy = actual_h / saved_h
                saved = [int(saved[0]*sx), int(saved[1]*sy),
                         int(saved[2]*sx), int(saved[3]*sy)]
                print(f"Scaled crop from {saved_w}x{saved_h} → {actual_w}x{actual_h}")
            rect = tuple(saved)
            print(f"Loaded existing crop: {rect}")
    else:
        print(f"No existing calibration for camera {args.camera} — draw a new one")

    if cam_profiles:
        print("\nAll saved camera profiles:")
        for k, v in cam_profiles.items():
            marker = "  <- current" if k == cam_key else ""
            print(f"  Camera {k}: {v.get('crop_width','?')}x{v.get('crop_height','?')}px{marker}")

    print("\nInstructions:")
    print("  Click and drag to draw rectangle around your white board")
    print("  ENTER = save   R = reset   ESC = cancel\n")

    mode_str = "Static" if _static_frame is not None else ("Stream" if args.stream else f"Camera {args.camera}")
    win = f"Calibrate {mode_str} — Draw rectangle, then ENTER"
    cv2.namedWindow(win, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(win, min(actual_w, 1280), min(actual_h, 800))
    cv2.setMouseCallback(win, mouse_cb)

    _fail_count = 0
    while True:
        if _static_frame is not None:
            frame, ret = _static_frame.copy(), True
        elif _http_stream:
            import numpy as np
            frame, ret = None, False
            for _ in range(300):
                _http_buf += _http_stream.read(4096)
                a  = _http_buf.find(b'\xff\xd8')
                b_ = _http_buf.find(b'\xff\xd9')
                if a != -1 and b_ != -1 and b_ > a:
                    jpg = _http_buf[a:b_+2]; _http_buf = _http_buf[b_+2:]
                    arr   = np.frombuffer(jpg, dtype=np.uint8)
                    frame = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                    ret   = frame is not None
                    break
        else:
            ret, frame = cap.read()
        if not ret or frame is None:
            _fail_count += 1
            if _fail_count > 50:
                print("Camera stopped responding — exiting")
                break
            import time as _t; _t.sleep(0.05)
            continue  # retry
        _fail_count = 0  # reset on good frame

        display = frame.copy()
        draw_r = temp_rect if temp_rect is not None else rect

        if draw_r:
            x1, y1, x2, y2 = draw_r
            overlay = display.copy()
            cv2.rectangle(overlay, (0,0), (display.shape[1], display.shape[0]), (0,0,0), -1)
            cv2.rectangle(overlay, (x1,y1), (x2,y2), (0,0,0), -1)
            cv2.addWeighted(overlay, 0.35, display, 0.65, 0, display)
            color = (0,255,0) if draw_r is rect else (0,200,255)
            cv2.rectangle(display, (x1,y1), (x2,y2), color, 2)
            label = f"{x2-x1}x{y2-y1}px  ({x1},{y1})-({x2},{y2})"
            cv2.putText(display, label, (x1, max(y1-8, 20)),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

        msg   = "ENTER = save  |  R = redraw  |  ESC = cancel" if rect else "Click and drag to select the board area"
        color = (0,255,0) if rect else (0,200,255)
        cv2.putText(display, msg, (10, display.shape[0]-12),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)
        cv2.putText(display, f"Camera {args.camera}  |  {len(cam_profiles)} profile(s) saved",
                    (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (200,200,200), 1)

        cv2.imshow(win, display)
        key = cv2.waitKey(1) & 0xFF

        if key in (13, ord('\r')):
            if rect:
                break
            else:
                print("Draw a rectangle first!")
        elif key in (ord('r'), ord('R')):
            rect = None
            temp_rect = None
            print("Reset — draw again")
        elif key == 27:
            print("Cancelled — nothing saved")
            if cap: cap.release()
            cv2.destroyAllWindows()
            sys.exit(0)

    if cap: cap.release()
    cv2.destroyAllWindows()

    if rect:
        x1, y1, x2, y2 = rect
        cfg = load_cfg()
        if "cameras" not in cfg:
            cfg["cameras"] = {}

        # Save this camera's profile
        cfg["cameras"][cam_key] = {
            "camera_width":  actual_w,
            "camera_height": actual_h,
            "crop":          [x1, y1, x2, y2],
            "crop_width":    x2 - x1,
            "crop_height":   y2 - y1,
        }
        cfg["active_camera"] = args.camera

        # Legacy flat fields so existing code still works
        cfg["camera"]        = args.camera
        cfg["camera_width"]  = actual_w
        cfg["camera_height"] = actual_h
        cfg["crop"]          = [x1, y1, x2, y2]
        cfg["crop_width"]    = x2 - x1
        cfg["crop_height"]   = y2 - y1

        save_cfg(cfg)

        print(f"\n  Saved profile for camera {args.camera}")
        print(f"    Resolution: {actual_w}x{actual_h}")
        print(f"    Crop: ({x1},{y1}) to ({x2},{y2})  =  {x2-x1}x{y2-y1}px")

        all_cams = cfg["cameras"]
        if len(all_cams) > 1:
            print(f"\nAll saved profiles:")
            for k, v in all_cams.items():
                marker = " <- active" if k == cam_key else ""
                print(f"  Camera {k}: {v.get('crop_width','?')}x{v.get('crop_height','?')}px{marker}")

        print(f"\nTo calibrate another camera:  py calibrate-station.py --camera 1")
        print(f"To start scanning:            py scan-gui.py")


if __name__ == "__main__":
    main()
