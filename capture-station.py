#!/usr/bin/env python3
"""
capture-station.py — Grab one frame from webcam, crop to calibrated region
===========================================================================
Called by scan-station.bat. Reads station.cfg for camera index and crop coords,
captures a single frame, saves to scans/scan-YYYYMMDD-HHMMSS.jpg

Usage (called by bat, not directly):
    py capture-station.py [--out PATH]
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

try:
    import cv2
except ImportError:
    print("Missing: pip install opencv-python --break-system-packages")
    sys.exit(1)

CFG_PATH = Path("station.cfg")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=None, help="Output path (default: scans/scan-TIMESTAMP.jpg)")
    args = parser.parse_args()

    if not CFG_PATH.exists():
        print("Error: station.cfg not found.")
        print("Run:  py calibrate-station.py  first!")
        sys.exit(1)

    cfg = json.loads(CFG_PATH.read_text())
    cam_idx = cfg.get("active_camera", cfg.get("camera", 0))

    # Load per-camera profile if available, fall back to legacy flat fields
    cam_profiles = cfg.get("cameras", {})
    cam_profile  = cam_profiles.get(str(cam_idx), {})
    crop         = cam_profile.get("crop") or cfg.get("crop")
    cam_w        = cam_profile.get("camera_width")  or cfg.get("camera_width", 1920)
    cam_h        = cam_profile.get("camera_height") or cfg.get("camera_height", 1080)

    if crop:
        print(f"Camera {cam_idx}: crop {crop[2]-crop[0]}x{crop[3]-crop[1]}px")
    else:
        print(f"Camera {cam_idx}: no crop profile — using full frame")

    cap = None
    for backend in [cv2.CAP_MSMF, cv2.CAP_ANY, cv2.CAP_DSHOW]:
        c = cv2.VideoCapture(cam_idx, backend)
        if c.isOpened():
            cap = c
            break
        c.release()

    if cap is None or not cap.isOpened():
        print(f"Error: Cannot open camera {cam_idx}")
        sys.exit(1)

    cap.set(cv2.CAP_PROP_FRAME_WIDTH,  cam_w)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, cam_h)

    # Discard a few frames — webcam needs warmup or first frames are dark
    for _ in range(5):
        cap.read()

    ret, frame = cap.read()
    cap.release()

    if not ret:
        print("Error: Failed to capture frame")
        sys.exit(1)

    # Crop to calibrated region
    if crop:
        x1, y1, x2, y2 = crop
        frame = frame[y1:y2, x1:x2]

    # Save
    if args.out:
        out_path = Path(args.out)
    else:
        scans_dir = Path("scans")
        scans_dir.mkdir(exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = scans_dir / f"scan-{ts}.jpg"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(out_path), frame, [cv2.IMWRITE_JPEG_QUALITY, 95])

    h, w = frame.shape[:2]
    print(f"Captured: {out_path}  ({w}×{h}px)")

    # Print path so bat file can read it
    print(f"OUTPUT_PATH={out_path}")


if __name__ == "__main__":
    main()
