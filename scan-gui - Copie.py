#!/usr/bin/env python3
"""
scan-gui.py — LEGO Scan Station GUI
=====================================
Requirements:
    pip install PyQt5 opencv-python --break-system-packages
    pip install pyserial --break-system-packages  (optional, for USB scale)

Usage:
    py scan-gui.py
"""

import sys, os, json, subprocess, threading, re, platform, webbrowser
from pathlib import Path
CREATE_NO_WINDOW = 0x08000000  # Windows: suppress black cmd console

# ── Rebrickable parts.csv ─────────────────────────────────────────────────────
_RB_PARTS: dict = {}
_RB_IMG: dict = {}

def _load_rb_inv_csv_bg():
    """Load inventory_parts.csv in background — builds (part_num, color_id) -> img_url index."""
    import csv as _csv, threading as _th
    def _load():
        global _RB_IMG
        csv_path = Path("inventory_parts.csv")
        if not csv_path.exists():
            return
        tmp = {}
        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                for row in _csv.DictReader(f):
                    pn  = row.get("part_num", "").strip().lower()
                    cid = row.get("color_id", "").strip()
                    url = row.get("img_url", "").strip()
                    if pn and cid and url and url.startswith("http"):
                        try:
                            key = (pn, int(cid))
                            if key not in tmp:   # keep first (usually element photo)
                                tmp[key] = url
                        except ValueError:
                            pass
        except Exception:
            pass
        _RB_IMG = tmp
    _th.Thread(target=_load, daemon=True).start()

def rb_img_url(part_num: str, color_id: int) -> str:
    """Look up Rebrickable image URL for part+color. Returns "" if not found."""
    return _RB_IMG.get((part_num.lower().strip(), color_id), "")

def _rb_cached_path(part_num: str, color_id: int) -> Path:
    _IMAGE_CACHE_DIR.mkdir(exist_ok=True)
    ext = "jpg"
    return _IMAGE_CACHE_DIR / f"{part_num.lower()}_{color_id}.{ext}"

def _load_rb_parts_csv_bg():
    """Load parts.csv in background thread — never blocks UI."""
    import csv as _csv, threading as _th
    def _load():
        global _RB_PARTS
        csv_path = Path("parts.csv")
        if not csv_path.exists():
            return
        count = 0
        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                for row in _csv.DictReader(f):
                    pn = row.get("part_num", "").strip().lower()
                    if pn:
                        _RB_PARTS[pn] = {
                            "name":     row.get("name", "").strip(),
                            "cat_id":   row.get("part_cat_id", "").strip(),
                            "material": row.get("part_material", "").strip(),
                        }
                        count += 1
        except Exception:
            pass
    _th.Thread(target=_load, daemon=True).start()

def rb_part_name(part_num: str) -> str:
    return _RB_PARTS.get(part_num.lower().strip(), {}).get("name", "")

def rb_part_material(part_num: str) -> str:
    return _RB_PARTS.get(part_num.lower().strip(), {}).get("material", "")

_python_exe_cache = None
def _python_exe():
    """Return pythonw.exe on Windows (windowless), python elsewhere. Cached after first call."""
    global _python_exe_cache
    if _python_exe_cache:
        return _python_exe_cache
    import sys, platform
    if platform.system() != "Windows":
        _python_exe_cache = sys.executable
        return _python_exe_cache
    exe = sys.executable
    for s in ["python.exe", "python3.exe"]:
        if exe.lower().endswith(s):
            w = exe[:-len(s)] + "pythonw.exe"
            if Path(w).exists():
                _python_exe_cache = w
                return w
    # Last resort: use current exe (console will show but won't hang)
    _python_exe_cache = exe
    return exe

def _hidden_popen_kwargs():
    """Return kwargs that suppress the black cmd window on Windows."""
    import platform
    if platform.system() != "Windows":
        return {}
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0  # SW_HIDE
    return {"startupinfo": si, "creationflags": CREATE_NO_WINDOW}
os.environ['OPENCV_LOG_LEVEL'] = 'ERROR'  # suppress cv2 backend warnings
_IMAGE_CACHE_DIR = Path("image_cache")
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTextEdit, QComboBox, QSlider, QGroupBox,
    QGridLayout, QSpinBox, QDoubleSpinBox, QCheckBox, QSplitter,
    QScrollArea, QInputDialog, QDialog, QListWidget, QListWidgetItem,
    QDialogButtonBox, QLineEdit, QFileDialog, QMenu, QShortcut, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QColor, QPalette, QPixmap, QTextCursor, QImage, QKeySequence

# ── Theme — BrickStore-inspired professional palette ─────────────────────────
CFG_PATH   = Path("station.cfg")
DARK_BG    = "#1c1c1c"      # near-black window base
PANEL_BG   = "#242424"      # panel background
CARD_BG    = "#2e2e2e"      # card / input background
BORDER     = "#3a3a3a"      # subtle borders
ACCENT     = "#c0392b"      # error / delete red
ACCENT2    = "#e8a020"      # orange highlight (BrickStore's primary accent)
TEXT       = "#e8e8e8"      # primary text
TEXT_DIM   = "#777777"      # secondary / muted text
SUCCESS    = "#27ae60"      # confirmed green
WARNING    = "#d4851a"      # warning amber
CONSOLE_BG = "#141414"      # console darker than panels
ROW_ALT    = "#262626"      # alternating row tint
ROW_SEL    = "#3a3020"      # selected row tint

def _load_theme_pref(): return "dark"  # placeholder

COLORS = [
    ("— Auto detect —", ""),
    ("White (1)",              "1"),
    ("Black (11)",             "11"),
    ("Light Bluish Gray (86)", "86"),
    ("Dark Bluish Gray (85)",  "85"),
    ("Light Gray (9)",         "9"),
    ("Dark Gray (10)",         "10"),
    ("Red (5)",                "5"),
    ("Dark Red (59)",          "59"),
    ("Orange (4)",             "4"),
    ("Dark Orange (68)",       "68"),
    ("Yellow (3)",             "3"),
    ("Bright Light Orange (110)","110"),
    ("Lime (34)",              "34"),
    ("Bright Green (36)",      "36"),
    ("Green (6)",              "6"),
    ("Dark Green (80)",        "80"),
    ("Dark Turquoise (39)",    "39"),
    ("Medium Azure (156)",     "156"),
    ("Dark Azure (153)",       "153"),
    ("Blue (7)",               "7"),
    ("Dark Blue (63)",         "63"),
    ("Medium Blue (42)",       "42"),
    ("Dark Purple (89)",       "89"),
    ("Purple (24)",            "24"),
    ("Magenta (71)",           "71"),
    ("Dark Pink (47)",         "47"),
    ("Bright Pink (104)",      "104"),
    ("Reddish Brown (88)",     "88"),
    ("Brown (8)",              "8"),
    ("Dark Brown (120)",       "120"),
    ("Tan (2)",                "2"),
    ("Dark Tan (69)",          "69"),
    ("Light Nougat (90)",      "90"),
    ("Nougat (28)",            "28"),
    ("Sand Green (48)",        "48"),
    ("Olive Green (155)",      "155"),
    ("Pearl Gold (115)",       "115"),
    ("Flat Silver (95)",       "95"),
    ("Pearl Dark Gray (77)",   "77"),
    ("Trans-Clear (12)",           "12"),
    ("Trans-Brown (13)",            "13"),
    ("Trans-Red (17)",              "17"),
    ("Trans-Neon Orange (18)",      "18"),
    ("Trans-Orange (98)",           "98"),
    ("Trans-Yellow (19)",           "19"),
    ("Trans-Neon Yellow (121)",     "121"),
    ("Trans-Neon Green (16)",       "16"),
    ("Trans-Bright Green (108)",    "108"),
    ("Trans-Green (20)",            "20"),
    ("Trans-Dark Blue (14)",        "14"),
    ("Trans-Medium Blue (74)",      "74"),
    ("Trans-Light Blue (15)",       "15"),
    ("Trans-Aqua (113)",            "113"),
    ("Trans-Purple (51)",           "51"),
    ("Trans-Dark Pink (50)",        "50"),
    ("Trans-Pink (107)",            "107"),
    ("Trans-Black (251)",           "251"),
]


# BrickLink color palette — kept in sync with scan-heads.py
# Load color table from scan-heads.py at import time — single source of truth
def _load_bl_colors():
    try:
        import importlib.util, os
        spec = importlib.util.spec_from_file_location(
            "scan_heads", os.path.join(os.path.dirname(__file__) or ".", "scan-heads.py"))
        sh = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sh)
        return sh.BRICKLINK_COLORS
    except Exception:
        return {}

BL_COLORS = _load_bl_colors()
BL_COLORS_SORTED = sorted(BL_COLORS.items(), key=lambda x: x[1][0])

# ── Worker ────────────────────────────────────────────────────────────────────
class ScanWorker(QThread):
    log            = pyqtSignal(str, str)
    preview        = pyqtSignal(str)
    part_found     = pyqtSignal(dict)
    finished       = pyqtSignal(bool, str)
    progress       = pyqtSignal(int)
    step           = pyqtSignal(str)
    detected_count = pyqtSignal(int)

    def __init__(self, settings, main_win):
        super().__init__()
        self.settings = settings
        self.main_win = main_win

    def run(self):
        s = self.settings
        try:
            import sys as _sys
            def _thread_except(t, v, tb):
                import traceback
                traceback.print_exception(t, v, tb)
                self.log.emit(f"Unexpected error: {v}", "error")
                self.finished.emit(False, "")
            threading.excepthook = lambda a: _thread_except(a.exc_type, a.exc_value, a.exc_traceback)
        except Exception:
            pass
        try:
            self.step.emit("📷  Taking photo...")
            self.log.emit("📷  Step 1/3 — Taking photo...", "info")
            self.progress.emit(10)

            # Grab frame directly from forced image path or warm live camera if available
            scan_path = None
            # Forced image path — skip camera entirely
            if s.get("forced_image_path"):
                scan_path = s["forced_image_path"]
                self.log.emit(f"📂  Using loaded image: {Path(scan_path).name}", "success")
            if scan_path is None and self.main_win._last_frame is not None and (self.main_win._live_running or self.main_win._stream_active):
                import cv2
                from datetime import datetime
                scans_dir = Path("scans")
                scans_dir.mkdir(exist_ok=True)
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                scan_path = str(scans_dir / f"scan-{ts}.jpg")
                frame = self.main_win._last_frame.copy()
                crop  = self.main_win._live_crop

                def _crop_frame(f, c):
                    if c:
                        x1,y1,x2,y2 = c
                        fh,fw = f.shape[:2]
                        x1=max(0,min(x1,fw-1)); x2=max(x1+1,min(x2,fw))
                        y1=max(0,min(y1,fh-1)); y2=max(y1+1,min(y2,fh))
                        return f[y1:y2, x1:x2]
                    return f

                if s.get("dual_cam") and s.get("cam2_idx", -1) >= 0:
                    cam2_idx = s["cam2_idx"]
                    self.log.emit(f"📷+📷  Dual capture — grabbing camera {cam2_idx}...", "info")
                    frame2 = None
                    try:
                        known2   = getattr(self.main_win, "_cam_backends", {}).get(cam2_idx)
                        backends = list(dict.fromkeys(
                            [known2, cv2.CAP_DSHOW, cv2.CAP_MSMF] if known2
                            else [cv2.CAP_DSHOW, cv2.CAP_MSMF]))
                        for backend in backends:
                            cap2 = cv2.VideoCapture(cam2_idx, backend)
                            if not cap2.isOpened():
                                cap2.release(); continue
                            # Warmup — flush stale buffer frames
                            got2 = False
                            for _ in range(30):
                                ret2, f2 = cap2.read()
                                if ret2 and f2 is not None and f2.size > 0:
                                    frame2 = f2; got2 = True; break
                                time.sleep(0.05)
                            cap2.release()
                            if got2:
                                self.log.emit(f"📷+📷  Camera {cam2_idx} captured (backend {backend})", "info")
                                break
                            self.log.emit(f"📷+📷  Camera {cam2_idx} no frame on backend {backend} — retrying...", "info")
                            time.sleep(0.3)
                    except Exception as e:
                        self.log.emit(f"⚠  Camera 2 capture failed: {e}", "warning")

                    if frame2 is not None:
                        f1 = _crop_frame(frame, crop)
                        f2 = _crop_frame(frame2, crop)
                        if f1.shape[0] != f2.shape[0]:
                            f2 = cv2.resize(f2, (int(f2.shape[1]*f1.shape[0]/f2.shape[0]), f1.shape[0]))
                        import numpy as _np
                        divider = _np.zeros((f1.shape[0], 4, 3), dtype=_np.uint8)
                        stitched = _np.hstack([f1, divider, f2])
                        cv2.imwrite(scan_path, stitched, [cv2.IMWRITE_JPEG_QUALITY, 95])
                        self.log.emit(f"📷+📷  Dual frame stitched ({f1.shape[1]}+{f2.shape[1]}px wide)", "success")
                    else:
                        self.log.emit("⚠  Camera 2 unavailable — single-cam fallback", "warning")
                        cv2.imwrite(scan_path, _crop_frame(frame, crop), [cv2.IMWRITE_JPEG_QUALITY, 95])
                else:
                    frame = _crop_frame(frame, crop)
                    cv2.imwrite(scan_path, frame, [cv2.IMWRITE_JPEG_QUALITY, 95])
                    self.log.emit(f"📸  Photo captured instantly from live feed", "success")
            if scan_path is None:
                # Fallback — no live feed and no forced path, launch capture-station.py
                self.log.emit("📷  No live feed — launching capture script...", "info")
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                result = subprocess.run([_python_exe(),"capture-station.py"],
                                     capture_output=True, text=True,
                                     encoding="utf-8", errors="replace",
                                     env=env, cwd=str(Path.cwd()),
                                     **_hidden_popen_kwargs())
                if result.returncode != 0:
                    self.log.emit(f"Capture failed:\n{result.stderr}", "error")
                    self.finished.emit(False, ""); return
                for line in result.stdout.splitlines():
                    if line.startswith("OUTPUT_PATH="):
                        scan_path = line.split("=",1)[1].strip()

            if not scan_path:
                self.log.emit("Could not capture image", "error")
                self.finished.emit(False, ""); return

            self.step.emit("🔎  Detecting parts...")
            self.preview.emit(scan_path)
            self.progress.emit(25)
            self.log.emit(f"🔎  Step 2/3 — Detecting parts...", "info")

            args = [_python_exe(),"scan-heads.py", scan_path,
                    "--mode",       s.get("mode","all"),
                    "--confidence", str(s.get("confidence",0.45)),
                    "--currency",   s.get("currency","CAD"),
                    "--qty",        str(s.get("qty",1)),
                    "--debug", "--annotate"]   # always save crops and annotated image

            if s.get("color",""):
                args += ["--color", s["color"]]
            if s.get("grid",False):
                args += ["--cols", str(s.get("cols",5)), "--rows", str(s.get("rows",4))]
            if s.get("gap", 0) > 0:
                args += ["--gap", str(s.get("gap", 0))]

            if s.get("padding", 10) != 10:
                args += ["--padding", str(s.get("padding", 10))]
            if s.get("studs"):
                args += ["--studs"]
            if s.get("geometric"):
                args += ["--geometric"]
            if s.get("brightness_bias", 0) != 0:
                args += ["--brightness_bias", str(s.get("brightness_bias", 0))]
            if s.get("bg_color"):
                r,g,b = s["bg_color"]
                args += ["--bg-color", f"{r},{g},{b}"]
            if s.get("shadow_color"):
                r,g,b = s["shadow_color"]
                args += ["--shadow-color", f"{r},{g},{b}"]
            # NEW (2026): pass through HD max-side override if set
            if s.get("max_side", 0):
                args += ["--max-side", str(int(s["max_side"]))]

            self.step.emit("🌐  Querying API...")
            self.log.emit(f"🌐  Step 3/3 — Querying Brickognize API...", "info")
            self.progress.emit(35)

            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            self.log.emit(f"⚙  Running: {' '.join(args[:4])}...", "info")
            proc = subprocess.Popen(args, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT, text=True,
                                    encoding="utf-8", errors="replace",
                                    env=env, cwd=str(Path.cwd()), bufsize=1,
                                    **_hidden_popen_kwargs())
            self.main_win._scan_proc = proc  # allow cancel
            xml_path = None
            for line in proc.stdout:
                line = line.rstrip()
                if not line: continue
                self.log.emit(f"  {line}", "info")  # always show raw output
                level = ("success" if "✓" in line or "✅" in line else
                         "error"   if "✗" in line or "error" in line.lower() else
                         "warning" if "⚠" in line else "info")
                if "XML saved:" in line:
                    xml_path = line.split("XML saved:")[-1].strip()
                if "ANNOTATED_PATH=" in line:
                    ann_path = line.split("ANNOTATED_PATH=")[-1].strip()
                    self.preview.emit(ann_path)
                if "parts detected"   in line:
                    self.step.emit(f"🔎  {line.strip()}"); self.progress.emit(45)
                    import re as _re2; _m2 = _re2.search(r"([0-9]+) parts detected", line)
                    if _m2: self.detected_count.emit(int(_m2.group(1)))
                elif "Detecting parts"  in line: self.progress.emit(40)
                elif "Querying"        in line: self.progress.emit(55)
                elif "Fetching"        in line: self.progress.emit(75)
                elif "HTML report"     in line: self.progress.emit(90)

            proc.wait()
            self.progress.emit(100)

            if proc.returncode != 0 or not xml_path:
                self.log.emit("Scan did not complete successfully", "error")
                self.finished.emit(False, ""); return

            # Emit per-part results from JSON
            json_p = Path(xml_path).parent / "scan-results.json"
            crops_d = Path(xml_path).parent / "crops"
            if json_p.exists():
                try:
                    entries = json.loads(json_p.read_text())
                    self.log.emit(f"  📋  {len(entries)} result(s) from JSON", "info")

                    # ── Dual-cam deduplication ────────────────────────────────
                    # If dual_cam was used, the stitched image is two cam widths wide.
                    # A part detected in cam1 (left half) and cam2 (right half) at the
                    # same physical position will have bboxes offset by ~cam1_width px.
                    # We merge such pairs: keep highest confidence, inject the loser as
                    # an alternative so the user can still see the other camera's result.
                    if s.get("dual_cam") and entries:
                        # Estimate cam1 width = half the stitched image width
                        # Best source: first bbox x2 max gives us a hint, but we can
                        # read it directly from the scan image dimensions.
                        try:
                            import cv2 as _cv2
                            _si = _cv2.imread(scan_path)
                            stitch_w = _si.shape[1] if _si is not None else 0
                            cam1_w = stitch_w // 2  # approx — divider is only 4px
                        except Exception:
                            cam1_w = 0

                        if cam1_w > 0:
                            # Split entries into left-half (cam1) and right-half (cam2)
                            left  = [e for e in entries if e.get("bbox") and e["bbox"][0] < cam1_w]
                            right = [e for e in entries if e.get("bbox") and e["bbox"][0] >= cam1_w]
                            unmatched_right = list(right)
                            merged = []

                            for le in left:
                                lx1, ly1, lx2, ly2 = le["bbox"]
                                # Translate to right-half coordinates for overlap test
                                lx1r, lx2r = lx1 + cam1_w, lx2 + cam1_w
                                best_match = None
                                best_iou   = 0.0
                                for re_ in unmatched_right:
                                    rx1, ry1, rx2, ry2 = re_["bbox"]
                                    # IoU in translated space
                                    ix1 = max(lx1r, rx1); iy1 = max(ly1,  ry1)
                                    ix2 = min(lx2r, rx2); iy2 = min(ly2,  ry2)
                                    if ix2 > ix1 and iy2 > iy1:
                                        inter = (ix2-ix1)*(iy2-iy1)
                                        area_l = (lx2r-lx1r)*(ly2-ly1)
                                        area_r = (rx2-rx1)*(ry2-ry1)
                                        iou = inter / (area_l + area_r - inter)
                                        if iou > best_iou:
                                            best_iou = iou
                                            best_match = re_

                                if best_match and best_iou > 0.25:
                                    # Same physical part — keep highest confidence
                                    unmatched_right.remove(best_match)
                                    lconf = le.get("confidence", 0)
                                    rconf = best_match.get("confidence", 0)
                                    winner = le if lconf >= rconf else best_match
                                    loser  = best_match if lconf >= rconf else le
                                    # Inject loser as alternative on the winner
                                    loser_alt = {
                                        "part_id":   loser.get("part_id", ""),
                                        "id":        loser.get("part_id", ""),
                                        "name":      loser.get("part_name", ""),
                                        "score":     loser.get("confidence", 0),
                                        "color_id":  loser.get("color_id"),
                                        "color_name":loser.get("color_name"),
                                        "_dual_cam_alt": True,
                                    }
                                    existing_alts = list(winner.get("alternatives") or [])
                                    existing_alts.insert(0, loser_alt)
                                    winner["alternatives"] = existing_alts
                                    cam_label = "📷1" if winner is le else "📷2"
                                    self.log.emit(
                                        f"  📷+📷  Merged: {le.get('part_id')} "
                                        f"(kept {cam_label} conf "
                                        f"{max(lconf,rconf):.0%} > {min(lconf,rconf):.0%})", "info")
                                    merged.append(winner)
                                else:
                                    merged.append(le)

                            # Add unmatched right-half detections (parts only visible from cam2)
                            merged.extend(unmatched_right)
                            n_deduped = len(left) + len(right) - len(merged)
                            if n_deduped:
                                self.log.emit(
                                    f"  📷+📷  Deduplicated {n_deduped} duplicate part(s) "
                                    f"— alt result stored on each row", "success")
                            entries = merged

                    # ── Emit all entries ──────────────────────────────────────
                    for r in entries:
                        idx = r.get("index", 0)
                        cf = crops_d / f"p1_crop_{idx:03d}.jpg"
                        r["crop_image"] = str(cf) if cf.exists() else ""
                        if not r.get("source_image"):
                            r["source_image"] = scan_path
                        self.part_found.emit(r)
                except Exception as e:
                    import traceback
                    self.log.emit(f"⚠  JSON load failed: {e}", "error")
                    self.log.emit(traceback.format_exc(), "error")
            else:
                self.log.emit(f"⚠  No JSON file at {json_p}", "warning")

            self.log.emit("✅  Done! XML ready.", "success")
            self.finished.emit(True, xml_path)

        except Exception as e:
            self.log.emit(f"Unexpected error: {e}", "error")
            self.finished.emit(False, "")

# ── Main Window ───────────────────────────────────────────────────────────────
class ScanStation(QMainWindow):
    cameras_detected  = pyqtSignal(list)
    log_message       = pyqtSignal(str, str)
    _stream_died      = pyqtSignal()       # emitted from stream thread → triggers camera detection on main thread
    camera_frame      = pyqtSignal(int, int, bytes)  # w, h, rgb bytes
    _bl_img_signal    = pyqtSignal(object, object)     # (QLabel, QPixmap) for BL ref images
    _part_out_ready   = pyqtSignal(object, object)     # (rd, parts_list) for part-out
    _rescan_ready     = pyqtSignal(object, object)     # (rd, entries_list) for rescan
    _po_value_ready   = pyqtSignal(object, float)      # (rd, total_value) for part-out value preview
    _clone_colors_ready = pyqtSignal(object, object)   # (rd, clones_list) for color clone
    _value_colors_ready = pyqtSignal(object, object)   # (rd, clones_list) for value clone
    _value_result_ready = pyqtSignal(str, str, object, object, int)  # discovery popup
    _wt_raw_signal      = pyqtSignal(float)   # raw grams from scale thread
    _wt_unit_ready      = pyqtSignal(float)   # unit weight fetched from API
    calibration_done  = pyqtSignal()
    _price_ready        = pyqtSignal(object, object)  # (row_data dict, price_or_None)
    _price_guide_ready  = pyqtSignal(object, object)  # (row_data, guide_results_dict)
    _iphone_recheck_ready    = pyqtSignal(object, object)  # (rd, best_result_dict)
    _iphone_crop_refresh     = pyqtSignal(object)          # (rd) – refresh thumbnail even if no match

    def __init__(self):
        super().__init__()
        self.setWindowTitle("🧱 LEGO Scan Station")
        _load_rb_parts_csv_bg()   # load parts.csv in background
        _load_rb_inv_csv_bg()      # load inventory_parts.csv image index
        self.setMinimumSize(1150, 720)
        self.worker = None
        self.last_xml_path = None
        self.last_html_path = None
        self._part_count = 0
        self._rows = []
        self._current_source_image = None  # group header tracking for batch
        self._last_preview_source  = None  # last source image shown in preview
        self._last_preview_bbox    = None  # bbox for that preview
        self._output_dir           = "reports"  # matches scan-heads --output default
        self._captures_folder_logged = False
        # iPhone recheck alignment (DroidCam → iPhone transform offset in iPhone pixels)
        self._iphone_dx = 0
        self._iphone_dy = 0

        self.cameras_detected.connect(self._cameras_found)
        self.log_message.connect(self._log)
        self._price_ready.connect(self._on_price_ready)
        self._price_guide_ready.connect(self._on_price_guide_ready)
        self._iphone_recheck_ready.connect(self._on_iphone_recheck_ready)
        self._part_out_ready.connect(self._on_part_out_ready)
        self._rescan_ready.connect(self._on_rescan_ready)
        self._po_value_ready.connect(self._on_po_value_ready)
        self._clone_colors_ready.connect(self._on_clone_colors_ready)
        self._value_colors_ready.connect(self._on_clone_colors_ready)  # reuse same handler
        self._value_result_ready.connect(self._on_value_result)
        self._wt_raw_signal.connect(self._wt_on_raw)
        self._wt_unit_ready.connect(self._wt_on_unit_ready)
        self.camera_frame.connect(self._update_live_preview)
        self._bl_img_signal.connect(self._on_bl_img)
        self.calibration_done.connect(self._load_config)
        self.calibration_done.connect(self._resume_after_calibration)
        self._live_cap = None
        self._live_running = False
        self._live_paused  = False
        self._last_frame = None
        self._live_crop = None   # (x1,y1,x2,y2) from station.cfg, or None
        self._use_http_stream = False
        self._cam_generation  = 0  # incremented each _start_live_camera call — kills stale threads
        self._http_stream_url = ""
        self._stream_active   = False
        self._stream_lock     = threading.Lock()
        self._apply_theme()
        self._build_ui()
        # NEW (2026): keyboard shortcut for HD scan toggle
        # Ctrl+H = toggle HD scan (higher --max-side)
        self._hd_shortcut = QShortcut(QKeySequence("Ctrl+H"), self)
        self._hd_shortcut.activated.connect(self._toggle_hd_scan)
        self._iphone_recheck_ready.connect(self._on_iphone_recheck_ready)
        self._iphone_crop_refresh.connect(self._on_iphone_crop_refresh)
        self._load_config()
        # NEW (2026): persist splitter sizes (debounced) so pane layout is remembered.
        self._ui_save_timer = QTimer(self)
        self._ui_save_timer.setSingleShot(True)
        self._ui_save_timer.timeout.connect(self._save_ui_layout)
        self._wire_splitter_persistence()
        self._stream_died.connect(self._on_stream_died)
        def _maybe_detect_cameras():
            url = self.stream_url.currentText().strip() if hasattr(self, "stream_url") else ""
            if url and not url.startswith("—"):
                self._log("📡  IP stream configured — starting stream...", "info")
                self._live_paused = False
                self._use_http_stream = True
                QTimer.singleShot(200, self._toggle_stream)
            else:
                # No stream — start webcam via camera detection
                self._detect_cameras()
        QTimer.singleShot(100, _maybe_detect_cameras)
        QTimer.singleShot(0, self._update_bulk_bar)   # set initial disabled state
        QTimer.singleShot(0, self._update_preview_toggle)  # set live btn green

    # ── Theme ─────────────────────────────────────────────────────────────────
    def _apply_theme(self):
        self.setStyleSheet(f"""
            QMainWindow {{ background:{DARK_BG}; }}
            QWidget {{
                color:{TEXT};
                font-family:'Segoe UI',Arial,sans-serif; font-size:12px;
            }}
            QGroupBox {{
                background:{PANEL_BG}; border:1px solid {BORDER};
                border-radius:4px; margin-top:10px; padding:6px 4px 4px 4px;
                font-size:10px; font-weight:bold; color:{TEXT_DIM};
                text-transform:uppercase; letter-spacing:1px;
            }}
            QGroupBox::title {{
                subcontrol-origin:margin; left:8px; padding:0 4px;
                color:{ACCENT2};
            }}
            QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit {{
                background:{CARD_BG}; border:1px solid {BORDER};
                border-radius:3px; padding:2px 4px; color:{TEXT};
                font-size:12px; min-height:22px;
                selection-background-color:{ACCENT2}; selection-color:#000;
            }}
            QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover, QLineEdit:hover {{
                border-color:#555;
            }}
            QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QLineEdit:focus {{
                border-color:{ACCENT2};
            }}
            QComboBox::drop-down {{ border:none; width:20px; }}
            QComboBox QAbstractItemView {{
                background:{CARD_BG}; color:{TEXT}; border:1px solid {BORDER};
                selection-background-color:{ACCENT2}; selection-color:#000;
                outline:none;
            }}
            QPushButton {{
                background:{CARD_BG}; color:{TEXT}; border:1px solid {BORDER};
                border-radius:3px; padding:5px 12px; font-size:12px;
            }}
            QPushButton:hover {{ background:#383838; border-color:#555; }}
            QPushButton:pressed {{ background:#222; }}
            QCheckBox {{ color:{TEXT}; font-size:12px; spacing:6px; }}
            QCheckBox::indicator {{
                width:14px; height:14px; border-radius:2px;
                border:1px solid #555; background:{CARD_BG};
            }}
            QCheckBox::indicator:checked {{
                background:{ACCENT2}; border-color:{ACCENT2};
                image:none;
            }}
            QCheckBox::indicator:indeterminate {{
                background:#555; border-color:#777;
            }}
            QLabel {{ color:{TEXT}; }}
            QSlider::groove:horizontal {{
                height:3px; background:{BORDER}; border-radius:2px;
            }}
            QSlider::handle:horizontal {{
                background:{ACCENT2}; width:14px; height:14px;
                margin:-6px 0; border-radius:7px; border:1px solid #222;
            }}
            QSlider::sub-page:horizontal {{
                background:{ACCENT2}; border-radius:2px;
            }}
            QTextEdit {{
                background:{CONSOLE_BG}; color:#aaddaa;
                font-family:'Consolas','Courier New',monospace;
                font-size:11px; border:1px solid {BORDER}; border-radius:3px;
            }}
            QScrollBar:vertical {{
                background:{DARK_BG}; width:8px; border:none;
            }}
            QScrollBar::handle:vertical {{
                background:{BORDER}; border-radius:4px; min-height:20px;
            }}
            QScrollBar::handle:vertical:hover {{ background:#555; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
            QScrollBar:horizontal {{
                background:{DARK_BG}; height:8px; border:none;
            }}
            QScrollBar::handle:horizontal {{
                background:{BORDER}; border-radius:4px; min-width:20px;
            }}
            QSplitter::handle {{ background:{BORDER}; }}
            QScrollArea {{ border:none; background:transparent; }}
            QMenu {{
                background:{CARD_BG}; color:{TEXT}; border:1px solid {BORDER};
                border-radius:4px; padding:4px;
            }}
            QMenu::item {{ padding:5px 20px 5px 12px; border-radius:2px; }}
            QMenu::item:selected {{ background:{ACCENT2}; color:#000; }}
            QMenu::separator {{ height:1px; background:{BORDER}; margin:4px 8px; }}
        """)

    # NEW (2026): HD scan toggle (higher scan-heads --max-side)
    def _toggle_hd_scan(self):
        """Toggle HD scan override. Affects scan-heads downscale limit, not live preview FPS."""
        if getattr(self, "_max_side_override", 0):
            self._max_side_override = 0
            if hasattr(self, "hd_btn"):
                self.hd_btn.setText("HD")
                self.hd_btn.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};border-radius:4px;")
                self.hd_btn.setToolTip("HD scan is OFF (default speed)\nClick to enable for next scans")
            self._log("HD scan: OFF (max-side default)", "info")
        else:
            self._max_side_override = 4200
            if hasattr(self, "hd_btn"):
                self.hd_btn.setText("HD✓")
                self.hd_btn.setStyleSheet(f"background:#1a3a1a;color:#5fca7a;border:1px solid #2a5a2a;border-radius:4px;")
                self.hd_btn.setToolTip("HD scan is ON (slower, more detail)\nUses scan-heads --max-side 4200\nClick to disable")
            self._log("HD scan: ON (max-side 4200)", "success")

    # ── UI Build ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setSpacing(12)
        root.setContentsMargins(6,6,6,6)

        # ── Main horizontal splitter (left panel | right panel) ──────────────
        main_splitter = QSplitter(Qt.Horizontal)
        self._main_splitter = main_splitter
        main_splitter.setChildrenCollapsible(False)
        main_splitter.setHandleWidth(5)
        main_splitter.setStyleSheet("QSplitter::handle:horizontal{background:#404040;width:4px;}")
        # Usability: make right pane take extra space by default.
        main_splitter.setStretchFactor(0, 0)
        main_splitter.setStretchFactor(1, 1)

        left_scroll = QScrollArea()
        left_scroll.setMinimumWidth(220)
        # Usability: don't hard-cap left panel width — let the user resize freely.
        # (Still keeps it sane via minimum width.)
        left_scroll.setMaximumWidth(9999)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setStyleSheet("QScrollArea{border:none;background:transparent;}")
        left_scroll.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        left = QWidget()
        left_scroll.setWidget(left)
        lv = QVBoxLayout(left)
        lv.setSpacing(10)
        lv.setContentsMargins(0,0,4,0)

        # Calibration status
        # Show saved calibration info if station.cfg exists
        import json as _json
        _cfg_ok = False
        _cfg_txt = "ℹ  No calibration yet — live preview shows full frame"
        _cfg_sty = f"font-size:10px;color:{WARNING};padding:4px 8px;background:#2a1800;border-radius:3px;border:1px solid #3a2800;"
        if CFG_PATH.exists():
            try:
                _c = _json.loads(CFG_PATH.read_text())
                _cw = _c.get("crop_width","?"); _ch = _c.get("crop_height","?")
                _cam = _c.get("camera", 0)
                _cfg_txt = f"✓  Camera {_cam} — crop {_cw}×{_ch}px"
                _cfg_sty = f"font-size:10px;color:{SUCCESS};padding:4px 8px;background:#0a1f0a;border-radius:3px;border:1px solid #1a3a1a;"
                _cfg_ok = True
            except Exception:
                pass
        self.cfg_label = QLabel(_cfg_txt)
        self.cfg_label.setStyleSheet(_cfg_sty)
        self.cfg_label.setWordWrap(True)
        lv.addWidget(self.cfg_label)

        # ── Settings group ────────────────────────────────────────────────────
        sg = QGroupBox("")
        gl = QGridLayout(sg)
        gl.setSpacing(5)
        gl.setColumnStretch(0, 0)   # label col — fixed
        gl.setColumnStretch(1, 1)   # value col — stretches
        gl.setColumnMinimumWidth(0, 90)  # label min width

        # Camera
        gl.addWidget(QLabel("Camera:"), 0, 0)
        cw = QWidget(); ch = QHBoxLayout(cw); ch.setContentsMargins(0,0,0,0); ch.setSpacing(4)
        self.camera_combo = QComboBox()
        self.camera_combo.addItem("Detecting...", -1)
        self.refresh_cam_btn = self._btn("⟳", CARD_BG, self._detect_cameras, w=30)
        self.calibrate_btn = self._btn("Calibrate", CARD_BG, self._run_calibrate, w=72)
        self.test_cam_btn = self._btn("Test", "#1a3a1a", self._test_camera, w=40)
        ch.addWidget(self.camera_combo); ch.addWidget(self.refresh_cam_btn)
        ch.addWidget(self.calibrate_btn); ch.addWidget(self.test_cam_btn)
        gl.addWidget(cw, 0, 1)

        # Dual-cam row
        gl.addWidget(QLabel("Dual cam:"), 1, 0)
        dw = QWidget(); dh = QHBoxLayout(dw); dh.setContentsMargins(0,0,0,0); dh.setSpacing(4)
        self.dual_cam_chk = QCheckBox("📷+📷 Enable")
        self.dual_cam_chk.setStyleSheet(f"color:{TEXT};font-size:11px;")
        self.dual_cam_chk.setToolTip(
            "Scan with two cameras simultaneously (different angles, same area).\n"
            "Results are merged — duplicate parts deduplicated by position.")
        self.camera2_combo = QComboBox()
        self.camera2_combo.addItem("— select 2nd cam —", -1)
        self.camera2_combo.setEnabled(False)
        self.camera2_combo.setToolTip("Second camera index for dual-angle scan")
        self.dual_cam_chk.stateChanged.connect(
            lambda s: self.camera2_combo.setEnabled(s == Qt.Checked))
        dh.addWidget(self.dual_cam_chk); dh.addWidget(self.camera2_combo, 1)
        gl.addWidget(dw, 1, 1)

        # DroidCam HTTP stream — hidden (using DroidCam as virtual webcam instead)
        self.stream_url = QComboBox(); self.stream_url.setEditable(True)
        self.stream_url.addItem("— webcam (no stream) —", "")
        self.stream_url.addItem("192.168.1.37:4747/video", "192.168.1.37:4747/video")
        self.stream_url.setCurrentIndex(0)
        self.stream_url.hide()
        self.use_stream_btn = self._btn("Use", CARD_BG, self._toggle_stream, w=38)
        self.use_stream_btn.hide()
        self.webcam_btn = self._btn("📷", CARD_BG, self._switch_to_webcam, w=30)
        self.webcam_btn.hide()

        # Confidence
        gl.addWidget(QLabel("Confidence:"), 3, 0)
        cfw = QWidget(); cfh = QHBoxLayout(cfw); cfh.setContentsMargins(0,0,0,0); cfh.setSpacing(5)
        self.conf_slider = QSlider(Qt.Horizontal)
        self.conf_slider.setRange(20,90); self.conf_slider.setValue(45)
        self.conf_val = QLabel("45%")
        self.conf_val.setFixedWidth(34)
        self.conf_val.setStyleSheet(f"color:{ACCENT2};font-weight:bold;")
        self.conf_slider.valueChanged.connect(lambda v: self.conf_val.setText(f"{v}%"))
        cfh.addWidget(self.conf_slider); cfh.addWidget(self.conf_val)
        gl.addWidget(cfw, 3, 1)

        # Currency
        gl.addWidget(QLabel("Currency:"), 4, 0)
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(["CAD","USD","EUR","GBP"])
        gl.addWidget(self.currency_combo, 4, 1)

        # Force color
        gl.addWidget(QLabel("Force color:"), 5, 0)
        self.color_combo = QComboBox()
        for name, val in COLORS:
            self.color_combo.addItem(name, val)
        gl.addWidget(self.color_combo, 5, 1)

        # Qty
        gl.addWidget(QLabel("Qty per part:"), 6, 0)
        self.qty_spin = QSpinBox()
        self.qty_spin.setRange(1,99); self.qty_spin.setValue(1)
        gl.addWidget(self.qty_spin, 6, 1)

        gl.addWidget(QLabel("Blob gap:"), 7, 0)
        self.gap_spin = QSpinBox()
        self.gap_spin.setRange(0, 300); self.gap_spin.setValue(0)
        self.gap_spin.setSpecialValueText("auto")
        self.gap_spin.setToolTip("Blob merge gap px (0=auto). Lower = less merging between close parts")
        gl.addWidget(self.gap_spin, 7, 1)

        gl.addWidget(QLabel("Detection:"), 8, 0)
        det_w = QWidget(); det_h = QHBoxLayout(det_w)
        det_h.setContentsMargins(0,0,0,0); det_h.setSpacing(4)
        self.det_standard  = QPushButton("Standard")
        self.det_studs     = QPushButton("Flood Fill")
        self.det_studs.setToolTip("Flood-fill from background outward\nRequires background color to be set")
        self.det_geometric = QPushButton("Geometric")
        self._det_mode = "standard"
        def _make_det_style(active):
            return (f"font-size:10px;border-radius:3px;border:1px solid {BORDER};"
                    f"padding:0 4px;background:{SUCCESS if active else CARD_BG};"
                    f"color:{'#000' if active else TEXT};")
        def _set_det(mode):
            self._det_mode = mode
            self.det_standard.setStyleSheet(_make_det_style(mode=="standard"))
            self.det_studs.setStyleSheet(_make_det_style(mode=="studs"))  # studs = flood fill mode
            self.det_geometric.setStyleSheet(_make_det_style(mode=="geometric"))
        for btn, mode in [(self.det_standard,"standard"),(self.det_studs,"studs"),(self.det_geometric,"geometric")]:
            btn.setCheckable(True); btn.setFixedHeight(22)
            btn.setStyleSheet(_make_det_style(mode=="standard"))
            btn.clicked.connect(lambda _, m=mode: _set_det(m))
        det_h.addWidget(self.det_standard); det_h.addWidget(self.det_studs); det_h.addWidget(self.det_geometric)
        gl.addWidget(det_w, 8, 1)

        gl.addWidget(QLabel("Crop padding:"), 9, 0)
        self.padding_spin = QSpinBox()
        self.padding_spin.setRange(0, 80); self.padding_spin.setValue(10)
        self.padding_spin.setSuffix(" %")
        self.padding_spin.setToolTip("Padding around each detected part crop (% of part size)")
        gl.addWidget(self.padding_spin, 9, 1)

        gl.addWidget(QLabel("Color bias:"), 10, 0)
        bias_w = QWidget()
        bias_h = QHBoxLayout(bias_w); bias_h.setContentsMargins(0,0,0,0); bias_h.setSpacing(4)
        self.brightness_bias = QSlider(Qt.Horizontal)
        # Default must be 0 (neutral). Previously this was accidentally non-zero while the label showed "0".
        self.brightness_bias.setRange(-60, 60); self.brightness_bias.setValue(0)
        self.brightness_bias.setTickInterval(20); self.brightness_bias.setTickPosition(QSlider.TicksBelow)
        self.brightness_bias.setToolTip("Shift color sampling brighter (+) or darker (−) before matching\nUse + if your lighting makes parts look darker than they are")
        self.bias_lbl = QLabel("0")
        self.bias_lbl.setFixedWidth(24)
        self.bias_lbl.setStyleSheet(f"color:{TEXT_DIM};font-size:10px;")
        self.brightness_bias.valueChanged.connect(
            lambda v: self.bias_lbl.setText(("+" if v > 0 else "") + str(v)))
        bias_h.addWidget(QLabel("−", styleSheet=f"color:{TEXT_DIM};font-size:9px;"))
        bias_h.addWidget(self.brightness_bias, 1)
        bias_h.addWidget(QLabel("+", styleSheet=f"color:{TEXT_DIM};font-size:9px;"))
        bias_h.addWidget(self.bias_lbl)
        gl.addWidget(bias_w, 10, 1)

        # iPhone photos folder (optional) — used for manual “recheck with latest iPhone photo”
        gl.addWidget(QLabel("iPhone photos:"), 11, 0)
        ipw = QWidget(); iph = QHBoxLayout(ipw); iph.setContentsMargins(0,0,0,0); iph.setSpacing(4)
        self._iphone_photo_dir = ""
        self.iphone_dir_lbl = QLabel("not set")
        self.iphone_dir_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        self.iphone_dir_lbl.setToolTip("Folder containing iPhone photos (newest file is used)")
        self.iphone_dir_btn = self._btn("Set…", CARD_BG, self._set_iphone_photo_dir, w=52)
        self.iphone_dir_btn.setFixedHeight(22)
        self.iphone_calib_btn = self._btn("Calib", CARD_BG, self._calibrate_iphone_from_selection, w=52)
        self.iphone_calib_btn.setFixedHeight(22)
        iph.addWidget(self.iphone_dir_lbl, 1)
        iph.addWidget(self.iphone_dir_btn)
        iph.addWidget(self.iphone_calib_btn)
        gl.addWidget(ipw, 11, 1)

        # Grid mode
        self.grid_check = QCheckBox("Fixed grid mode")
        self.grid_check.setToolTip("Use when parts are arranged in a regular grid")
        gl.addWidget(self.grid_check, 11, 0, 1, 2)
        grid_sub = QWidget()
        gsub = QHBoxLayout(grid_sub); gsub.setContentsMargins(0,0,0,0); gsub.setSpacing(5)
        gsub.addWidget(QLabel("Cols:"))
        self.cols_spin = QSpinBox(); self.cols_spin.setRange(1,20); self.cols_spin.setValue(5)
        gsub.addWidget(self.cols_spin)
        gsub.addWidget(QLabel("Rows:"))
        self.rows_spin = QSpinBox(); self.rows_spin.setRange(1,20); self.rows_spin.setValue(4)
        gsub.addWidget(self.rows_spin)
        gl.addWidget(grid_sub, 12, 0, 1, 2)
        # Grid sub always visible — cols/rows always editable
        self.grid_check.toggled.connect(lambda _: (self._redraw_preview(), self._save_grid_prefs()))
        self.cols_spin.valueChanged.connect(lambda _: (self._redraw_preview(), self._save_grid_prefs()))
        self.rows_spin.valueChanged.connect(lambda _: (self._redraw_preview(), self._save_grid_prefs()))

        # Background color — single clickable button, right-click to clear
        # Restore saved grid prefs
        self._load_grid_prefs()

        gl.addWidget(QLabel("Background:"), 13, 0)
        bg_sub = QWidget()
        bg_h = QHBoxLayout(bg_sub); bg_h.setContentsMargins(0,0,0,0); bg_h.setSpacing(0)
        self._bg_rgb = None

        self.bg_swatch = QLabel("🖱 Right-click preview or here to set background")
        self.bg_swatch.setFixedHeight(24)
        self.bg_swatch.setAlignment(Qt.AlignCenter)
        self.bg_swatch.setCursor(Qt.PointingHandCursor)
        self.bg_swatch.setStyleSheet(
            f"background:{CARD_BG};color:{TEXT_DIM};border:1px solid {BORDER};"
            f"border-radius:3px;font-size:10px;padding:0 6px;")
        self.bg_swatch.setText("Right-click preview or here to set background")
        self.bg_swatch.setToolTip(
            "Right-click preview to sample\n"
            "Right-click here to reset")

        def _update_bg_swatch():
            if self._bg_rgb:
                r,g,b = self._bg_rgb
                lum = 0.299*r+0.587*g+0.114*b
                fg = "#000" if lum > 128 else "#fff"
                self.bg_swatch.setStyleSheet(
                    f"background:rgb({r},{g},{b});color:{fg};"
                    f"border:1px solid {BORDER};border-radius:3px;"
                    f"font-size:10px;padding:0 6px;")
                self.bg_swatch.setText(f" #{r:02x}{g:02x}{b:02x}  (right-click to reset)")
            else:
                self.bg_swatch.setStyleSheet(
                    f"background:{CARD_BG};color:{TEXT_DIM};border:1px solid {BORDER};"
                    f"border-radius:3px;font-size:10px;padding:0 6px;")
                self.bg_swatch.setText("🖱 Right-click preview or here to set background")

        def _pick_bg_color():
            from PyQt5.QtWidgets import QColorDialog
            init = QColor(*self._bg_rgb) if self._bg_rgb else QColor(220,220,220)
            col = QColorDialog.getColor(init, self, "Pick background color")
            if col.isValid():
                self._bg_rgb = (col.red(), col.green(), col.blue())
                _update_bg_swatch()

        def _clear_bg_color():
            self._bg_rgb = None
            _update_bg_swatch()

        def _bg_mouse(event):
            if event.button() == Qt.LeftButton:
                _pick_bg_color()
            elif event.button() == Qt.RightButton:
                _clear_bg_color()

        self.bg_swatch.mousePressEvent = _bg_mouse
        self.bg_clear_btn = type('_Dummy', (), {'setEnabled': lambda s,v: None})()
        bg_h.addWidget(self.bg_swatch, 1)

        # ── Shadow color swatch — same row as background ──────────────────────
        self._shadow_rgb = None
        self.shadow_swatch = QLabel("🌑 Left-click to set shadow")
        self.shadow_swatch.setFixedHeight(24)
        self.shadow_swatch.setAlignment(Qt.AlignCenter)
        self.shadow_swatch.setCursor(Qt.PointingHandCursor)
        self.shadow_swatch.setStyleSheet(
            f"background:{CARD_BG};color:{TEXT_DIM};border:1px solid {BORDER};"
            f"border-radius:3px;font-size:10px;padding:0 6px;")
        self.shadow_swatch.setToolTip(
            "Left-click preview to sample\n"
            "Right-click here to clear")

        def _update_shadow_swatch():
            if self._shadow_rgb:
                r,g,b = self._shadow_rgb
                lum = 0.299*r+0.587*g+0.114*b
                fg = "#000" if lum > 128 else "#fff"
                self.shadow_swatch.setStyleSheet(
                    f"background:rgb({r},{g},{b});color:{fg};"
                    f"border:1px solid {BORDER};border-radius:3px;"
                    f"font-size:10px;padding:0 6px;")
                self.shadow_swatch.setText(f"🌑 #{r:02x}{g:02x}{b:02x}  (right-click to reset)")
            else:
                self.shadow_swatch.setStyleSheet(
                    f"background:{CARD_BG};color:{TEXT_DIM};border:1px solid {BORDER};"
                    f"border-radius:3px;font-size:10px;padding:0 6px;")
                self.shadow_swatch.setText("🌑 Left-click to set shadow")

        def _shadow_mouse(event):
            if event.button() == Qt.LeftButton:
                from PyQt5.QtWidgets import QColorDialog
                init = QColor(*self._shadow_rgb) if self._shadow_rgb else QColor(80,80,80)
                col = QColorDialog.getColor(init, self, "Pick shadow color")
                if col.isValid():
                    self._shadow_rgb = (col.red(), col.green(), col.blue())
                    _update_shadow_swatch()
            elif event.button() == Qt.RightButton:
                self._shadow_rgb = None
                _update_shadow_swatch()

        self.shadow_swatch.mousePressEvent = _shadow_mouse
        bg_h.addWidget(self.shadow_swatch, 1)
        gl.addWidget(bg_sub, 13, 1)

        # iPhone photos folder (optional) — used for manual “recheck with latest iPhone photo”
        gl.addWidget(QLabel("iPhone photos:"), 14, 0)
        ipw = QWidget(); iph = QHBoxLayout(ipw); iph.setContentsMargins(0,0,0,0); iph.setSpacing(4)
        self._iphone_photo_dir = ""
        self.iphone_dir_lbl = QLabel("not set")
        self.iphone_dir_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        self.iphone_dir_lbl.setToolTip("Folder containing iPhone photos (newest file is used)")
        self.iphone_dir_btn = self._btn("Set…", CARD_BG, self._set_iphone_photo_dir, w=52)
        self.iphone_dir_btn.setFixedHeight(22)
        self.iphone_calib_btn = self._btn("Calib", CARD_BG, self._calibrate_iphone_from_selection, w=52)
        self.iphone_calib_btn.setFixedHeight(22)
        iph.addWidget(self.iphone_dir_lbl, 1)
        iph.addWidget(self.iphone_dir_btn)
        iph.addWidget(self.iphone_calib_btn)
        gl.addWidget(ipw, 14, 1)

        lv.addWidget(sg)

        # ── Preview ───────────────────────────────────────────────────────────
        pg = QGroupBox("📷 Live Preview")
        pl = QVBoxLayout(pg)
        self.preview_label = QLabel("Starting camera...")
        self.preview_label.setAlignment(Qt.AlignCenter)
        # Usability: allow preview to grow vertically; keep a reasonable minimum.
        self.preview_label.setMinimumHeight(220)
        self.preview_label.setMinimumWidth(200)
        self.preview_label.setStyleSheet(f"color:{TEXT_DIM};background:{CARD_BG};border-radius:6px;")
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_label.setCursor(Qt.PointingHandCursor)
        self.preview_label.setToolTip("Left-click → set shadow color\nRight-click → set background color")

        def _preview_dbl(event):
            src  = getattr(self, "_last_preview_source", None)
            bbox = getattr(self, "_last_preview_bbox",   None)
            if src:
                self._show_enlarged(src, bbox)
        self.preview_label.mouseDoubleClickEvent = _preview_dbl

        def _preview_click(event):
            pm = self.preview_label.pixmap()
            if pm is None or pm.isNull(): return
            lw = self.preview_label.width(); lh = self.preview_label.height()
            pw = pm.width(); ph = pm.height()
            ox = (lw - pw) // 2; oy = (lh - ph) // 2
            px = event.x() - ox;  py = event.y() - oy
            if px < 0 or py < 0 or px >= pw or py >= ph: return
            img = pm.toImage()
            col = img.pixelColor(px, py)
            r, g, b = col.red(), col.green(), col.blue()
            hex_col = f"#{r:02x}{g:02x}{b:02x}"
            lum = 0.299*r + 0.587*g + 0.114*b
            fg = "#000" if lum > 128 else "#fff"

            if event.button() == Qt.RightButton:
                # Right-click → Background color
                self._bg_rgb = (r, g, b)
                self.bg_swatch.setStyleSheet(
                    f"background:rgb({r},{g},{b});color:{fg};"
                    f"border:1px solid {BORDER};border-radius:3px;font-size:10px;padding:0 6px;")
                self.bg_swatch.setText(f" {hex_col}  (right-click to reset)")
                self._log(f"🖼  Background set to {hex_col}", "info")

            elif event.button() == Qt.LeftButton:
                # Left-click → Shadow color
                self._shadow_rgb = (r, g, b)
                self.shadow_swatch.setStyleSheet(
                    f"background:rgb({r},{g},{b});color:{fg};"
                    f"border:1px solid {BORDER};border-radius:3px;font-size:10px;padding:0 6px;")
                self.shadow_swatch.setText(f"🌑 {hex_col}  (right-click to reset)")
                self._log(f"🌑  Shadow set to {hex_col}", "info")

        self.preview_label.mousePressEvent = _preview_click
        pl.addWidget(self.preview_label)
        # Toggle live/last capture
        tog_h = QWidget(); tog_hl = QHBoxLayout(tog_h)
        tog_hl.setContentsMargins(0,0,0,0); tog_hl.setSpacing(4)
        self.live_btn = self._btn("● Live", SUCCESS, lambda: self._show_live(), w=87)
        self.snap_btn = self._btn("⊡ Last", CARD_BG, lambda: self._show_last_snap(), w=87)
        def _enlarge_preview():
            from PyQt5.QtWidgets import QDialog, QLabel, QVBoxLayout
            pm = self.preview_label.pixmap()
            if pm is None or pm.isNull(): return
            dlg = QDialog(self)
            dlg.setWindowTitle("Preview")
            dlg.setStyleSheet(f"background:{DARK_BG};")
            vl = QVBoxLayout(dlg); vl.setContentsMargins(4,4,4,4)
            lbl = QLabel(); lbl.setPixmap(pm.scaled(900, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            lbl.setAlignment(Qt.AlignCenter)
            vl.addWidget(lbl)
            dlg.exec_()
        self.enlarge_btn = self._btn("⛶", CARD_BG, _enlarge_preview, w=37)
        self.enlarge_btn.setToolTip("Enlarge preview")
        # NEW (2026): “HD scan” toggle button (Ctrl+H shortcut also available)
        self.hd_btn = self._btn("HD", CARD_BG, self._toggle_hd_scan, w=42)
        self.hd_btn.setFixedHeight(26)
        self.hd_btn.setToolTip("HD scan is OFF (default speed)\nClick to enable for next scans")
        tog_hl.addWidget(self.live_btn); tog_hl.addWidget(self.snap_btn); tog_hl.addWidget(self.hd_btn); tog_hl.addWidget(self.enlarge_btn); tog_hl.addStretch()
        pl.addWidget(tog_h)
        lv.addWidget(pg)

        # ── Weight counter ────────────────────────────────────────────────────

        lv.addStretch()

        # ── SCAN button ───────────────────────────────────────────────────────
        self.scan_btn = QPushButton("▶  SCAN")
        self.scan_btn.setFixedHeight(62)
        self.scan_btn.setStyleSheet(f"""
            QPushButton {{background:{ACCENT};color:white;font-size:22px;font-weight:bold;
                          border-radius:10px;border:none;}}
            QPushButton:hover {{background:#ff6b7a;}}
            QPushButton:disabled {{background:#444;color:#777;}}
        """)
        self.scan_btn.clicked.connect(self._scan_or_cancel)
        self._auto_merge_on = True
        self.auto_merge_btn = self._btn("🔀 Auto-merge ✓", "#1a4a1a", self._toggle_auto_merge)
        self.auto_merge_btn.setFixedHeight(62)
        self.auto_merge_btn.setToolTip("Auto-merge ON — click to disable")
        scan_row = QHBoxLayout(); scan_row.setSpacing(4)
        scan_row.addWidget(self.scan_btn, 4)
        scan_row.addWidget(self.auto_merge_btn, 1)
        lv.addLayout(scan_row)

        _H = 26  # button height for all 3 rows

        # Row 1: Image · Batch · Folder · Sessions
        row1 = QHBoxLayout(); row1.setSpacing(4)
        self.load_img_btn = self._btn("📂 Image",    CARD_BG, self._load_image_file)
        self.batch_btn    = self._btn("📂 Batch",    CARD_BG, self._load_batch_files)
        self.folder_btn   = self._btn("📁 Folder",   CARD_BG, self._load_scan_folder)
        self.sessions_btn = self._btn("🕐 Sessions", CARD_BG, self._show_sessions)
        self.load_img_btn.setToolTip("Load a single image to scan")
        self.batch_btn.setToolTip("Select multiple images — scanned sequentially")
        self.folder_btn.setToolTip("Select a folder — all images scanned sequentially")
        self.sessions_btn.setToolTip("Reopen a recent scan session")
        for b in [self.load_img_btn, self.batch_btn, self.folder_btn, self.sessions_btn]:
            b.setFixedHeight(_H); row1.addWidget(b)
        lv.addLayout(row1)

        # Row 2: XML · Copy · Report · Merge Dupes · BrickOwl
        row2 = QHBoxLayout(); row2.setSpacing(4)
        self.copy_btn = self._btn("💾 XML",    ACCENT2, self._download_xml)
        self.clip_btn = self._btn("📋 Copy",   CARD_BG, self._copy_xml_to_clipboard)
        self.html_btn = self._btn("🌐 Report", CARD_BG, self._open_html)
        self.merge_r2 = self._btn("🔀 Merge",  CARD_BG, self._merge_duplicate_lots)
        self.bo_btn   = self._btn("🦉 BrickOwl", CARD_BG, self._export_brickowl_csv)
        self.copy_btn.setToolTip("Save XML file")
        self.clip_btn.setToolTip("Copy XML to clipboard")
        self.merge_r2.setToolTip("Merge rows with same part+color into one lot")
        self.bo_btn.setToolTip("Export BrickOwl batch upload CSV")
        for b in [self.copy_btn, self.clip_btn, self.html_btn, self.merge_r2, self.bo_btn]:
            b.setFixedHeight(_H); row2.addWidget(b)
        lv.addLayout(row2)

        # Row 3: Capture · BrickStore · Merge · Upload
        row3 = QHBoxLayout(); row3.setSpacing(4)
        self.capture_btn = self._btn("📸 Capture",     "#1a3a1a", self._start_capture_countdown)
        self.bs_btn      = self._btn("🧱 BrickStore",  CARD_BG,   self._open_in_brickstore)
        self.merge_btn   = self._btn("🔀 Merge",       CARD_BG,   self._merge_duplicate_lots)
        self.merge_btn.setToolTip("Merge rows with same part+color into one lot")
        self.upload_btn  = self._btn("🚀 Upload",      "#1a3a1a", self._upload_to_bricklink)
        self.capture_btn.setToolTip("Capture frame from live feed (3s countdown)")
        self.merge_btn.setToolTip("Merge rows with same part+color into one lot")
        self.upload_btn.setToolTip("Upload inventory directly to BrickLink via API")
        for b in [self.capture_btn, self.merge_btn, self.upload_btn, self.bs_btn]:
            b.setFixedHeight(_H); row3.addWidget(b)
        lv.addLayout(row3)
        self.scan_btn.setText("▶  SCAN")
        if hasattr(self, "cancel_btn"): self.cancel_btn.setEnabled(False)
        self.scan_btn.setStyleSheet("""
            QPushButton { background: #1a3a1a; color: white; font-size: 22px;
            font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background: #2a5a2a; }
        """)

        for b in [self.copy_btn, self.clip_btn, self.html_btn,
                  self.merge_r2, self.merge_btn, self.upload_btn, self.bs_btn, self.bo_btn]:
            b.setEnabled(False)

        main_splitter.addWidget(left_scroll)

        # ── RIGHT PANEL ───────────────────────────────────────────────────────
        right = QWidget()
        rv = QVBoxLayout(right); rv.setContentsMargins(0,0,0,0); rv.setSpacing(8)

        splitter = QSplitter(Qt.Vertical)

        # Results list
        res_w = QWidget(); res_w.setObjectName("resPanel")
        res_w.setStyleSheet(f"QWidget#resPanel {{ background:{PANEL_BG}; border-radius:4px; border:1px solid {BORDER}; }}")
        res_l = QVBoxLayout(res_w); res_l.setContentsMargins(8,8,8,8); res_l.setSpacing(4)

        rh = QHBoxLayout()
        rt = QLabel("Scan Results"); rt.setStyleSheet(f"font-size:13px;font-weight:bold;color:{TEXT_DIM};")
        self.results_count = QLabel("0 parts"); self.results_count.setStyleSheet(f"font-size:11px;color:{ACCENT2};")
        self.clear_results_btn = self._btn("Clear",CARD_BG,self._clear_results,w=55)
        self.del_errors_btn = self._btn("✗ Delete errors", ACCENT, self._delete_error_rows, w=110)
        rh.addWidget(rt); rh.addWidget(self.results_count); rh.addStretch()
        self._results_fullscreen = False
        self.fullscreen_btn = self._btn("⛶", CARD_BG, self._toggle_results_fullscreen, w=44)
        self.fullscreen_btn.setToolTip("Toggle fullscreen results panel")
        rh.addWidget(self.del_errors_btn); rh.addWidget(self.clear_results_btn); rh.addWidget(self.fullscreen_btn)
        res_l.addLayout(rh)

        # Bulk action bar — hidden until rows are selected
        self.bulk_bar = QWidget()
        self.bulk_bar.setObjectName("bulkBar")
        self.bulk_bar.setAutoFillBackground(True)
        _bpal = self.bulk_bar.palette()
        _bpal.setColor(self.bulk_bar.backgroundRole(), QColor("#2a2010"))
        self.bulk_bar.setPalette(_bpal)
        self.bulk_bar.setStyleSheet(f"QWidget#bulkBar {{ background:#2a2010; border:1px solid {ACCENT2}40; }}")
        bbl = QHBoxLayout(self.bulk_bar); bbl.setContentsMargins(8,4,8,4); bbl.setSpacing(6)
        self.bulk_label = QLabel("bulk")
        self.bulk_label.setStyleSheet(f"font-size:11px;color:{ACCENT2};font-weight:bold;")
        bbl.addWidget(self.bulk_label)
        bbl.addStretch()
        _bh = 26   # button height
        self._bulk_sel_all  = self._btn("☑ All",    CARD_BG, self._select_all_rows,          w=78)
        self._bulk_sel_none = self._btn("☐ None",   CARD_BG, self._select_no_rows,           w=78)
        self._bulk_set_qty  = self._btn("✏ Qty",    CARD_BG, self._bulk_edit_qty,            w=78)
        self._bulk_set_med  = self._btn("💲 Med",    CARD_BG, self._bulk_set_medium,          w=78)
        self._bulk_set_med15= self._btn("💲 +15%",  CARD_BG, self._bulk_set_medium15,        w=86)
        self._bulk_min01    = self._btn("⬇ Min", CARD_BG, self._bulk_set_min_01_all,     w=80)
        self._bulk_price    = self._btn("✏ Price",  CARD_BG, self._bulk_set_price,           w=86)
        self._bulk_color    = self._btn("🎨 Color",    CARD_BG, self._bulk_override_color,   w=97)
        self._bulk_bg_color = self._btn("🔬 BG Color", CARD_BG, self._bulk_apply_scan_color, w=97)
        self._bulk_recolor  = self._btn("🎯 Recolor",  CARD_BG, self._bulk_pixel_recolor,    w=97)
        self._bulk_del      = self._btn("🗑 Del",    ACCENT,  self._bulk_delete,              w=78)
        self._bulk_used     = self._btn("U Used",   "#2a1800", lambda: self._bulk_set_condition("U"), w=78)
        self._bulk_new      = self._btn("N New",    "#1a3a1a", lambda: self._bulk_set_condition("N"), w=73)
        self._bulk_remark   = self._btn("✏ Remark", CARD_BG, self._bulk_set_remark,           w=97)
        self._bulk_undo = self._btn("↩ Undo", "#1a1a2e", self._undo_last, w=78)
        self._bulk_undo.setStyleSheet("background:#1a1a2e;color:#7a9fff;font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #2a3a6e;")
        self._bulk_undo.setToolTip("Undo last change (one level)")
        # Merge button removed to keep bulk bar compact and avoid splitter getting stuck
        self._bulk_color.setToolTip("Set color for all selected rows")
        self._bulk_bg_color.setToolTip("Restore Brickognize scan color for selected rows\n(click again to toggle back to manual color)")
        self._bulk_recolor.setToolTip("Re-sample pixel color from crop image — ignores Brickognize,\nmatches against full BrickLink color table by Lab ΔE")
        self._bulk_price.setToolTip("Set a fixed price for all selected rows")
        self._bulk_min01.setToolTip("Set a minimum price floor for ALL rows\n(only applies to unpriced or lower-than-floor rows)")
        self._bulk_used.setStyleSheet(f"background:#2a1800;color:{ACCENT2};font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #4a3000;")
        self._bulk_new.setStyleSheet(f"background:#1a3a1a;color:#5fca7a;font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #2a5a2a;")
        _all_bulk_btns = [self._bulk_sel_all, self._bulk_sel_none, self._bulk_set_qty,
                          self._bulk_set_med, self._bulk_set_med15, self._bulk_min01, self._bulk_price,
                          self._bulk_color, self._bulk_bg_color, self._bulk_recolor,
                          self._bulk_used, self._bulk_new,
                          self._bulk_remark, self._bulk_del, self._bulk_undo]
        for b in _all_bulk_btns:
            b.setFixedHeight(_bh)
            bbl.addWidget(b)
        self.bulk_bar.setFixedHeight(34)
        res_l.addWidget(self.bulk_bar)

        # Column headers — clickable to sort
        self._sort_col = None
        self._sort_asc = True
        self._selected = set()  # set of row indices currently selected
        self._UNDO_MAX = 25
        self._undo_stack = []  # stack of snapshots (last item is most recent)
        self._last_clicked_rd = None   # row dict anchor for shift-click range
        # bulk_bar always visible — dims and disables when nothing selected
        hdr = QWidget()
        hdr.setObjectName("resultsHeader")
        hdr.setAutoFillBackground(True)
        _hpal = hdr.palette()
        _hpal.setColor(hdr.backgroundRole(), QColor("#1a1a1a"))
        hdr.setPalette(_hpal)
        hdr.setStyleSheet(f"QWidget#resultsHeader {{ background:#1a1a1a; border-bottom:2px solid {ACCENT2}; }}")
        hdrl = QHBoxLayout(hdr); hdrl.setContentsMargins(8,3,8,3); hdrl.setSpacing(6)
        self._hdr_btns = {}
        # (label, width, sort_key)
        # Select-all checkbox in header
        self._sel_all_chk = QCheckBox()
        self._sel_all_chk.setTristate(False)
        self._sel_all_chk.setStyleSheet("margin-left:4px;")
        self._sel_all_chk.setToolTip("Select / deselect all")
        self._sel_all_chk.stateChanged.connect(lambda s: self._select_all_rows() if s == Qt.Checked else self._select_no_rows())

        cols = [("#",24,None),("Crop",110,None),("BL",110,None),
                ("Part ID",90,"part_id"),("Name",185,"part_name"),
                ("Color",115,"color_name"),("Conf",48,"confidence"),
                ("Price",58,"price"),("Qty",38,"qty"),("Remark",80,"remark"),("",30,None)]
        self._sel_all_chk.setFixedWidth(18)
        hdrl.addWidget(self._sel_all_chk)
        for lbl, w, key in cols:
            if key:
                btn = QPushButton(lbl)
                btn.setFixedWidth(w)
                btn.setFlat(True)
                btn.setStyleSheet(f"font-size:10px;font-weight:bold;color:{TEXT_DIM};text-align:left;padding:0 4px;background:transparent;border:none;letter-spacing:0.5px;")
                btn.setCursor(Qt.PointingHandCursor)
                def make_sort(k): return lambda: self._sort_results(k)
                btn.clicked.connect(make_sort(key))
                hdrl.addWidget(btn)
                self._hdr_btns[key] = btn
            else:
                lb = QLabel(lbl); lb.setFixedWidth(w)
                lb.setStyleSheet(f"font-size:10px;font-weight:bold;color:{TEXT_DIM};")
                hdrl.addWidget(lb)
        hdrl.addStretch()
        res_l.addWidget(hdr)

        self.results_scroll = QScrollArea(); self.results_scroll.setWidgetResizable(True)
        self.results_container = QWidget()
        self.results_container.setObjectName("resultsContainer")
        # Use palette, not stylesheet, so Qt doesn't cascade background to child rows
        _rc_pal = self.results_container.palette()
        _rc_pal.setColor(self.results_container.backgroundRole(), QColor(PANEL_BG))
        self.results_container.setPalette(_rc_pal)
        self.results_container.setAutoFillBackground(True)
        self.results_list_layout = QVBoxLayout(self.results_container)

        self.results_list_layout.setSpacing(2); self.results_list_layout.setContentsMargins(0,0,0,0)
        self.results_list_layout.addStretch()
        self.results_scroll.setWidget(self.results_container)
        # Drag-select: press+drag across rows to select multiple at once
        self._drag_selecting  = False
        self._drag_start_rd   = None
        self.results_container.setMouseTracking(True)
        self.results_scroll.viewport().setMouseTracking(True)
        res_l.addWidget(self.results_scroll)
        splitter.addWidget(res_w)

        # ── Price Guide Panel (BrickStore-style) ──────────────────────────────
        # ── Price Guide Panel (BrickStore-style: Used + New side by side) ────────
        self._pg_panel_widget = QWidget()
        self._pg_panel_widget.setObjectName("priceGuidePanel")
        self._pg_panel_widget.setStyleSheet(
            f"QWidget#priceGuidePanel {{ background:{PANEL_BG}; border-radius:4px; border:1px solid {BORDER}; }}")
        pg_l = QVBoxLayout(self._pg_panel_widget)
        pg_l.setContentsMargins(8, 5, 8, 5); pg_l.setSpacing(3)

        # Header row: title | part label | stretch | Sold/Stock toggle | close button
        pg_hdr = QHBoxLayout(); pg_hdr.setSpacing(5)
        pg_title = QLabel("Price Guide")
        pg_title.setStyleSheet(f"font-size:11px;font-weight:bold;color:{TEXT_DIM};")
        self.pg_part_label = QLabel("—")
        self.pg_part_label.setStyleSheet(f"font-size:11px;color:{ACCENT2};font-weight:bold;")
        self.pg_toggle_sold  = self._btn("Sold",  ACCENT2, lambda: self._pg_set_guide("sold"),  w=57)
        self.pg_toggle_stock = self._btn("Stock", CARD_BG, lambda: self._pg_set_guide("stock"), w=59)
        self.pg_toggle_sold.setFixedHeight(20); self.pg_toggle_stock.setFixedHeight(20)
        pg_close_btn = self._btn("✕", "#5a2020", self._pg_close, w=28)
        pg_close_btn.setFixedHeight(20)
        pg_close_btn.setToolTip("Hide price guide (click any row to reopen)")
        pg_hdr.addWidget(pg_title)
        pg_hdr.addWidget(self.pg_part_label)
        pg_hdr.addStretch()
        pg_hdr.addWidget(self.pg_toggle_sold)
        pg_hdr.addWidget(self.pg_toggle_stock)
        pg_hdr.addWidget(pg_close_btn)
        pg_l.addLayout(pg_hdr)

        # Column headers: blank | Used (col 1) | New (col 2)
        pg_grid = QGridLayout(); pg_grid.setSpacing(2); pg_grid.setContentsMargins(0,1,0,1)
        pg_grid.setColumnStretch(0, 0); pg_grid.setColumnStretch(1, 1); pg_grid.setColumnStretch(2, 1)
        _hdr_style = f"font-size:10px;font-weight:bold;color:{ACCENT2};text-align:center;"
        _used_hdr = QLabel("Used"); _used_hdr.setStyleSheet(_hdr_style); _used_hdr.setAlignment(Qt.AlignCenter)
        _new_hdr  = QLabel("New");  _new_hdr.setStyleSheet(_hdr_style);  _new_hdr.setAlignment(Qt.AlignCenter)
        pg_grid.addWidget(QLabel(""), 0, 0)
        pg_grid.addWidget(_used_hdr, 0, 1)
        pg_grid.addWidget(_new_hdr,  0, 2)

        _lbl_style = f"font-size:10px;color:{TEXT_DIM};"
        _val_style = f"font-size:10px;color:{TEXT};font-weight:bold;"

        def _pg_row2(row, label):
            lbl  = QLabel(label); lbl.setStyleSheet(_lbl_style)
            used = QLabel("—");   used.setStyleSheet(_val_style); used.setAlignment(Qt.AlignCenter)
            new  = QLabel("—");   new.setStyleSheet(_val_style);  new.setAlignment(Qt.AlignCenter)
            pg_grid.addWidget(lbl,  row, 0)
            pg_grid.addWidget(used, row, 1)
            pg_grid.addWidget(new,  row, 2)
            return used, new

        self.pg_min_u,     self.pg_min_n     = _pg_row2(1, "Min")
        self.pg_avg_u,     self.pg_avg_n     = _pg_row2(2, "Avg")
        self.pg_qty_avg_u, self.pg_qty_avg_n = _pg_row2(3, "Qty Avg")
        self.pg_max_u,     self.pg_max_n     = _pg_row2(4, "Max")
        self.pg_lots_u,    self.pg_lots_n    = _pg_row2(5, "Lots")
        self.pg_units_u,   self.pg_units_n   = _pg_row2(6, "Units")
        pg_l.addLayout(pg_grid)

        # Apply buttons row
        pg_apply = QHBoxLayout(); pg_apply.setSpacing(4)
        self.pg_apply_btn = self._btn("↓ Apply Used avg to selected", CARD_BG, self._pg_apply_avg)
        self.pg_apply_btn.setFixedHeight(22)
        self.pg_apply_btn.setEnabled(False)
        self.pg_apply_med_btn = self._btn("↓ Apply to all rows", CARD_BG, self._pg_apply_to_all)
        self.pg_apply_med_btn.setFixedHeight(22)
        self.pg_apply_med_btn.setEnabled(False)
        pg_apply.addWidget(self.pg_apply_btn)
        pg_apply.addWidget(self.pg_apply_med_btn)
        pg_l.addLayout(pg_apply)

        # Keep backward-compat aliases so _pg_set_condition callers still work
        self.pg_toggle_new  = None   # no longer needed — both columns always visible
        self.pg_toggle_used = None

        splitter.addWidget(self._pg_panel_widget)
        self._pg_splitter = splitter   # keep ref for open/close
        self._pg_condition = "U"   # U=used, N=new
        self._pg_guide     = "sold" # sold / stock
        self._pg_current_rd = None  # row dict currently shown in price guide

        # Console
        con_w = QWidget()
        con_l = QVBoxLayout(con_w); con_l.setContentsMargins(0,0,0,0); con_l.setSpacing(4)
        con_h = QHBoxLayout()
        con_t = QLabel("Console"); con_t.setStyleSheet(f"font-size:12px;font-weight:bold;color:{TEXT_DIM};")
        self.clear_btn = self._btn("Clear", CARD_BG, self._clear_console, w=50)
        self._con_close_btn = self._btn("✕", "#5a2020", lambda: self._pg_splitter.setSizes(
            [self._pg_splitter.sizes()[0]+self._pg_splitter.sizes()[1]+self._pg_splitter.sizes()[2],
             0, 0, self._pg_splitter.sizes()[3]]), w=22)
        self._con_close_btn.setToolTip("Hide console")
        con_h.addWidget(con_t); con_h.addStretch()
        con_h.addWidget(self.clear_btn); con_h.addWidget(self._con_close_btn)
        con_l.addLayout(con_h)
        self.console = QTextEdit(); self.console.setReadOnly(True)
        self.console.setLineWrapMode(QTextEdit.WidgetWidth)
        con_l.addWidget(self.console)
        splitter.addWidget(con_w)

        # ── Weight Calculator Panel ──────────────────────────────────────────
        wt_w = QWidget()
        wt_w.setStyleSheet(f"background:{PANEL_BG};border-radius:4px;border:1px solid {BORDER};")
        wt_l = QVBoxLayout(wt_w); wt_l.setContentsMargins(8,6,8,6); wt_l.setSpacing(5)

        # Header
        wt_hdr = QHBoxLayout()
        wt_title = QLabel("⚖ Weight → Qty")
        wt_title.setStyleSheet(f"font-size:11px;font-weight:bold;color:{TEXT_DIM};")
        self._wt_status = QLabel("No scale")
        self._wt_status.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        wt_con_btn = self._btn("Connect", CARD_BG, self._wt_connect_scale, w=84)
        wt_con_btn.setFixedHeight(20)
        self._wt_con_btn = wt_con_btn
        wt_hdr.addWidget(wt_title); wt_hdr.addWidget(self._wt_status)
        wt_hdr.addStretch(); wt_hdr.addWidget(wt_con_btn)
        wt_l.addLayout(wt_hdr)

        # Live weight display
        wt_row1 = QHBoxLayout()
        self._wt_live = QLabel("— g")
        self._wt_live.setStyleSheet(f"font-size:18px;font-weight:bold;color:{ACCENT2};min-width:80px;")
        wt_tare_btn = self._btn("Tare", CARD_BG, self._wt_tare, w=48)
        wt_tare_btn.setFixedHeight(22)
        wt_row1.addWidget(self._wt_live)
        wt_row1.addStretch()
        wt_row1.addWidget(wt_tare_btn)
        wt_l.addLayout(wt_row1)

        # Part selector + unit weight
        wt_row2 = QHBoxLayout()
        self.weight_part_id = QComboBox()
        self.weight_part_id.setStyleSheet(f"background:{CARD_BG};color:{TEXT};font-size:10px;border:1px solid {BORDER};")
        self.weight_part_id.setFixedHeight(22)
        self.weight_part_id.setToolTip("Part to count — populated after scan")
        self.weight_part_id.addItem("— select part —", "")
        self.weight_part_id.currentIndexChanged.connect(self._wt_part_selected)
        wt_row2.addWidget(self.weight_part_id, 3)
        self._wt_unit_lbl = QLabel("unit: —")
        self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        wt_row2.addWidget(self._wt_unit_lbl, 1)
        wt_l.addLayout(wt_row2)

        # Manual part ID entry
        wt_pid_row = QHBoxLayout()
        wt_pid_lbl = QLabel("Part ID:")
        wt_pid_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        self._wt_pid_edit = QLineEdit()
        self._wt_pid_edit.setFixedHeight(20)
        self._wt_pid_edit.setPlaceholderText("e.g. 3004")
        self._wt_pid_edit.setStyleSheet(f"background:{CARD_BG};color:{TEXT};font-size:10px;border:1px solid {BORDER};padding:2px;")
        self._wt_pid_edit.setToolTip("Enter BrickLink part number manually")
        wt_pid_row.addWidget(wt_pid_lbl); wt_pid_row.addWidget(self._wt_pid_edit)
        wt_l.addLayout(wt_pid_row)

        # Result
        wt_row3 = QHBoxLayout()
        self._wt_result = QLabel("Qty: —")
        self._wt_result.setStyleSheet(f"font-size:13px;font-weight:bold;color:{SUCCESS};")
        wt_apply_btn = self._btn("→ Apply qty", CARD_BG, self._wt_apply_qty, w=80)
        wt_apply_btn.setFixedHeight(22)
        self._wt_apply_btn = wt_apply_btn
        self._wt_apply_btn.setEnabled(False)
        wt_row3.addWidget(self._wt_result)
        wt_row3.addStretch()
        wt_row3.addWidget(wt_apply_btn)
        wt_l.addLayout(wt_row3)

        # Manual weight entry fallback
        wt_row4 = QHBoxLayout()
        wt_manual_lbl = QLabel("Manual g:")
        wt_manual_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        self._wt_manual = QLineEdit()
        self._wt_manual.setFixedHeight(20)
        self._wt_manual.setPlaceholderText("enter grams")
        self._wt_manual.setStyleSheet(f"background:{CARD_BG};color:{TEXT};font-size:10px;border:1px solid {BORDER};padding:2px;")
        self._wt_manual.textChanged.connect(self._wt_manual_changed)
        self._wt_pid_edit.textChanged.connect(lambda _: self._wt_recalculate_from_inputs())
        wt_row4.addWidget(wt_manual_lbl); wt_row4.addWidget(self._wt_manual)
        wt_l.addLayout(wt_row4)

        splitter.addWidget(wt_w)
        self._wt_tare_val = 0.0
        self._wt_raw_g    = 0.0
        self._wt_serial   = None
        self._wt_thread   = None

        # Allow panels to collapse to zero when handle is dragged
        splitter.setCollapsible(0, False)   # results — never collapse
        splitter.setCollapsible(1, True)    # price guide — collapsible
        splitter.setCollapsible(2, True)    # console — collapsible
        splitter.setCollapsible(3, True)    # weight — collapsible
        splitter.setStretchFactor(0, 1)     # results absorbs extra space
        splitter.setStretchFactor(1, 0)
        splitter.setStretchFactor(2, 0)
        splitter.setStretchFactor(3, 0)
        # Usability: responsive initial panel widths (doesn't fight user after first layout).
        # Goal: give most space to Results; keep Price Guide/Console readable on wide screens.
        def _init_right_splitter_sizes():
            try:
                total = max(400, splitter.width())
                # Default targets as ratios (approx old 480/180/100/0)
                r0 = int(total * 0.65)   # results
                r1 = int(total * 0.23)   # price guide
                r2 = max(0, total - r0 - r1)  # console (remainder)
                # Keep weight panel collapsed by default
                splitter.setSizes([r0, r1, r2, 0])
            except Exception:
                splitter.setSizes([480, 180, 100, 0])
        QTimer.singleShot(0, _init_right_splitter_sizes)

        rv.addWidget(splitter)
        status_row = QHBoxLayout(); status_row.setContentsMargins(0,2,0,2)
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet(f"font-size:11px;color:{TEXT_DIM};")
        self.bl_store_label = QLabel("  🏪 Store: —  ")
        self.bl_store_label.setStyleSheet(
            f"font-size:11px;color:{TEXT_DIM};font-weight:bold;padding:0 8px;"
            f"background:#1a1a1a;border-radius:4px;border:1px solid {BORDER};min-width:160px;")
        self.bl_store_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.bl_store_label.setFixedHeight(22)
        self.bl_store_label.setCursor(Qt.PointingHandCursor)
        self.bl_store_label.setToolTip("BrickLink store total — click to refresh")
        self.bl_store_label.mousePressEvent = lambda e: self._fetch_bl_store_total()
        self.total_value_label = QLabel("")
        self.total_value_label.setStyleSheet(
            f"font-size:11px;color:{ACCENT2};font-weight:bold;padding:0 8px;"
            f"background:#1a1a1a;border-radius:4px;border:1px solid {BORDER};"
            f"min-width:320px;")
        self.total_value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.total_value_label.setFixedHeight(22)
        status_row.addWidget(self.status_label)
        status_row.addStretch()
        status_row.addWidget(self.bl_store_label)
        status_row.addSpacing(6)
        status_row.addWidget(self.total_value_label)
        rv.addLayout(status_row)
        QTimer.singleShot(3000, self._fetch_bl_store_total)

        main_splitter.addWidget(right)
        # Usability: responsive initial left/right split.
        # Keeps left panel at ~30% (clamped) so the results area is roomy by default.
        def _init_main_splitter_sizes():
            try:
                total = max(600, main_splitter.width() or self.width())
                left_w = int(total * 0.30)
                left_w = max(260, min(left_w, 520))
                main_splitter.setSizes([left_w, max(200, total - left_w)])
            except Exception:
                main_splitter.setSizes([440, 1100])
        QTimer.singleShot(0, _init_main_splitter_sizes)
        root.addWidget(main_splitter)
        self._log("✓  Scan Station ready.", "success")

    # ── Helper: small button ──────────────────────────────────────────────────
    def _btn(self, text, bg, cb, w=None):
        b = QPushButton(text)
        b.setFixedHeight(34)
        if w: b.setFixedWidth(w)
        b.setStyleSheet(f"""
            QPushButton {{background:{bg};color:{TEXT};font-size:12px;font-weight:bold;
                          border-radius:6px;border:1px solid #333;padding:0 10px;}}
            QPushButton:hover {{background:#3a3a6a;}}
            QPushButton:disabled {{color:#555;}}
        """)
        b.clicked.connect(cb)
        return b

    # ── Camera detection ──────────────────────────────────────────────────────
    def _detect_cameras(self):
        # Bump generation to kill the live thread, then probe in background.
        # The probe thread waits for the device to be released before opening it.
        self._stop_live_camera()  # sets _live_running=False, bumps _cam_generation
        self._detecting = True    # guard against re-entrant detection

        self.camera_combo.blockSignals(True)
        self.camera_combo.clear()
        self.camera_combo.addItem("Detecting...", -1)
        self._log("🔍  Scanning for local cameras...", "info")

        def probe():
            import time as _time
            # Wait for live camera thread to release the device before probing.
            # _live_running goes False quickly, but cap.release() takes a moment.
            _time.sleep(0.5)
            found = []
            try:
                import cv2, concurrent.futures, subprocess

                # ── Pull friendly device names from Windows (PowerShell/WMI) ──
                device_names = {}
                try:
                    ps = (
                        'Get-CimInstance Win32_PnPEntity | '
                        'Where-Object { $_.PNPClass -eq "Camera" -or $_.PNPClass -eq "Image" } | '
                        'Select-Object -ExpandProperty Name'
                    )
                    r = subprocess.run(
                        ["powershell", "-NoProfile", "-Command", ps],
                        capture_output=True, text=True, timeout=5, **_hidden_popen_kwargs())
                    names = [l.strip() for l in r.stdout.splitlines() if l.strip()]
                    for k, name in enumerate(names):
                        device_names[k] = name
                except Exception:
                    pass

                def try_open(i, backend):
                    try:
                        cap = cv2.VideoCapture(i, backend)
                        if cap.isOpened():
                            ret, _ = cap.read()   # confirm device is truly alive
                            w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                            h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                            cap.release()
                            if ret and w > 0:
                                return (i, w, h)
                        cap.release()
                    except Exception:
                        pass
                    return None

                # Probe all indices in parallel — much faster than sequential
                with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
                    dshow_futs = {ex.submit(try_open, i, cv2.CAP_DSHOW): i for i in range(6)}
                    dshow_results = {}
                    for fut, i in dshow_futs.items():
                        try:
                            r = fut.result(timeout=3)
                            if r: dshow_results[i] = r
                        except (concurrent.futures.TimeoutError, Exception):
                            pass

                for i in range(6):
                    if i in dshow_results:
                        result, bname = dshow_results[i], "DSHOW"
                    else:
                        # MSMF fallback only for indices that had no DSHOW result
                        try:
                            result = try_open(i, cv2.CAP_MSMF)
                            bname = "MSMF" if result else ""
                        except Exception:
                            result, bname = None, ""
                    if result and not any(f[0] == i for f in found):
                        dname = device_names.get(i, "")
                        found.append((result[0], result[1], result[2], bname, dname))

            except ImportError:
                pass
            self.cameras_detected.emit(found)

        threading.Thread(target=probe, daemon=True).start()

    def _on_stream_died(self):
        """Called on main thread when the IP stream thread exits.
        If cameras already probed → start webcam immediately.
        Otherwise → run detection first (which auto-starts on completion)."""
        self._use_http_stream = False
        self._stream_active   = False
        if getattr(self, "_detecting", False):
            return  # detection already in progress — don't interfere
        if getattr(self, "_switching_to_webcam", False):
            return  # _switch_to_webcam already called _start_live_camera — don't double-start
        if self.camera_combo.count() > 0 and self.camera_combo.itemData(0) != -1:
            # Cameras already known — go straight to webcam
            QTimer.singleShot(200, self._start_live_camera)
        else:
            # No cameras probed yet — detect first; _cameras_found will start live
            self._detect_cameras()

    # NOTE: legacy USB iPhone capture (idevicescreenshot) has been removed.
    # If needed again in the future, implement it via the new native iPhone
    # capture path instead of libimobiledevice CLI tools.

    def _switch_to_webcam(self):
        """Stop IP stream and switch to local webcam."""
        self._use_http_stream = False
        self._switching_to_webcam = True  # tells _on_stream_died not to interfere
        self._log("📷  Switching to webcam...", "info")
        self._stop_live_camera()
        # Clear flag after a short delay — long enough for _stream_died signal
        # to be delivered and processed, so _on_stream_died sees it and skips.
        QTimer.singleShot(500, lambda: setattr(self, "_switching_to_webcam", False))
        QTimer.singleShot(300, self._start_live_camera)

    def _auto_start_stream(self):
        """On startup, auto-connect HTTP stream if URL is pre-filled, else start webcam."""
        self._live_paused = False          # always start in Live mode
        url = self.stream_url.currentText().strip()
        if url and not url.startswith("—") and not self._stream_active:
            self._toggle_stream()
        elif not self._stream_active:
            self._start_live_camera()

    def _toggle_stream(self):
        """Switch live preview to DroidCam HTTP stream."""
        url_text = self.stream_url.currentText().strip()
        if not url_text or url_text.startswith("—"):
            self._use_http_stream = False
            self._log("📷  Switched to webcam mode", "info")
            self._start_live_camera()
            return
        # Normalize URL
        if not url_text.startswith("http"):
            # Only add /video if no path specified
            if "/" not in url_text.split(":")[-1]:
                url_text = f"http://{url_text}/video"
            else:
                url_text = f"http://{url_text}"
        self._http_stream_url = url_text
        self._use_http_stream = True
        self._log(f"📷  Connecting to {url_text}...", "info")
        self._stop_live_camera()
        threading.Thread(target=self._stream_http, daemon=True).start()

    def _stream_http(self):
        """Read MJPEG stream from DroidCam HTTP endpoint."""
        with self._stream_lock:
            if self._stream_active:
                return  # already streaming
            self._stream_active = True
        import urllib.request, time, cv2, numpy as np
        url = self._http_stream_url
        self._live_running = True
        try:
            stream = urllib.request.urlopen(url, timeout=3)
            self.log_message.emit("✓  HTTP stream connected", "success")
            buf = b""
            while self._live_running:
                buf += stream.read(4096)
                a  = buf.find(b'\xff\xd8')
                b_ = buf.find(b'\xff\xd9')
                if a != -1 and b_ != -1 and b_ > a:
                    jpg = buf[a:b_+2]
                    buf = buf[b_+2:]
                    arr   = np.frombuffer(jpg, dtype=np.uint8)
                    frame = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                    if frame is not None:
                        self._last_frame = frame
                        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                        h2, w2 = rgb.shape[:2]
                        try:
                            self.camera_frame.emit(w2, h2, rgb.tobytes())
                        except RuntimeError:
                            break  # window deleted — exit gracefully
        except Exception as e:
            err = str(e)
            if "10061" in err or "refused" in err.lower():
                self.log_message.emit("⚠  IP stream lost — switching to local camera...", "warning")
            else:
                self.log_message.emit(f"⚠  Stream error: {e} — switching to local camera...", "warning")
        finally:
            self._live_running    = False
            self._stream_active   = False
            self._use_http_stream = False
            self.log_message.emit("📷  Switching to local camera...", "info")
            self._live_paused = False
            # Emit signal — safely triggers _detect_cameras on the main thread
            self._stream_died.emit()

    def _get_active_cam_idx(self):
        idx = self.camera_combo.currentData()
        return idx if idx is not None and idx >= 0 else 0

    def _test_camera(self):
        """Cycle through all camera indices, stream each for 3 seconds so user can see the feed."""
        self._log("🎥  Scanning all cameras — watch the preview...", "info")
        self._stop_live_camera()

        def run():
            import cv2, time
            for idx in range(8):
                for backend in [cv2.CAP_MSMF, cv2.CAP_DSHOW]:
                    try:
                        cap = cv2.VideoCapture(idx, backend)
                        if cap.isOpened():
                            # Warmup — try up to 2 seconds for frame
                            ret, frame = False, None
                            for _ in range(20):
                                ret, frame = cap.read()
                                if ret and frame is not None and frame.size > 0:
                                    break
                                time.sleep(0.1)
                            if ret and frame is not None and frame.size > 0:
                                self.log_message.emit(
                                    f"📷  Camera {idx} — is this your feed? (showing 3 sec...)", "info")
                                end = time.time() + 3
                                while time.time() < end and not self._live_running:
                                    ret, frame = cap.read()
                                    if ret and frame is not None and frame.size > 0:
                                        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                                        h2, w2 = frame_rgb.shape[:2]
                                        try:
                                            self.camera_frame.emit(w2, h2, frame_rgb.tobytes())
                                        except RuntimeError:
                                            break  # window deleted
                                    time.sleep(0.067)
                                cap.release()
                                break
                            cap.release()
                    except Exception:
                        pass
            self.log_message.emit("✓  Done scanning — select the right camera and calibrate", "success")

        threading.Thread(target=run, daemon=True).start()

    def _start_live_camera(self):
        """Stop existing stream then start fresh — supports both webcam and HTTP stream.
        Uses generation counter so only the latest thread ever runs the frame loop.
        """
        self._live_running = False
        self._live_paused  = False   # always show live when camera starts
        self._update_preview_toggle()
        import time

        # If HTTP stream active, use that instead
        if self._use_http_stream and self._http_stream_url:
            if not self._stream_active:
                time.sleep(0.8)  # wait for any webcam thread to fully exit
                threading.Thread(target=self._stream_http, daemon=True).start()
            return

        # Bump generation — any older delayed_start thread will see stale gen and exit
        self._cam_generation += 1
        my_gen = self._cam_generation

        def delayed_start():
            time.sleep(0.3)
            if self._cam_generation != my_gen:
                return  # superseded by a newer call

            # Load crop from config if available
            crop = None
            cam_w, cam_h = 1920, 1080
            if CFG_PATH.exists():
                try:
                    cfg      = json.loads(CFG_PATH.read_text())
                    cam_idx  = self._get_active_cam_idx()
                    profiles = cfg.get("cameras", {})
                    profile  = profiles.get(str(cam_idx), {})
                    crop     = profile.get("crop") or cfg.get("crop")
                    cam_w    = profile.get("camera_width")  or cfg.get("camera_width",  1920)
                    cam_h    = profile.get("camera_height") or cfg.get("camera_height", 1080)
                except Exception:
                    pass

            # Camera index: always use the combo box (user has explicitly selected it).
            # station.cfg camera is only a startup hint — after detection the combo is authoritative.
            cam_idx = self._get_active_cam_idx()

            if self._cam_generation != my_gen:
                return  # superseded before opening device

            try:
                import cv2
                cap = None
                known = getattr(self, "_cam_backends", {}).get(cam_idx)
                backends = ([known, cv2.CAP_DSHOW, cv2.CAP_MSMF] if known
                            else [cv2.CAP_DSHOW, cv2.CAP_MSMF])
                for backend in dict.fromkeys(backends):  # deduplicate, preserve order
                    try:
                        c = cv2.VideoCapture(cam_idx, backend)
                        if c.isOpened():
                            cap = c; break
                        c.release()
                    except Exception:
                        pass

                if cap is None:
                    self.log_message.emit(f"⚠  Could not open camera {cam_idx}", "warning")
                    return

                cap.set(cv2.CAP_PROP_FRAME_WIDTH,  cam_w)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, cam_h)
                cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
                self._live_running = True

                # Warmup — try up to 2 full attempts with different backends
                self.log_message.emit(f"📷  Connecting to camera {cam_idx}...", "info")
                got_frame = False
                for attempt in range(2):
                    if attempt > 0:
                        # First attempt failed — release and retry with alternate backend
                        cap.release()
                        time.sleep(0.5)
                        if self._cam_generation != my_gen:
                            return
                        alt_backend = cv2.CAP_MSMF if known == cv2.CAP_DSHOW else cv2.CAP_DSHOW
                        cap = cv2.VideoCapture(cam_idx, alt_backend)
                        if not cap.isOpened():
                            break
                        cap.set(cv2.CAP_PROP_FRAME_WIDTH,  cam_w)
                        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, cam_h)
                        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
                        self.log_message.emit(f"📷  Retrying camera {cam_idx} with alt backend...", "info")
                    for _ in range(40):
                        if self._cam_generation != my_gen:
                            cap.release(); return
                        ret, frame = cap.read()
                        if ret and frame is not None and frame.size > 0:
                            got_frame = True; break
                        time.sleep(0.1)
                    if got_frame:
                        break

                if not got_frame:
                    self.log_message.emit(f"⚠  Camera {cam_idx} — no frame after 2 attempts", "warning")
                    cap.release(); return

                self.log_message.emit(f"📷  Live preview — camera {cam_idx}", "success")

                while self._live_running and self._cam_generation == my_gen:
                    ret, frame = cap.read()
                    if not ret or frame is None or frame.size == 0:
                        time.sleep(0.05)
                        continue
                    # Store full uncropped frame — ScanWorker applies _live_crop at capture time
                    self._last_frame = frame
                    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    h2, w2 = rgb.shape[:2]
                    try:
                        self.camera_frame.emit(w2, h2, rgb.tobytes())
                    except RuntimeError:
                        break  # window deleted — exit gracefully
                    time.sleep(0.067)

            except Exception as e:
                self.log_message.emit(f"⚠  Camera error: {e}", "warning")
            finally:
                try: cap.release()
                except Exception: pass  # ignorable
                self._live_cap = None
                if self._cam_generation == my_gen:
                    self.log_message.emit("📷  Camera released", "info")

        threading.Thread(target=delayed_start, daemon=True).start()

    def _stop_live_camera(self):
        self._live_running  = False
        self._stream_active = False
        self._cam_generation += 1  # invalidate any running delayed_start thread
        # Small yield so background threads notice the flag before we restart
        import time; time.sleep(0.05)

    def _update_live_preview(self, w, h, data):
        if getattr(self, "_live_paused", False):
            return
        if not hasattr(self, '_frame_count'): self._frame_count = 0
        self._frame_count += 1
        if self._frame_count == 1:
            self._log("🎥  First frame received", "success")
            self._live_paused = False
            self._update_preview_toggle()
        try:
            img = QImage(data, w, h, w * 3, QImage.Format_RGB888).copy()
            # Crop to calibrated area if available
            crop = getattr(self, "_live_crop", None)
            if crop:
                x1, y1, x2, y2 = crop
                # Clamp to actual frame dimensions
                x1 = max(0, min(x1, w-1)); x2 = max(x1+1, min(x2, w))
                y1 = max(0, min(y1, h-1)); y2 = max(y1+1, min(y2, h))
                img = img.copy(x1, y1, x2-x1, y2-y1)
            lw  = max(self.preview_label.width(),  300)
            lh  = max(self.preview_label.height(), 220)
            pix = QPixmap.fromImage(img).scaled(lw, lh, Qt.KeepAspectRatio, Qt.SmoothTransformation)

            # Grid overlay — draw when Fixed Grid Mode is checked
            if getattr(self, "grid_check", None) and self.grid_check.isChecked():
                cols = self.cols_spin.value()
                rows = self.rows_spin.value()
                from PyQt5.QtGui import QPainter, QPen, QFont as QF
                from PyQt5.QtCore import Qt as _Qt
                painter = QPainter(pix)
                painter.setRenderHint(QPainter.Antialiasing, False)
                # Semi-transparent white lines
                pen = QPen(QColor(255, 200, 0, 200))
                pen.setWidth(1)
                painter.setPen(pen)
                pw, ph = pix.width(), pix.height()
                cell_w = pw / cols
                cell_h = ph / rows
                for c in range(1, cols):
                    x = int(c * cell_w)
                    painter.drawLine(x, 0, x, ph)
                for r in range(1, rows):
                    y = int(r * cell_h)
                    painter.drawLine(0, y, pw, y)
                # Draw outer border
                pen2 = QPen(QColor(255, 200, 0, 160))
                pen2.setWidth(2)
                painter.setPen(pen2)
                painter.drawRect(0, 0, pw-1, ph-1)
                # Cell numbers
                painter.setPen(QPen(QColor(255, 220, 0, 220)))
                font = QF(); font.setPointSize(7); font.setBold(True)
                painter.setFont(font)
                for r in range(rows):
                    for c in range(cols):
                        n = r * cols + c + 1
                        x = int(c * cell_w) + 3
                        y = int(r * cell_h) + 10
                        painter.drawText(x, y, str(n))
                painter.end()

            self.preview_label.setPixmap(pix)
            self.preview_label.setText("")
        except Exception as e:
            self._log(f"Preview error: {e}", "error")

    def _cameras_found(self, found):
        self.camera_combo.blockSignals(True)
        self.camera_combo.clear()
        # Also repopulate camera2_combo
        if hasattr(self, "camera2_combo"):
            self.camera2_combo.blockSignals(True)
            self.camera2_combo.clear()
            self.camera2_combo.addItem("— select 2nd cam —", -1)
        if found:
            self._cam_backends = {}  # idx -> cv2.CAP_DSHOW or cv2.CAP_MSMF
            import cv2 as _cv2
            for item in found:
                idx   = item[0]; w = item[1]; h = item[2]
                bname = item[3] if len(item) > 3 else ""
                dname = item[4] if len(item) > 4 else ""
                lbl   = f"{dname or f'Camera {idx}'}  ({w}x{h})"
                if bname: lbl += f"  [{bname}]"
                self.camera_combo.addItem(lbl, idx)
                if hasattr(self, "camera2_combo"):
                    self.camera2_combo.addItem(lbl, idx)
                self._cam_backends[idx] = (_cv2.CAP_DSHOW if bname == "DSHOW"
                                           else _cv2.CAP_MSMF)
            if CFG_PATH.exists():
                try:
                    saved = json.loads(CFG_PATH.read_text()).get("active_camera",
                            json.loads(CFG_PATH.read_text()).get("camera", 0))
                    for i in range(self.camera_combo.count()):
                        if self.camera_combo.itemData(i) == saved:
                            self.camera_combo.setCurrentIndex(i); break
                except Exception as e:
                    self._log(f"⚠  Camera combo restore: {e}", "warning")
            stream_active = self._stream_active
            if stream_active:
                self._log(f"✓  Found {len(found)} local camera(s) — IP stream active", "info")
            else:
                self._log(f"✓  Found {len(found)} camera(s)", "success")
        else:
            for i in range(3):
                self.camera_combo.addItem(f"Camera {i}", i)
                if hasattr(self, "camera2_combo"):
                    self.camera2_combo.addItem(f"Camera {i}", i)
            self._log("⚠  Could not detect cameras — select manually", "warning")
        self.camera_combo.blockSignals(False)
        if hasattr(self, "camera2_combo"):
            self.camera2_combo.blockSignals(False)
        # Connect save signal only once
        if not getattr(self, "_cam_combo_connected", False):
            self.camera_combo.currentIndexChanged.connect(self._save_camera)
            # Only restart camera when user manually changes the combo
            self.camera_combo.activated.connect(lambda _: self._start_live_camera())
            self._cam_combo_connected = True
        # Start webcam only if stream is not already running
        self._detecting = False  # detection complete — re-entrant guard off
        if not self._stream_active:
            self._live_paused = False
            QTimer.singleShot(200, self._start_live_camera)

    def _resume_after_calibration(self):
        """After calibration window closes, reload config and restart live preview."""
        self._live_paused = False          # always go back to Live, never stay on Last
        self._update_preview_toggle()
        if self._http_stream_url:
            # Was using HTTP stream — restore flag and reconnect
            self._use_http_stream = True
            self._toggle_stream()
        else:
            idx = self.camera_combo.currentData()
            if idx is not None and idx >= 0:
                self._log(f"📷  Resuming live camera {idx} after calibration", "success")
            self._start_live_camera()

    def _run_calibrate(self):
        self._log("📐  Launching calibration...", "info")
        self._log("    Close the calibration window when done.", "info")

        # Save current frame to temp file BEFORE stopping camera.
        # calibrate-station uses --image so it never needs to open the camera at all.
        import time, tempfile
        frame_path = ""
        if self._last_frame is not None:
            try:
                import cv2 as _cv2
                tmp = Path(tempfile.gettempdir()) / "scanstation_calib_frame.jpg"
                _cv2.imwrite(str(tmp), self._last_frame, [_cv2.IMWRITE_JPEG_QUALITY, 95])
                frame_path = str(tmp)
                self._log(f"    Frame saved for calibration", "info")
            except Exception as e:
                self._log(f"⚠  Could not save frame: {e}", "warning")
        else:
            # No live frame (e.g. folder/batch image) — fall back to current preview image if available
            p = getattr(self, "_last_preview_path", None)
            try:
                if p and Path(p).exists():
                    frame_path = str(Path(p))
                    self._log("    Using current preview image for calibration", "info")
            except Exception:
                pass

        self._stop_live_camera()
        cam_idx = self.camera_combo.currentData()
        if cam_idx is None or cam_idx < 0:
            cam_idx = 0

        def _launch():
            time.sleep(0.3)
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"

            if frame_path:
                cmd = [_python_exe(), "calibrate-station.py",
                       "--image", frame_path,
                       "--camera", str(cam_idx)]
                self.log_message.emit(f"    Using saved frame — camera {cam_idx}", "info")
            else:
                # No frame — open camera directly (fallback)
                cmd = [_python_exe(), "calibrate-station.py", "--camera", str(cam_idx)]
                self.log_message.emit(f"    No frame available — opening camera {cam_idx}", "info")

            try:
                proc = subprocess.Popen(cmd, env=env, cwd=str(Path.cwd()),
                                        **_hidden_popen_kwargs())
            except Exception as e:
                self.log_message.emit(f"⚠  Could not launch calibration: {e}", "error")
                return

            proc.wait()
            self.log_message.emit("📐  Calibration closed — config updated.", "success")
            self.calibration_done.emit()

        threading.Thread(target=_launch, daemon=True).start()

    def _save_camera(self):
        idx = self.camera_combo.currentData()
        if idx is None or idx < 0: return
        if CFG_PATH.exists():
            try:
                cfg = json.loads(CFG_PATH.read_text())
                cfg["camera"] = idx
                CFG_PATH.write_text(json.dumps(cfg, indent=2))
                self._log(f"📷  Camera {idx} saved", "info")
            except Exception as e:
                self._log(f"⚠  Could not save camera config: {e}", "warning")

    # ── Config load ───────────────────────────────────────────────────────────
    def _load_config(self):
        if CFG_PATH.exists():
            try:
                cfg = json.loads(CFG_PATH.read_text())
                cw = cfg.get("crop_width","?"); ch = cfg.get("crop_height","?")
                cam = cfg.get("camera",0)
                self.cfg_label.setText(f"✓  Camera {cam} — crop {cw}×{ch}px")
                self.cfg_label.setStyleSheet(
                    f"font-size:10px;color:{SUCCESS};padding:4px 8px;background:#0a1f0a;border-radius:3px;border:1px solid #1a3a1a;")
                # Load crop rect for live preview cropping
                crop = cfg.get("crop")
                if crop and len(crop) == 4:
                    self._live_crop = tuple(int(v) for v in crop)
                else:
                    self._live_crop = None

                # iPhone photos folder (optional)
                iph = cfg.get("iphone_photo_dir", "")
                if isinstance(iph, str) and iph:
                    self._iphone_photo_dir = iph
                if hasattr(self, "iphone_dir_lbl"):
                    self._update_iphone_dir_label()

                # iPhone offset (alignment between DroidCam and iPhone crops, in iPhone px)
                off = cfg.get("iphone_offset", None)
                if isinstance(off, list) and len(off) == 2:
                    try:
                        self._iphone_dx = int(off[0])
                        self._iphone_dy = int(off[1])
                    except Exception:
                        self._iphone_dx = self._iphone_dy = 0

                # iPhone photos folder (optional)
                iph = cfg.get("iphone_photo_dir", "")
                if isinstance(iph, str) and iph:
                    self._iphone_photo_dir = iph
                if hasattr(self, "iphone_dir_lbl"):
                    self._update_iphone_dir_label()

                # NEW (2026): restore last splitter sizes (if present).
                # This runs after _build_ui, so splitters exist and won't fight user resizing.
                ms = cfg.get("ui_main_splitter_sizes")
                ps = cfg.get("ui_right_splitter_sizes")
                def _apply():
                    try:
                        if ms and isinstance(ms, list) and len(ms) == 2 and hasattr(self, "_main_splitter"):
                            self._main_splitter.setSizes([int(ms[0]), int(ms[1])])
                        if ps and isinstance(ps, list) and len(ps) == 4 and hasattr(self, "_pg_splitter"):
                            self._pg_splitter.setSizes([int(ps[0]), int(ps[1]), int(ps[2]), int(ps[3])])
                    except Exception:
                        pass
                QTimer.singleShot(0, _apply)
            except Exception as e:
                self._log(f"⚠  Could not load station.cfg: {e}", "warning")

    def _wire_splitter_persistence(self):
        """Hook splitterMoved → debounced save into station.cfg."""
        try:
            if hasattr(self, "_main_splitter"):
                self._main_splitter.splitterMoved.connect(lambda *_: self._ui_save_timer.start(350))
            if hasattr(self, "_pg_splitter"):
                self._pg_splitter.splitterMoved.connect(lambda *_: self._ui_save_timer.start(350))
        except Exception:
            pass

    def _save_ui_layout(self):
        """Persist current splitter sizes into station.cfg (keeps other config intact)."""
        try:
            cfg = {}
            if CFG_PATH.exists():
                try:
                    cfg = json.loads(CFG_PATH.read_text())
                except Exception:
                    cfg = {}
            if hasattr(self, "_main_splitter"):
                cfg["ui_main_splitter_sizes"] = [int(x) for x in self._main_splitter.sizes()[:2]]
            if hasattr(self, "_pg_splitter"):
                sz = self._pg_splitter.sizes()
                if len(sz) >= 4:
                    cfg["ui_right_splitter_sizes"] = [int(sz[0]), int(sz[1]), int(sz[2]), int(sz[3])]
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception:
            # Silent: layout persistence should never crash the scanner.
            pass

    def _update_iphone_dir_label(self):
        try:
            d = getattr(self, "_iphone_photo_dir", "") or ""
            if not d:
                self.iphone_dir_lbl.setText("not set")
                self.iphone_dir_lbl.setToolTip("Folder containing iPhone photos (newest file is used)")
                return
            short = d
            if len(short) > 38:
                short = "…" + short[-37:]
            self.iphone_dir_lbl.setText(short)
            self.iphone_dir_lbl.setToolTip(d)
        except Exception:
            pass

    def _set_iphone_photo_dir(self):
        from PyQt5.QtWidgets import QFileDialog
        start = getattr(self, "_iphone_photo_dir", "") or str(Path.cwd())
        folder = QFileDialog.getExistingDirectory(self, "Select iPhone photos folder", start)
        if not folder:
            return
        self._iphone_photo_dir = folder
        self._update_iphone_dir_label()
        # Persist
        try:
            cfg = {}
            if CFG_PATH.exists():
                try:
                    cfg = json.loads(CFG_PATH.read_text())
                except Exception:
                    cfg = {}
            cfg["iphone_photo_dir"] = folder
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception:
            pass

    def _get_latest_iphone_photo(self) -> str:
        """Return newest JPG/PNG path in iPhone photo dir, or ''."""
        d = getattr(self, "_iphone_photo_dir", "") or ""
        if not d or not Path(d).exists():
            return ""
        exts = {".jpg", ".jpeg", ".png", ".webp", ".heic"}
        try:
            files = [p for p in Path(d).iterdir() if p.is_file() and p.suffix.lower() in exts]
            if not files:
                return ""
            files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return str(files[0])
        except Exception:
            return ""

    def _save_iphone_offset(self):
        """Persist current iPhone offset into station.cfg."""
        try:
            cfg = {}
            if CFG_PATH.exists():
                try:
                    cfg = json.loads(CFG_PATH.read_text())
                except Exception:
                    cfg = {}
            cfg["iphone_offset"] = [int(self._iphone_dx), int(self._iphone_dy)]
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception:
            pass

    def _calibrate_iphone_from_selection(self):
        """User-facing entry: calibrate iPhone offset using a selected row and latest photo."""
        rows = self._selected_rows()
        if not rows:
            self._log("📱  Select a row first, then click Calib to align with iPhone photo", "warning")
            return
        rd = rows[0]
        src = rd.get("source_image", "")
        bbox = rd.get("bbox")
        if not src or not bbox or not Path(src).exists():
            self._log("📱  Cannot calibrate — selected row has no source/bbox", "warning")
            return
        iph_path = self._get_latest_iphone_photo()
        if not iph_path:
            self._log("📱  Cannot calibrate — no iPhone photos found (set folder first)", "warning")
            return
        self._calibrate_iphone_offset(src, iph_path, bbox)

    def _calibrate_iphone_offset(self, src, iph_path, bbox):
        """Interactive calibration: nudge iPhone crop so it matches DroidCam crop."""
        try:
            from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel as _QL, QPushButton
            from PyQt5.QtGui import QPixmap as _QP
            from PyQt5.QtCore import Qt as _Qt
            from PIL import Image as _PIL
            dlg = QDialog(self)
            dlg.setWindowTitle("Calibrate iPhone offset")
            layout = QVBoxLayout(dlg)
            imgs = QHBoxLayout()
            layout.addLayout(imgs)
            scan_lbl = _QL(); iphone_lbl = _QL()
            scan_lbl.setAlignment(_Qt.AlignCenter); iphone_lbl.setAlignment(_Qt.AlignCenter)
            imgs.addWidget(scan_lbl, 1); imgs.addWidget(iphone_lbl, 1)
            btns = QHBoxLayout()
            layout.addLayout(btns)
            for name in ["↑","↓","←","→","OK","Cancel"]:
                b = QPushButton(name)
                btns.addWidget(b)
                if name=="OK":
                    b.clicked.connect(dlg.accept)
                elif name=="Cancel":
                    b.clicked.connect(dlg.reject)
            x1,y1,x2,y2 = [int(v) for v in bbox]
            # Prepare static scan crop
            try:
                scan_img = _PIL.open(src).convert("RGB")
                sx = scan_img.crop((x1,y1,x2,y2))
                tmp = Path.cwd() / "_tmp_scan_crop_preview.jpg"
                sx.save(str(tmp), "JPEG", quality=85)
                scan_lbl.setPixmap(_QP(str(tmp)).scaled(220,220,_Qt.KeepAspectRatio,_Qt.SmoothTransformation))
            except Exception:
                pass
            iph_img = _PIL.open(iph_path).convert("RGB")
            wd,hd = scan_img.size
            wi,hi = iph_img.size
            # Apply existing rotation heuristic
            if (wd > hd and wi < hi) or (wd < hd and wi > hi):
                iph_img = iph_img.rotate(90, expand=True)
                wi,hi = iph_img.size
            sx_scale = wi/float(wd); sy_scale = hi/float(hd)
            dx = self._iphone_dx; dy = self._iphone_dy
            def _update_iphone():
                ix1 = max(0, min(int(x1 * sx_scale + dx), wi - 2))
                ix2 = max(ix1 + 1, min(int(x2 * sx_scale + dx), wi))
                iy1 = max(0, min(int(y1 * sy_scale + dy), hi - 2))
                iy2 = max(iy1 + 1, min(int(y2 * sy_scale + dy), hi))
                c = iph_img.crop((ix1,iy1,ix2,iy2))
                tmp2 = Path.cwd() / "_tmp_iphone_crop_preview.jpg"
                c.save(str(tmp2), "JPEG", quality=85)
                iphone_lbl.setPixmap(_QP(str(tmp2)).scaled(220,220,_Qt.KeepAspectRatio,_Qt.SmoothTransformation))
            _update_iphone()
            def _mk_nudge(dx0,dy0):
                def _do():
                    nonlocal dx,dy
                    dx += dx0; dy += dy0
                    _update_iphone()
                return _do
            for b in dlg.findChildren(QPushButton):
                if b.text()=="↑": b.clicked.connect(_mk_nudge(0,-5))
                elif b.text()=="↓": b.clicked.connect(_mk_nudge(0,5))
                elif b.text()=="←": b.clicked.connect(_mk_nudge(-5,0))
                elif b.text()=="→": b.clicked.connect(_mk_nudge(5,0))
            if dlg.exec_()==QDialog.Accepted:
                self._iphone_dx = int(dx); self._iphone_dy = int(dy)
                self._save_iphone_offset()
        except Exception:
            pass

    def _iphone_recheck_row(self, rd):
        """User-triggered: recheck ONE row using newest iPhone photo (no time limit)."""
        src = rd.get("source_image", "")
        bbox = rd.get("bbox")
        if not src or not bbox or not Path(src).exists():
            self._log("📱  iPhone recheck: missing source image/bbox", "warning")
            return
        iph_path = self._get_latest_iphone_photo()
        if not iph_path:
            self._log("📱  iPhone recheck: no photos found (set iPhone photos folder first)", "warning")
            return
        self._log(f"📱  Rechecking with latest iPhone photo: {Path(iph_path).name}", "info")
        self._iphone_recheck_row_fixed_iph(rd, iph_path)

    def _iphone_recheck_selected(self, rows):
        """Recheck many rows using newest iPhone photo (batched, no time limit)."""
        rows = [r for r in (rows or []) if r and not r.get("_deleted")]
        if not rows:
            return
        iph_path = self._get_latest_iphone_photo()
        if not iph_path:
            self._log("📱  iPhone recheck: no photos found (set iPhone photos folder first)", "warning")
            return
        usable = []
        for rd in rows:
            src = rd.get("source_image", "")
            bbox = rd.get("bbox")
            if src and bbox and Path(src).exists():
                usable.append(rd)
        if not usable:
            self._log("📱  iPhone recheck: selected rows missing source image/bbox", "warning")
            return
        self._save_undo_snapshot(f"iphone recheck ({len(usable)})")
        self._log(f"📱  Rechecking {len(usable)} selected row(s) with latest iPhone photo: {Path(iph_path).name}", "info")
        import queue
        q = queue.Queue()
        for r in usable:
            q.put(r)
        def _worker():
            while True:
                try:
                    rd = q.get_nowait()
                except Exception:
                    break
                try:
                    self._iphone_recheck_row_fixed_iph(rd, iph_path)
                finally:
                    try:
                        q.task_done()
                    except Exception:
                        pass
        n_workers = min(3, max(1, len(usable)))
        for _ in range(n_workers):
            threading.Thread(target=_worker, daemon=True).start()

    def _iphone_recheck_row_fixed_iph(self, rd, iph_path: str):
        """Internal: recheck ONE row using a provided iphone image path."""
        src = rd.get("source_image", "")
        bbox = rd.get("bbox")
        if not src or not bbox or not Path(src).exists():
            return
        if not iph_path or not Path(iph_path).exists():
            return
        def _work():
            try:
                from PIL import Image as _PIL
                import io, requests
                scan_img = _PIL.open(src).convert("RGB")
                iph_img  = _PIL.open(iph_path).convert("RGB")
                # Normalize orientation: if rotated 90° relative to scan, rotate left.
                wd, hd = scan_img.size
                wi, hi = iph_img.size
                if (wd > hd and wi < hi) or (wd < hd and wi > hi):
                    iph_img = iph_img.rotate(90, expand=True)
                    wi, hi = iph_img.size
                wd, hd = scan_img.size
                x1, y1, x2, y2 = [int(v) for v in bbox]
                sx = wi / float(wd)
                sy = hi / float(hd)
                dx = getattr(self, "_iphone_dx", 0)
                dy = getattr(self, "_iphone_dy", 0)
                # Apply scale + calibrated offset (offset is in iPhone pixels).
                ix1 = max(0, min(int(x1 * sx + dx), wi - 2))
                ix2 = max(ix1 + 1, min(int(x2 * sx + dx), wi))
                iy1 = max(0, min(int(y1 * sy + dy), hi - 2))
                iy2 = max(iy1 + 1, min(int(y2 * sy + dy), hi))
                crop = iph_img.crop((ix1, iy1, ix2, iy2))
                # Overwrite or create crop_image so GUI can show the iPhone crop.
                try:
                    crop_path = rd.get("crop_image", "")
                    if crop_path:
                        out_p = Path(crop_path)
                    else:
                        base = Path(src).parent if src else Path(iph_path).parent
                        out_p = base / f"iphone_crop_{rd.get('index',0):03d}.jpg"
                        rd["crop_image"] = str(out_p)
                    out_p.parent.mkdir(parents=True, exist_ok=True)
                    crop.save(str(out_p), "JPEG", quality=92)
                    # Always ask main thread to refresh thumbnail, even if Brickognize finds no match.
                    self._iphone_crop_refresh.emit(rd)
                except Exception:
                    pass
                buf = io.BytesIO()
                crop.save(buf, "JPEG", quality=90)
                buf.seek(0)
                resp = requests.post(
                    "https://api.brickognize.com/predict/",
                    files={"query_image": ("iphone_crop.jpg", buf.getvalue(), "image/jpeg")},
                    headers={"accept": "application/json"},
                    timeout=30,
                )
                if resp.status_code != 200:
                    self.log_message.emit(f"📱  iPhone recheck error — Brickognize HTTP {resp.status_code}", "error")
                    return
                data = resp.json() or {}
                items = data.get("items") or []
                if not items:
                    self.log_message.emit("📱  iPhone recheck: no match (Brickognize returned no items)", "info")
                    return
                best = items[0]
                best["_alts"] = items[1:6]
                self._iphone_recheck_ready.emit(rd, best)
            except Exception:
                # Keep quiet in batch; main log already indicates batch start
                pass
        threading.Thread(target=_work, daemon=True).start()

    def _on_iphone_recheck_ready(self, rd, best):
        """Main thread: apply Brickognize result from iPhone crop to a row."""
        try:
            new_pid = best.get("id", "")
            new_name = best.get("name", "") or "Unknown"
            new_conf = float(best.get("score", 0) or 0)
            new_type = best.get("type", "P")
            if not new_pid:
                return
            self._save_undo_snapshot("iphone recheck")
            rd["part_id"] = new_pid
            rd["part_name"] = new_name
            rd["item_type"] = self._effective_item_type({"part_id": new_pid, "item_type": new_type})
            rd["confidence"] = new_conf
            rd["thumb_url"] = best.get("img_url", rd.get("thumb_url", ""))
            # Replace alternatives with new Brickognize candidates
            alts = []
            for c in (best.get("_alts") or []):
                if not c.get("id"):
                    continue
                alts.append({
                    "id": c.get("id", ""),
                    "name": c.get("name", ""),
                    "score": c.get("score", 0),
                    "type": c.get("type", "P"),
                    "img_url": c.get("img_url", ""),
                })
            rd["alternatives"] = alts
            # UI refresh
            if rd.get("_pid_lbl"):
                rd["_pid_lbl"].setText(("👤 " if rd["item_type"] == "M" else "") + new_pid)
            if rd.get("_name_lbl"):
                label = new_name[:38] + "…" if len(new_name) > 38 else new_name
                rd["_name_lbl"].setText(label)
                rd["_name_lbl"].setToolTip(new_name)
            if rd.get("_confl"):
                cc = SUCCESS if new_conf >= 0.7 else WARNING if new_conf >= 0.4 else ACCENT
                rd["_confl"].setText(f"{new_conf:.0%}")
                rd["_confl"].setStyleSheet(f"color:{cc};font-weight:bold;font-size:12px;")
            # Refresh crop thumbnail if we have one and crop_image now points to iPhone crop
            try:
                cl = rd.get("_crop_lbl")
                cp = rd.get("crop_image", "")
                if cl is not None and cp and Path(cp).exists():
                    from PyQt5.QtGui import QPixmap as _QP
                    px = _QP(cp).scaled(110,110,Qt.KeepAspectRatio,Qt.SmoothTransformation)
                    cl.setPixmap(px)
                    cl.setText("")
            except Exception:
                pass
            # Reload BL reference image
            bl_lbl = rd.get("_bl_lbl")
            if bl_lbl:
                bl_lbl.setPixmap(QPixmap())
                bl_lbl.setText("…")
                bl_lbl._itype = rd.get("item_type","P")
                pid_for_img = rd.get("part_id","")
                threading.Thread(
                    target=self._load_bl_img_for,
                    args=(pid_for_img, rd.get("color_id", 0), rd.get("item_type","P"), bl_lbl, rd.get("thumb_url","")),
                    daemon=True,
                ).start()
            # Refresh pricing
            self._fetch_price_for_row(rd)
            self._fetch_price_guide(rd)
            self._log(f"📱  Row updated from iPhone crop: {new_pid} ({new_conf:.0%})", "success")
        except Exception as e:
            self._log(f"📱  Apply iPhone result failed: {e}", "warning")

    def _on_iphone_crop_refresh(self, rd):
        """Main thread: refresh crop thumbnail from rd['crop_image'] (even if no match)."""
        try:
            cl = rd.get("_crop_lbl")
            cp = rd.get("crop_image", "")
            if cl is None or not cp or not Path(cp).exists():
                return
            from PyQt5.QtGui import QPixmap as _QP
            px = _QP(cp).scaled(110,110,Qt.KeepAspectRatio,Qt.SmoothTransformation)
            cl.setPixmap(px)
            cl.setText("")
        except Exception:
            pass

    # ── Scale detection ───────────────────────────────────────────────────────
    # ── Console ───────────────────────────────────────────────────────────────
    def _log(self, msg, level="info"):
        colors = {"info":TEXT,"success":SUCCESS,"warning":WARNING,"error":ACCENT}
        self.console.setTextColor(QColor(colors.get(level, TEXT)))
        self.console.append(msg)
        self.console.moveCursor(QTextCursor.End)

    def _clear_console(self):
        self.console.clear()

    # ── Results list ──────────────────────────────────────────────────────────
    def _on_price_ready(self, rd, price):
        """Main-thread slot — safely update price label after background fetch."""
        if price is not None and price > 0:
            rd["price"]        = price
            rd["medium_price"] = price
            rd["_price_lbl"].setText(f"${price:.2f}")
            rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;")
        else:
            rd["_price_lbl"].setText("—")
            rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
        if hasattr(self, "total_value_label"):
            self._update_total_value()

    def _load_bl_img_for(self, pid, color_id, itype, lbl, thumb_url=""):
        """Background thread: fetch BL reference image and emit signal to update label."""
        try:
            import urllib.request
            from PyQt5.QtGui import QPixmap as _QP
            from PyQt5.QtCore import QByteArray as _QBA
            lbl._itype = itype
            pid_lo = (pid or "").strip().lower()
            pid_up = pid_lo.upper()

            if itype == "M":
                # BL minifig image pattern: ItemImage/MN/0/{id}.png
                # Try both cases; put BL URLs first — Brickognize thumbs are often
                # generic placeholders for minifigs and pollute the result.
                urls = [
                    f"https://img.bricklink.com/ItemImage/MN/0/{pid_lo}.png",
                    f"https://img.bricklink.com/ItemImage/MN/0/{pid_up}.png",
                    f"https://img.bricklink.com/ItemImage/MF/0/{pid_lo}.png",
                    f"https://img.bricklink.com/ItemImage/MF/0/{pid_up}.png",
                ]
                if thumb_url:
                    urls.append(thumb_url)   # Brickognize thumb as last resort
            else:
                # Parts: color-specific first, then color 0, then Brickognize thumb
                urls = [
                    f"https://img.bricklink.com/ItemImage/PN/{color_id}/{pid_lo}.png",
                    f"https://img.bricklink.com/ItemImage/PN/0/{pid_lo}.png",
                ]
                if thumb_url:
                    urls.append(thumb_url)

            for url in urls:
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                    resp = urllib.request.urlopen(req, timeout=8)
                    data = resp.read()
                    # Reject suspiciously tiny responses (BL sometimes returns a
                    # 1×1 transparent PNG ~68 bytes instead of a proper 404)
                    if len(data) < 200:
                        continue
                    pm = _QP()
                    pm.loadFromData(_QBA(data))
                    if pm.isNull() or pm.width() < 4:
                        continue
                    pm = pm.scaled(110, 110, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self._bl_img_signal.emit(lbl, pm)
                    return
                except Exception:
                    continue
            # All BrickLink URLs failed — try Rebrickable inventory_parts.csv
            rb_url = rb_img_url(pid_lo, int(color_id))
            if rb_url:
                # Check disk cache first
                cached = _rb_cached_path(pid_lo, color_id)
                if cached.exists():
                    data = cached.read_bytes()
                else:
                    try:
                        req2 = urllib.request.Request(rb_url, headers={"User-Agent": "Mozilla/5.0"})
                        resp2 = urllib.request.urlopen(req2, timeout=8)
                        data = resp2.read()
                        if len(data) > 200:
                            cached.write_bytes(data)
                        else:
                            data = b""
                    except Exception:
                        data = b""
                if data:
                    pm2 = _QP()
                    pm2.loadFromData(_QBA(data))
                    if not pm2.isNull() and pm2.width() > 4:
                        pm2 = pm2.scaled(110, 110, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        self._bl_img_signal.emit(lbl, pm2)
                        return
            self._bl_img_signal.emit(lbl, _QP())
        except Exception:
            try:
                from PyQt5.QtGui import QPixmap as _QP2
                self._bl_img_signal.emit(lbl, _QP2())
            except Exception:
                pass

    def _on_bl_img(self, lbl, pm):
        """Thread-safe slot to update BL reference image label — runs on main thread."""
        try:
            # QLabel may already be deleted if rows were cleared while the image was loading.
            from sip import isdeleted as _isdel
            if lbl is None or _isdel(lbl):
                return
            if not pm.isNull():
                lbl.setPixmap(pm)
                lbl.setText("")
            else:
                lbl.setText("🧍" if getattr(lbl, "_itype", "P") == "M" else "?")
                lbl.setStyleSheet(lbl.styleSheet() + "color:#888;font-size:16px;")
        except Exception:
            # Swallow any errors here; BL image updates should never crash the app.
            pass

    # BrickLink minifig catalog prefixes — whole figures listed under Minifigs,
    # not Parts.  Brickognize returns type="P" for all of these; we must override.
    _MINIFIG_CATALOG_PREFIXES = (
    "ac",  # Atlantis
    "adv",  # Adventurers
    "agt",  # Agents
    "alp",  # Alpha Team
    "arc",  # Arctic
    "atl",  # Atlantis alt
    "bat",  # Batman
    "cas",  # Castle
    "cca",  # City/Castle alt
    "cc",  # City Chase
    "cl",  # Classic
    "col",  # Collectible Minifigures
    "coldnd",  # Collectible D&D
    "colhp",  # Collectible Harry Potter
    "colmar",  # Collectible Marvel
    "collon",  # Collectible LOTR
    "coltlm",  # Collectible LEGO Movie
    "colnin",  # Collectible Ninjago
    "cre",  # Creator
    "cty",  # City
    "dim",  # Dimensions
    "dis",  # Disney
    "dp",  # Disney Princess
    "edu",  # Education/Dacta
    "elf",  # Elves
    "fab",  # Fabuland
    "fst",  # Forestmen
    "frnd",  # Friends
    "gal",  # Galaxy Squad
    "gen",  # Generic
    "har",  # Harry generic
    "hol",  # Holiday/Seasonal
    "hob",  # The Hobbit
    "hp",  # Harry Potter
    "hs",  # Hidden Side
    "ice",  # Ice Planet
    "idea",  # Ideas
    "ind",  # Indiana Jones
    "jw",  # Jurassic World
    "lor",  # Lord of the Rings
    "mk",  # Monkie Kid
    "min",  # Minifigures generic
    "mof",  # Modulars alt
    "mvl",  # Marvel
    "nba",  # NBA
    "nin",  # Ninjago old
    "njo",  # Ninjago new
    "njr",  # Ninjago reboot
    "ora",  # Overwatch
    "pac",  # Pac-Man
    "pi",  # Pirates
    "pm",  # Power Miners
    "potc",  # Pirates of the Caribbean
    "prince",  # Prince of Persia
    "pur",  # Pursuits
    "rac",  # Racers
    "res",  # Rescue
    "sc",  # Space Classic
    "sh",  # Super Heroes
    "she",  # She-Ra
    "shf",  # Super Heroes Female
    "soc",  # Soccer/Sports
    "spd",  # Speed Champions
    "spj",  # Spider-Man
    "sw",  # Star Wars
    "tlbm",  # LEGO Batman Movie
    "tlm",  # LEGO Movie
    "tlnm",  # LEGO Ninjago Movie
    "toy",  # Toy Story
    "twn",  # Town
    "vik",  # Vikings
    "wc",  # World City
    "ww",  # Wizarding World
)

    def _effective_item_type(self, rd):
        """Return 'M' if the part_id belongs to BL Minifigs catalog, else 'P'.
        Checks item_type already stored (set by scan-heads via BL API verification),
        then falls back to prefix list for offline use."""
        stored = rd.get("item_type", "P")
        if stored == "M":
            return "M"
        pid = str(rd.get("part_id", "")).lower().strip()
        if pid and not pid[0].isdigit():
            # Check prefix list as fast offline fallback
            if any(pid.startswith(p) for p in self._MINIFIG_CATALOG_PREFIXES):
                return "M"
        return stored

    # ── Alternatives navigation (per-row) ─────────────────────────────────────
    def _alt_state_from_row(self, rd):
        """Capture the current identity fields for alternative navigation."""
        return {
            "part_id": rd.get("part_id"),
            "part_name": rd.get("part_name"),
            "item_type": rd.get("item_type"),
            "confidence": rd.get("confidence"),
            # include dual-cam color swaps if they happened
            "color_id": rd.get("color_id"),
            "color_name": rd.get("color_name"),
        }

    def _alt_nav_push(self, rd, action=""):
        """Push current identity onto back stack, clear forward stack."""
        back = rd.setdefault("_alt_back", [])
        back.append(self._alt_state_from_row(rd))
        if len(back) > 25:
            del back[:-25]
        rd["_alt_fwd"] = []
        if action:
            rd["_alt_last_action"] = action

    def _alt_apply_state(self, rd, st, log_prefix="↩"):
        """Apply a previously captured alt state to the row and refresh UI."""
        if not st:
            return
        pid = st.get("part_id") or ""
        nm = st.get("part_name") or ""
        it = st.get("item_type") or "P"
        cf = st.get("confidence") or 0

        rd["part_id"] = pid
        rd["part_name"] = nm
        rd["item_type"] = self._effective_item_type({"part_id": pid, "item_type": it})
        rd["confidence"] = cf

        # restore color if present (useful for dual-cam swaps)
        if st.get("color_id") is not None:
            rd["color_id"] = st.get("color_id")
        if st.get("color_name") is not None:
            rd["color_name"] = st.get("color_name")

        # UI labels
        if rd.get("_pid_lbl"):
            rd["_pid_lbl"].setText(("👤 " if rd["item_type"] == "M" else "") + (pid or "—"))
            rd["_pid_lbl"].setStyleSheet(f"color:{ACCENT2};font-weight:bold;font-size:12px;")
        if rd.get("_name_lbl"):
            name_short = nm[:38] + "…" if len(nm) > 38 else nm
            rd["_name_lbl"].setText(name_short if name_short else "—")
            rd["_name_lbl"].setToolTip(nm or "")
        if rd.get("_confl"):
            cc = SUCCESS if cf >= 0.7 else WARNING if cf >= 0.4 else ACCENT
            rd["_confl"].setText(f"{cf:.0%}")
            rd["_confl"].setStyleSheet(f"color:{cc};font-weight:bold;font-size:12px;")

        # Refresh BL image + pricing (pid/color may have changed)
        if rd.get("_bl_lbl"):
            rd["_bl_lbl"].setText("…")
            threading.Thread(
                target=self._load_bl_img_for,
                args=(pid, rd.get("color_id", 0), rd.get("item_type", "P"), rd["_bl_lbl"]),
                daemon=True,
            ).start()
        rd["price"] = None
        rd["medium_price"] = None
        if rd.get("_price_lbl"):
            rd["_price_lbl"].setText("—")
            rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
        self._fetch_price_for_row(rd)
        self._fetch_price_guide(rd)
        self._update_total_value()
        self._log(f"{log_prefix}  Row: {pid} — {nm} ({cf:.0%})", "success")

    def _alt_back(self, rd):
        back = rd.get("_alt_back") or []
        if not back:
            self._log("No previous alternative for this row", "info")
            return
        cur = self._alt_state_from_row(rd)
        fwd = rd.setdefault("_alt_fwd", [])
        fwd.append(cur)
        st = back.pop()
        self._alt_apply_state(rd, st, log_prefix="↩ Back")

    def _alt_forward(self, rd):
        fwd = rd.get("_alt_fwd") or []
        if not fwd:
            self._log("No forward alternative for this row", "info")
            return
        cur = self._alt_state_from_row(rd)
        back = rd.setdefault("_alt_back", [])
        back.append(cur)
        st = fwd.pop()
        self._alt_apply_state(rd, st, log_prefix="↪ Forward")

    def _fetch_price_for_row(self, rd):
        """Re-fetch medium price for a row in background, update label when done."""
        pid = rd.get("part_id")
        if not pid: return
        rd["_price_lbl"].setText("…")
        rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")

        # Semaphore limits concurrent BL API calls — prevents rate-limit timeouts
        if not hasattr(self, "_bl_price_sem"):
            self._bl_price_sem = threading.Semaphore(3)

        def fetch_and_emit():
            color_id = rd.get("color_id", 0)
            currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if not line or line.startswith("#") or "=" not in line: continue
                        k, v = line.split("=", 1)
                        v = v.split("#")[0].strip().strip('"').strip("'")
                        env[k.strip()] = v
                km = {"CONSUMER_KEY":   ["CONSUMER_KEY"],
                      "CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":          ["TOKEN", "ACCESS_TOKEN"],
                      "TOKEN_SECRET":   ["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                missing = [k for k in km if k not in bl]
                if missing:
                    self.log_message.emit(f"⚠  .env missing: {missing}", "warning")
                    self._price_ready.emit(rd, None); return
                from requests_oauthlib import OAuth1
                import requests as _req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])
                itype = self._effective_item_type(rd)
                rd["item_type"] = itype
                if itype == "M":
                    url    = f"https://api.bricklink.com/api/store/v1/items/minifig/{pid}/price"
                    params = {"guide_type": "sold", "new_or_used": "U", "currency_code": currency}
                else:
                    url    = f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                    params = {"guide_type": "sold", "new_or_used": "U",
                              "currency_code": currency, "color_id": color_id}

                # Acquire semaphore — max 3 concurrent BL requests
                r = None
                with self._bl_price_sem:
                    for attempt in range(3):
                        try:
                            r = _req.get(url, params=params, auth=auth, timeout=12)
                            break
                        except Exception as e:
                            if attempt < 2 and ("timeout" in str(e).lower() or "timed out" in str(e).lower()):
                                import time as _t; _t.sleep(1.5 * (attempt + 1))
                                continue
                            raise
                if r is None:
                    self._price_ready.emit(rd, None); return

                if r.status_code == 404 and itype == "M":
                    url2    = f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                    params2 = {"guide_type": "sold", "new_or_used": "U",
                               "currency_code": currency, "color_id": color_id}
                    with self._bl_price_sem:
                        r2 = _req.get(url2, params=params2, auth=auth, timeout=12)
                    if r2.status_code == 200:
                        rd["item_type"] = "P"; r = r2
                    else:
                        self._price_ready.emit(rd, None); return
                elif r.status_code == 404:
                    self._price_ready.emit(rd, None); return
                elif r.status_code != 200:
                    self.log_message.emit(f"⚠  Price API {r.status_code} for {pid}: {r.text[:80]}", "warning")
                    self._price_ready.emit(rd, None); return
                data  = r.json().get("data", {})
                avg   = data.get("avg_price") or data.get("qty_avg_price")
                price = float(avg) if avg is not None and avg != "0.0000" and float(avg) > 0 else None
                self._price_ready.emit(rd, price)
            except Exception as e:
                self.log_message.emit(f"⚠  Price fetch ({pid}): {e}", "warning")
                self._price_ready.emit(rd, None)

        threading.Thread(target=fetch_and_emit, daemon=True).start()


    def _sort_results(self, key):
        """Sort visible rows by key, toggle asc/desc on repeated click."""
        active = [r for r in self._rows if not r.get("_deleted")]
        if not active: return

        # Toggle direction
        if self._sort_col == key:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = key
            self._sort_asc = True

        # Update header button labels
        for k, btn in self._hdr_btns.items():
            base = {"part_id":"Part ID","part_name":"Name","color_name":"Color",
                    "confidence":"Conf","price":"Price","qty":"Qty","remark":"Remark"}.get(k, k)
            if k == key:
                arrow = " ▲" if self._sort_asc else " ▼"
                btn.setText(base + arrow)
                btn.setStyleSheet(btn.styleSheet().replace(TEXT_DIM, ACCENT2))
            else:
                btn.setText(base)
                btn.setStyleSheet(btn.styleSheet().replace(ACCENT2, TEXT_DIM))

        # Sort key function
        def sort_val(r):
            v = r.get(key)
            if v is None:
                return "zzzzz" if key in ("color_name","part_name","part_id","remark") else float("inf")
            if isinstance(v, str): return v.lower()
            return v

        active.sort(key=sort_val, reverse=not self._sort_asc)

        # Reorder widgets in layout AND keep self._rows in sync so that
        # index-based operations (shift-click range, ctrl-click, _selected set)
        # remain correct after sorting.
        deleted = [r for r in self._rows if r.get("_deleted")]
        self._rows = active + deleted   # deleted rows keep their data, just hidden

        # Rebuild _selected by id — identity-based so sort never breaks it
        self._selected = {id(r) for r in self._rows if not r.get("_deleted") and r["_chk"].isChecked()}
        # Keep _last_clicked_rd as-is — it's an object reference, unaffected by sort

        for r in active:
            w = r["_widget"]
            self.results_list_layout.removeWidget(w)
        for r in active:
            self.results_list_layout.insertWidget(self.results_list_layout.count()-1, r["_widget"])

    # ── Selection & bulk actions ─────────────────────────────────────────────
    def _update_bulk_bar(self):
        if hasattr(self, "total_value_label"):
            self._update_total_value()
        n = len(self._selected)
        if n:
            self.bulk_label.setText(f"{n} selected")
            self.bulk_label.setStyleSheet(f"font-size:11px;color:{ACCENT2};font-weight:bold;")
            self.bulk_bar.setStyleSheet(
                f"QWidget#bulkBar {{ background:#2a2010; border:1px solid {ACCENT2}60; }}")
        else:
            self.bulk_label.setText("bulk")
            self.bulk_label.setStyleSheet(f"font-size:11px;color:{TEXT_DIM};font-style:italic;")
            self.bulk_bar.setStyleSheet(
                f"QWidget#bulkBar {{ background:#202020; border:1px solid {BORDER}; }}")
        # Enable/disable action buttons
        _action_btns = [self._bulk_set_qty, self._bulk_set_med, self._bulk_set_med15,
                        self._bulk_price, self._bulk_color, self._bulk_bg_color,
                        self._bulk_recolor, self._bulk_del]
        for b in _action_btns:
            b.setEnabled(n > 0)
        # Keep select-all checkbox in sync without triggering its signal
        self._sel_all_chk.blockSignals(True)
        active = [r for r in self._rows if not r.get("_deleted")]
        self._sel_all_chk.setCheckState(
            Qt.Checked if n == len(active) and active else Qt.Unchecked)
        self._sel_all_chk.blockSignals(False)

    def _select_all_rows(self):
        for r in self._rows:
            if not r.get("_deleted"):
                r["_chk"].blockSignals(True)
                r["_chk"].setChecked(True)
                r["_chk"].blockSignals(False)
                self._selected.add(id(r))
                w = r["_widget"]
                _p = w.palette()
                _p.setColor(w.backgroundRole(), QColor(ROW_SEL))
                w.setPalette(_p)
                w.setStyleSheet(f"QWidget#resultRow {{ background:{ROW_SEL}; border-bottom:1px solid {BORDER}; }}")
        self._update_bulk_bar()

    def _select_no_rows(self):
        for r in self._rows:
            r["_chk"].blockSignals(True)
            r["_chk"].setChecked(False)
            r["_chk"].blockSignals(False)
            w = r["_widget"]
            orig = r.get("_bg", "#2c2c2c")
            _p = w.palette()
            _p.setColor(w.backgroundRole(), QColor(orig))
            w.setPalette(_p)
            w.setStyleSheet(f"QWidget#resultRow {{ background:{orig}; border-bottom:1px solid {BORDER}; }}")
        self._selected.clear()
        self._last_clicked_rd = None
        self._update_bulk_bar()

    def _selected_rows(self):
        """Return list of row dicts in _selected, in current display order."""
        sel_ids = self._selected
        return [r for r in self._rows if id(r) in sel_ids and not r.get("_deleted")]

    def _bulk_edit_qty(self):
        self._save_undo_snapshot("bulk set qty")
        rows = self._selected_rows()
        if not rows: return
        val, ok = QInputDialog.getInt(self, "Bulk Edit Qty",
            f"Set quantity for {len(rows)} rows:", 1, 1, 9999)
        if not ok: return
        for rd in rows:
            rd["qty"] = val
            rd["_qty_lbl"].setText(str(val))
        self._log(f"✏  {len(rows)} rows → qty {val}", "info")

    def _bulk_set_medium(self):
        self._save_undo_snapshot("bulk set medium price")
        rows = [r for r in self._selected_rows() if r.get("medium_price") or r.get("price")]
        if not rows: self._log("No prices available for selected rows", "warning"); return
        for rd in rows:
            med = rd.get("medium_price") or rd.get("price")
            if not med: continue
            rd["price"] = med
            rd["_price_lbl"].setText(f"${med:.2f}")
            rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;")
        self._log(f"💲  {len(rows)} rows set to medium price", "info")
        self._update_total_value()

    def _bulk_set_medium15(self):
        self._save_undo_snapshot("bulk set +15% price")
        rows = [r for r in self._selected_rows() if r.get("medium_price") or r.get("price")]
        if not rows: self._log("No prices available for selected rows", "warning"); return
        for rd in rows:
            med = rd.get("medium_price") or rd.get("price")
            if not med: continue
            val = round(med * 1.15, 2)
            rd["price"] = val
            rd["_price_lbl"].setText(f"${val:.2f}")
            rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;")
        self._log(f"💲  {len(rows)} rows set to medium +15%", "info")
        self._update_total_value()

    def _bulk_set_min_01_all(self):
        """Set a minimum price floor on all non-deleted rows (user-configurable)."""
        from PyQt5.QtWidgets import QInputDialog
        # Load last used floor from station.cfg if available
        min_val = 0.10
        try:
            if CFG_PATH.exists():
                cfg = json.loads(CFG_PATH.read_text())
                mv = cfg.get("min_price_floor", 0.10)
                try:
                    min_val = max(0.0, float(mv))
                except Exception:
                    min_val = 0.10
        except Exception:
            min_val = 0.10
        val, ok = QInputDialog.getDouble(
            self,
            "Set minimum price",
            "Apply this minimum price to all rows (only raises prices below this):",
            min_val,
            0.0,
            999.0,
            2,
        )
        if not ok:
            return
        MIN_VAL = float(val)
        # Persist new floor
        try:
            cfg = {}
            if CFG_PATH.exists():
                try:
                    cfg = json.loads(CFG_PATH.read_text())
                except Exception:
                    cfg = {}
            cfg["min_price_floor"] = MIN_VAL
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception:
            pass
        rows = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        if not rows:
            self._log("No rows to update", "info")
            return
        self._save_undo_snapshot(f"set min ${MIN_VAL:.2f} (all)")
        changed = 0
        for rd in rows:
            p = rd.get("price")
            try:
                p_num = float(p) if p is not None and p != "" else None
            except Exception:
                p_num = None
            if p_num is None or p_num < MIN_VAL:
                rd["price"] = MIN_VAL
                if rd.get("_price_lbl"):
                    rd["_price_lbl"].setText(f"${MIN_VAL:.2f}")
                    rd["_price_lbl"].setStyleSheet(f"color:{WARNING};font-size:11px;")
                changed += 1
        self._update_total_value()
        self._log(f"⬇  {changed} row{'s' if changed != 1 else ''} set to min ${MIN_VAL:.2f}", "info")

    def _bulk_delete(self):
        self._save_undo_snapshot("bulk delete")
        rows = self._selected_rows()
        if not rows: return
        n = len(rows)
        for rd in rows:
            rd["_deleted"] = True
            rd["_widget"].hide()
        self._selected.clear()
        self._last_clicked_rd = None
        self._part_count = sum(1 for r in self._rows if not r.get("_deleted"))
        self.results_count.setText(f"{self._part_count} parts")
        self._update_bulk_bar()
        self._log(f"🗑  {n} rows deleted", "info")

    def _delete_error_rows(self):
        """Delete all unidentified (error) rows at once."""
        deleted = 0
        for rd in self._rows:
            if not rd.get("_deleted") and not rd.get("part_id"):
                rd["_deleted"] = True
                rd["_widget"].hide()
                deleted += 1
        if deleted:
            self._part_count = sum(1 for r in self._rows if not r.get("_deleted"))
            self.results_count.setText(f"{self._part_count} parts")
            self._log(f"🗑  {deleted} error row{'s' if deleted>1 else ''} deleted", "info")

    def _clear_results(self):
        self._selected.clear()
        if hasattr(self, 'bulk_bar'): self._update_bulk_bar()
        self._part_count = 0
        self._rows = []
        self._current_source_image = None  # reset group header tracking
        self.results_count.setText("0 parts")
        while self.results_list_layout.count() > 1:
            item = self.results_list_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()

    def _add_group_header(self, src_name, src_path, photo_idx):
        """Insert a thin separator row labelling a new source image (batch mode)."""
        from PyQt5.QtWidgets import QLabel as _QL
        hdr = QWidget()
        hdr.setObjectName("groupHeader")
        hdr.setAutoFillBackground(True)
        _hp = hdr.palette()
        _hp.setColor(hdr.backgroundRole(), QColor("#1a2535"))
        hdr.setPalette(_hp)
        hdr.setStyleSheet("QWidget#groupHeader { background:#1a2535; border-bottom:1px solid #2a4060; }")
        hdr.setFixedHeight(22)
        hl = QHBoxLayout(hdr); hl.setContentsMargins(10, 2, 8, 2); hl.setSpacing(6)
        icon_lbl = _QL("📷"); icon_lbl.setStyleSheet(f"font-size:10px;color:#5a8fc0;")
        name_lbl = _QL(src_name)
        name_lbl.setStyleSheet(f"font-size:10px;color:#7ab0d8;font-weight:bold;")
        idx_lbl  = _QL(f"  photo {photo_idx + 1}")
        idx_lbl.setStyleSheet(f"font-size:9px;color:{TEXT_DIM};")
        hl.addWidget(icon_lbl); hl.addWidget(name_lbl); hl.addWidget(idx_lbl)
        hl.addStretch()
        self.results_list_layout.insertWidget(self.results_list_layout.count() - 1, hdr)

    def _add_result_row(self, r):
        from PyQt5.QtWidgets import QLabel as _QL
        # Insert group header row when source image changes (batch mode)
        src_name = r.get("source_image_name", "")
        src_path = r.get("source_image", "")
        photo_idx = r.get("photo_idx", 0)
        if src_name and src_name != self._current_source_image:
            self._current_source_image = src_name
            self._add_group_header(src_name, src_path, photo_idx)
        self._part_count += 1
        identified   = bool(r.get("part_id"))
        color_method = r.get("color_method") or ""
        bg = ("#3d1515" if not identified else
              "#332010" if "unreliable" in color_method else
              "#28280d" if color_method == "unknown" else "#2c2c2c")

        row = QWidget()
        row.setObjectName("resultRow")
        row.setAutoFillBackground(True)
        _pal = row.palette()
        _pal.setColor(row.backgroundRole(), QColor(bg))
        row.setPalette(_pal)
        row.setStyleSheet(f"QWidget#resultRow {{ background:{bg}; border-bottom:1px solid {BORDER}; }}")
        row.setFixedHeight(148)
        # Outer vertical layout: top line (all fields) + bottom line (condition + remark)
        rv_outer = QVBoxLayout(row); rv_outer.setContentsMargins(0,0,0,0); rv_outer.setSpacing(0)
        top_w = QWidget(); top_w.setFixedHeight(116)
        rl = QHBoxLayout(top_w); rl.setContentsMargins(8,3,8,2); rl.setSpacing(6)
        rv_outer.addWidget(top_w)
        # Bottom line — condition toggle + remark/comment
        bot_w = QWidget(); bot_w.setFixedHeight(30)
        bot_w.setStyleSheet(f"background:transparent;")
        # Left margin = 8 (row margin) + 18(chk)+6 + 24(idx)+6 = 62 so remark starts under crop
        bl2 = QHBoxLayout(bot_w); bl2.setContentsMargins(62,0,8,3); bl2.setSpacing(6)
        rv_outer.addWidget(bot_w)

        # Selection checkbox
        chk = QCheckBox(); chk.setFixedWidth(18)
        chk.setStyleSheet("margin-left:2px;")
        rl.addWidget(chk)

        # Index
        il = _QL(str(r.get("index", self._part_count)))
        il.setFixedWidth(24); il.setAlignment(Qt.AlignCenter)
        il.setStyleSheet(f"color:{TEXT_DIM};font-size:10px;"); rl.addWidget(il)

        # (il is index only — qty gets its own label below after crop images)

        # Scan crop
        cl = _QL(); cl.setFixedSize(110,110)
        cl.setStyleSheet(f"background:{CARD_BG};border-radius:3px;"); cl.setAlignment(Qt.AlignCenter)
        cl.setCursor(Qt.PointingHandCursor)
        cl.setToolTip("Left-click on preview → set shadow color\nRight-click on preview → set background color")
        def _make_crop_rclick(rd_):
            def _crop_rclick(event):
                if event.button() == Qt.RightButton:
                    from PyQt5.QtWidgets import QMenu, QAction
                    menu = QMenu()
                    menu.setStyleSheet(f"QMenu{{background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};}}"
                                       f"QMenu::item{{padding:5px 16px;}}QMenu::item:selected{{background:{ACCENT2};color:#000;}}")
                    a_shadow = menu.addAction("🌑  Mark as shadow — suppress this color in detection")
                    a_bg     = menu.addAction("🖼  Mark as background — suppress this color in detection")
                    a_del    = menu.addAction("🗑  Delete this row (false detection)")
                    from PyQt5.QtGui import QCursor
                    chosen = menu.exec_(QCursor.pos())
                    if chosen in (a_shadow, a_bg):
                        # Sample dominant color from the crop image
                        crop_path = rd_.get("crop_image", "")
                        color = None
                        if crop_path and Path(crop_path).exists():
                            try:
                                import cv2, numpy as np
                                img = cv2.imread(crop_path)
                                if img is not None:
                                    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                                    # Sample center region to avoid edge artifacts
                                    h, w = img_rgb.shape[:2]
                                    cx, cy = w//2, h//2
                                    margin = max(5, min(w, h) // 4)
                                    region = img_rgb[max(0,cy-margin):cy+margin, max(0,cx-margin):cx+margin]
                                    med = np.median(region.reshape(-1, 3), axis=0)
                                    color = (int(med[0]), int(med[1]), int(med[2]))
                            except Exception:
                                pass
                        if color:
                            if chosen == a_shadow:
                                self._shadow_rgb = color
                                if hasattr(self, "shadow_swatch"):
                                    r2,g2,b2 = color
                                    lum = 0.299*r2+0.587*g2+0.114*b2
                                    fg = "#000" if lum > 128 else "#fff"
                                    self.shadow_swatch.setStyleSheet(
                                        f"background:rgb({r2},{g2},{b2});color:{fg};"
                                        f"border:1px solid {BORDER};border-radius:3px;font-size:10px;padding:0 6px;")
                                    self.shadow_swatch.setText(f"🌑 #{r2:02x}{g2:02x}{b2:02x}  (right-click to reset)")
                                self._log(f"🌑  Shadow color set to rgb{color} — active on next scan", "info")
                            else:
                                self._bg_rgb = color
                                self._log(f"🖼  Background color set to rgb{color} — active on next scan", "info")
                        else:
                            self._log("⚠  Could not sample color from crop image", "warning")
                    elif chosen == a_del:
                        rd_["_deleted"] = True
                        rd_["_widget"].hide()
                        self._part_count = sum(1 for x in self._rows if not x.get("_deleted"))
                        self.results_count.setText(f"{self._part_count} parts")
            return _crop_rclick
        cl.mousePressEvent = _make_crop_rclick(r)
        cp = r.get("crop_image","")
        if cp and Path(cp).exists():
            px = QPixmap(cp).scaled(110,110,Qt.KeepAspectRatio,Qt.SmoothTransformation)
            cl.setPixmap(px)
        else:
            cl.setText("?"); cl.setStyleSheet(f"color:{TEXT_DIM};background:{CARD_BG};border-radius:3px;")
        rl.addWidget(cl)

        # BrickLink reference image (loaded async)
        bl_img = _QL(); bl_img.setFixedSize(110,110)
        bl_img.setStyleSheet(f"background:{CARD_BG};border-radius:3px;"); bl_img.setAlignment(Qt.AlignCenter)
        bl_img.setCursor(Qt.PointingHandCursor); bl_img.setToolTip("Double-click to enlarge")
        bl_img.setText("…")
        rl.addWidget(bl_img)
        # Load BL image in background thread
        part_id_val = r.get("part_id")
        if part_id_val:
            itype_for_img = self._effective_item_type(r)  # corrects sw*/col*/sh* etc to "M"
            r["item_type"] = itype_for_img  # correct in-place so rest of row is consistent
            bl_img._itype = itype_for_img
            threading.Thread(target=self._load_bl_img_for,
                args=(part_id_val, r.get("color_id", 0), itype_for_img, bl_img,
                      r.get("thumb_url", "")),
                daemon=True).start()

        # Double-click crop → enlarge crop
        def make_crop_dbl(cp_):
            def on_dbl(event):
                if event.button() == Qt.LeftButton and cp_ and Path(cp_).exists():
                    self._show_enlarged("", None, cp_)
            return on_dbl
        cl.mouseDoubleClickEvent = make_crop_dbl(cp)

        # Double-click BL image → wired AFTER row_idx is assigned (see below)
        def make_bl_dbl(bl_lbl_, rd_ref_):
            def on_dbl(event):
                if event.button() != Qt.LeftButton:
                    return
                # Show full-size BrickLink catalog image in a popup
                pid   = rd_ref_.get("part_id","")
                itype = rd_ref_.get("item_type","P")
                color_id = rd_ref_.get("color_id", 0)
                if not pid:
                    return
                # Build full-size URLs (same pattern but no size limit)
                if itype == "M":
                    urls = [
                        f"https://img.bricklink.com/ItemImage/MN/0/{pid.lower()}.png",
                        f"https://img.bricklink.com/ItemImage/MN/0/{pid.upper()}.png",
                        f"https://img.bricklink.com/ItemImage/MF/0/{pid.lower()}.png",
                        f"https://img.bricklink.com/ItemImage/MF/0/{pid.upper()}.png",
                    ]
                else:
                    urls = [
                        f"https://img.bricklink.com/ItemImage/PN/{color_id}/{pid.lower()}.png",
                        f"https://img.bricklink.com/ItemImage/PN/0/{pid.lower()}.png",
                    ]
                from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel as _QLD, QHBoxLayout
                from PyQt5.QtCore import QByteArray, Qt as _Qt
                from PyQt5.QtGui import QPixmap as _QP
                dlg = QDialog(self)
                dlg.setWindowTitle(f"BrickLink — {pid}")
                dlg.resize(500, 500)
                dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
                vl = QVBoxLayout(dlg); vl.setContentsMargins(12,12,12,12); vl.setSpacing(8)
                img_lbl = _QLD("Loading…")
                img_lbl.setAlignment(_Qt.AlignCenter)
                img_lbl.setStyleSheet(f"background:{CARD_BG};border-radius:6px;font-size:12px;color:{TEXT_DIM};")
                img_lbl.setMinimumSize(400, 400)
                vl.addWidget(img_lbl)
                info = _QLD(f"{pid}  {'Minifig' if itype=='M' else f'Color {color_id}'}")
                info.setAlignment(_Qt.AlignCenter)
                info.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
                vl.addWidget(info)
                close_btn = self._btn("✕ Close", "#5a2020", dlg.accept)
                close_btn.setFixedHeight(26)
                bh = QHBoxLayout(); bh.addStretch(); bh.addWidget(close_btn)
                vl.addLayout(bh)

                def _fetch_full():
                    import urllib.request
                    for url in urls:
                        try:
                            req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
                            data = urllib.request.urlopen(req, timeout=10).read()
                            if len(data) < 200: continue
                            pm = _QP(); pm.loadFromData(QByteArray(data))
                            if not pm.isNull():
                                self._bl_img_signal.emit(img_lbl, pm)
                                return
                        except Exception:
                            continue
                    self._bl_img_signal.emit(img_lbl, _QP())

                # Slot: resize pixmap to fit dialog
                def _on_img(lbl, pm):
                    if lbl is not img_lbl: return
                    if not pm.isNull():
                        lbl.setPixmap(pm.scaled(480, 480, _Qt.KeepAspectRatio, _Qt.SmoothTransformation))
                        lbl.setText("")
                    else:
                        lbl.setText("Image not available")

                # Temporarily connect a one-shot slot
                _conn = [None]
                def _one_shot(lbl, pm):
                    _on_img(lbl, pm)
                    try: self._bl_img_signal.disconnect(_one_shot)
                    except Exception: pass
                self._bl_img_signal.connect(_one_shot)

                threading.Thread(target=_fetch_full, daemon=True).start()
                dlg.exec_()
                # Clean up connection if dialog closed before fetch finished
                try: self._bl_img_signal.disconnect(_one_shot)
                except Exception: pass
            return on_dbl

        # Part ID — clickable to edit/override
        pid_txt = r.get("part_id") or "—"
        if r.get("item_type") == "M": pid_txt = "👤 " + pid_txt
        pidl = _QL(pid_txt); pidl.setFixedWidth(90)
        pidl.setStyleSheet(f"color:{ACCENT2 if identified else TEXT_DIM};font-weight:bold;font-size:12px;")
        pidl.setCursor(Qt.PointingHandCursor)
        pidl.setToolTip("Click to override part ID")
        rl.addWidget(pidl)

        # Name
        full_name = r.get("part_name","—")
        name = full_name[:38]+"…" if len(full_name)>38 else full_name
        nl = _QL(name); nl.setFixedWidth(185)
        nl.setToolTip(full_name)
        nl.setStyleSheet(f"color:{TEXT};font-size:11px;"); rl.addWidget(nl)

        # Color swatch + name + source marker (who “won”: Brickognize vs scanner)
        cw = QWidget(); cw.setFixedWidth(115)
        cwh = QHBoxLayout(cw); cwh.setContentsMargins(0,0,0,0); cwh.setSpacing(4)
        is_minifig_row = (r.get("item_type") == "M")
        rgb = r.get("color_rgb",(128,128,128))
        sw = _QL(); sw.setFixedSize(13,13)
        if is_minifig_row:
            sw.setStyleSheet(f"background:#444;border-radius:3px;border:1px solid #444;")
        else:
            sw.setStyleSheet(f"background:rgb{tuple(rgb)};border-radius:3px;border:1px solid #444;")
        if is_minifig_row:
            cn = "—"
            cnl = _QL(cn); cnl.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
        else:
            cn = r.get("color_name","—"); cn = cn[:13] if len(cn)>13 else cn
            cnl = _QL(cn); cnl.setStyleSheet(f"color:{WARNING if 'unreliable' in color_method else TEXT};font-size:11px;")
        sw.setCursor(Qt.PointingHandCursor)
        sw.setToolTip("Click to override color")
        # NEW (2026): Source tag: "B" = Brickognize won, "S" = scanner(core pixels) won
        src = (r.get("color_source") or "").lower()
        src_tag = _QL("")
        src_tag.setFixedWidth(10)
        src_tag.setAlignment(Qt.AlignCenter)
        if not is_minifig_row and src in ("brickognize", "scan"):
            if src == "brickognize":
                src_tag.setText("B")
                src_tag.setStyleSheet(f"color:#000;background:{ACCENT2};border-radius:2px;font-size:9px;font-weight:bold;")
            else:
                src_tag.setText("S")
                src_tag.setStyleSheet(f"color:#000;background:#7a9fff;border-radius:2px;font-size:9px;font-weight:bold;")
        else:
            src_tag.setText("")
            src_tag.setStyleSheet(f"color:{TEXT_DIM};font-size:9px;")

        # Tooltip: show both candidates if present
        if not is_minifig_row:
            scid = r.get("scan_color_id"); scn = r.get("scan_color_name"); scc = r.get("scan_color_conf")
            bcid = r.get("brickognize_color_id"); bcn = r.get("brickognize_color_name"); bcc = r.get("brickognize_color_conf")
            if scid or bcid:
                cnl.setToolTip(
                    f"Final: {r.get('color_name','—')} ({r.get('color_id','—')})  via {r.get('color_source','?')}\n"
                    f"Scanner: {scn or '—'} ({scid or '—'})  conf {scc if scc is not None else '—'}\n"
                    f"Brickognize: {bcn or '—'} ({bcid or '—'})  conf {bcc if bcc is not None else '—'}"
                )

        cwh.addWidget(sw); cwh.addWidget(src_tag); cwh.addWidget(cnl); rl.addWidget(cw)

        # Confidence + alternatives picker
        conf = r.get("confidence",0)
        cc = SUCCESS if conf>=0.7 else WARNING if conf>=0.4 else ACCENT
        conf_w = QWidget(); conf_w.setFixedWidth(48)
        conf_layout = QVBoxLayout(conf_w); conf_layout.setContentsMargins(0,0,0,0); conf_layout.setSpacing(1)
        confl = _QL(f"{conf:.0%}"); confl.setAlignment(Qt.AlignCenter)
        confl.setStyleSheet(f"color:{cc};font-weight:bold;font-size:12px;")
        conf_layout.addWidget(confl)
        # For print variants, inject mirror (R↔L) and base mold into alternatives
        if r.get("part_id"):
            import re as _re3
            _pid = r["part_id"]
            _alts = list(r.get("alternatives") or [])
            _existing = {a.get("part_id","").lower() for a in _alts} | {a.get("id","").lower() for a in _alts}

            # Mirror variant: R↔L  (e.g. 60581pb038R ↔ 60581pb038L)
            if _pid.endswith("R") or _pid.endswith("r"):
                _mirror = _pid[:-1] + "L"
                if _mirror.lower() not in _existing:
                    _alts.append({"part_id": _mirror, "id": _mirror,
                                  "name": f"Mirror: {_mirror}", "score": 0.0, "is_mold_variant": True})
            elif _pid.endswith("L") or _pid.endswith("l"):
                _mirror = _pid[:-1] + "R"
                if _mirror.lower() not in _existing:
                    _alts.append({"part_id": _mirror, "id": _mirror,
                                  "name": f"Mirror: {_mirror}", "score": 0.0, "is_mold_variant": True})

            # Base mold: strip print suffix (pb, pr, pl, pat, ps)
            _bm = _re3.match(r"^([0-9]+[a-z]?)(pb|pr|pl|pat|ps)", _pid.lower())
            if _bm:
                _base = _bm.group(1)
                if _base not in _existing and _base != _pid.lower():
                    _alts.append({"part_id": _base, "id": _base,
                                  "name": f"Base mold: {_base}", "score": 0.0, "is_mold_variant": True})

            # Lettered dimensional siblings: 4460a → 4460b, 4460c, etc.
            # Parts ending in a single lowercase letter are dimensional variants
            # e.g. 4460a (2×1×3 slope) vs 4460b (2×2×3 slope)
            # Inject adjacent siblings (±1 letter) — limits noise from non-existent variants
            _lm = _re3.match(r"^([0-9]+)([a-f])$", _pid.lower())
            if _lm:
                _base_num = _lm.group(1)
                _cur_letter = _lm.group(2)
                _cur_ord = ord(_cur_letter)
                for _delta in [-1, 1, 2]:  # prev, next, next+1
                    _new_ord = _cur_ord + _delta
                    if ord("a") <= _new_ord <= ord("f"):
                        _sibling = _base_num + chr(_new_ord)
                        if _sibling not in _existing:
                            _alts.append({"part_id": _sibling, "id": _sibling,
                                          "name": f"Dimensional variant: {_sibling}",
                                          "score": 0.0, "is_mold_variant": True})

            if len(_alts) > len(r.get("alternatives") or []):
                r["alternatives"] = _alts
                r["has_mold_variants"] = True

        # Show "▾ alt" button whenever alternatives exist
        if r.get("alternatives"):
            # Re-derive has_mold_variants from alternatives content in case scan-heads didn't set it
            has_mv = r.get("has_mold_variants", False) or any(
                a.get("is_mold_variant") for a in r.get("alternatives", []))
            # Also check if part_id contains print variant pattern
            if not has_mv and r.get("part_id"):
                import re as _re2
                if _re2.match(r"^[0-9]+[a-z]?(pb|pr|pat|ps)", r["part_id"].lower()):
                    has_mv = True
            has_dual = any(a.get("_dual_cam_alt") for a in r.get("alternatives", []))
            if has_dual:
                btn_text = "▾ 📷 alt"
                btn_tip  = "Dual-cam alternative — click to see the other camera's result"
                btn_col  = "#7a9fff"
            elif has_mv:
                btn_text = "▾ 🔀"
                btn_tip  = "Mold variants available — click to see BrickLink alternatives"
                btn_col  = WARNING
            else:
                btn_text = "▾ alt"
                btn_tip  = "Pick alternative Brickognize result"
                btn_col  = ACCENT2
            alt_btn = _QL(btn_text); alt_btn.setAlignment(Qt.AlignCenter)
            alt_btn.setFixedHeight(14)
            alt_btn.setStyleSheet(f"color:{btn_col};font-size:9px;text-decoration:underline;font-weight:{'bold' if has_mv or has_dual else 'normal'};")
            alt_btn.setCursor(Qt.PointingHandCursor)
            alt_btn.setToolTip(btn_tip)
            conf_layout.addWidget(alt_btn)
        else:
            conf_layout.addStretch()
        rl.addWidget(conf_w)

        # Price
        price = r.get("price")
        price_txt = f"${price:.2f}" if price else "—"
        prl = _QL(price_txt); prl.setFixedWidth(58); prl.setAlignment(Qt.AlignCenter)
        prl.setStyleSheet(f"color:{'#7ec8a0' if price else TEXT_DIM};font-size:11px;")
        rl.addWidget(prl)

        # Qty label — after price to match header order
        qty_val = r.get("qty", 1)
        qtyl = _QL(str(qty_val)); qtyl.setFixedWidth(38); qtyl.setAlignment(Qt.AlignCenter)
        qtyl.setStyleSheet(f"color:{TEXT};font-size:12px;font-weight:bold;"); rl.addWidget(qtyl)

        # Status icon — last column
        if not identified:                    st,sc = "✗",ACCENT
        elif color_method=="minifig":         st,sc = "✓",SUCCESS
        elif "unreliable" in color_method:    st,sc = "⚠",WARNING
        elif color_method=="unknown":         st,sc = "?",WARNING
        else:                                 st,sc = "✓",SUCCESS
        stl = _QL(st); stl.setFixedWidth(30); stl.setAlignment(Qt.AlignCenter)
        stl.setStyleSheet(f"color:{sc};font-weight:bold;font-size:13px;"); rl.addWidget(stl)

        rl.addStretch()

        # ── Bottom line: condition toggle + remark + comment ─────────────────
        cond_val = r.get("condition", "U")

        def _make_cond_toggle(rd_ref_):
            cond_btn = _QL()
            cond_btn.setFixedSize(28, 20)
            cond_btn.setAlignment(Qt.AlignCenter)
            cond_btn.setCursor(Qt.PointingHandCursor)
            cond_btn.setToolTip("Click to toggle Used / New")
            def _refresh_cond():
                c = rd_ref_.get("condition", "U")
                cond_btn.setText(c)
                cond_btn.setStyleSheet(
                    f"background:{'#1a3a1a' if c=='N' else '#2a1800'};"
                    f"color:{'#5fca7a' if c=='N' else ACCENT2};"
                    f"font-size:10px;font-weight:bold;border-radius:3px;"
                    f"border:1px solid {'#2a5a2a' if c=='N' else '#4a3000'};")
            _refresh_cond()
            def _toggle(event):
                if event.button() == Qt.LeftButton:
                    rd_ref_["condition"] = "N" if rd_ref_.get("condition","U") == "U" else "U"
                    _refresh_cond()
            cond_btn.mousePressEvent = _toggle
            return cond_btn, _refresh_cond

        # Remark label — italic, dimmed, click to edit
        rmk_val = r.get("remark", "") or ""
        rmk_lbl = _QL(rmk_val if rmk_val else "remark…")
        rmk_lbl.setFixedHeight(18)
        rmk_lbl.setStyleSheet(
            f"color:{'#aaa' if not rmk_val else TEXT};font-size:10px;font-style:italic;"
            f"background:transparent;padding:0 4px;")
        rmk_lbl.setCursor(Qt.PointingHandCursor)
        rmk_lbl.setToolTip("Click to set remark (shown in BrickLink lot)")
        rmk_tpl_btn = QPushButton("📋")
        rmk_tpl_btn.setFixedSize(18, 18)
        rmk_tpl_btn.setStyleSheet(f"font-size:9px;background:transparent;border:none;color:{TEXT_DIM};")
        rmk_tpl_btn.setToolTip("Insert condition note into comment field")
        rmk_tpl_btn.setCursor(Qt.PointingHandCursor)
        bl2.addWidget(rmk_lbl, 2)
        bl2.addWidget(rmk_tpl_btn)

        # Comment label
        cmt_val = r.get("comment", "") or ""
        cmt_lbl = _QL(cmt_val if cmt_val else "comment…")
        cmt_lbl.setFixedHeight(18)
        cmt_lbl.setStyleSheet(
            f"color:{'#aaa' if not cmt_val else TEXT};font-size:10px;font-style:italic;"
            f"background:transparent;padding:0 4px;")
        cmt_lbl.setCursor(Qt.PointingHandCursor)
        cmt_lbl.setToolTip("Click to set comment (internal note)")
        bl2.addWidget(cmt_lbl, 2)

        # ⊞ Colors + 💎 Value buttons — parts only
        if not is_minifig_row:
            clr_btn = QPushButton("⊞ Colors")
            clr_btn.setFixedHeight(18)
            clr_btn.setCursor(Qt.PointingHandCursor)
            clr_btn.setStyleSheet(f"font-size:9px;color:{ACCENT2};background:transparent;"
                                   f"border:none;text-decoration:underline;padding:0 4px;")
            clr_btn.setToolTip("Clone row into top 9 colors by sales volume on BrickLink")
            bl2.addWidget(clr_btn)

            val_btn = QPushButton("💎 Value")
            val_btn.setFixedHeight(18)
            val_btn.setCursor(Qt.PointingHandCursor)
            val_btn.setStyleSheet(f"font-size:9px;color:#b8860b;background:transparent;"
                                   f"border:none;text-decoration:underline;padding:0 4px;")
            val_btn.setToolTip("Find the single most valuable color for this part — discovery tool")
            bl2.addWidget(val_btn)

        # Part Out button for minifigs
        if is_minifig_row:
            po_btn = QPushButton("⊞ Part Out")
            po_btn.setFixedHeight(18)
            po_btn.setCursor(Qt.PointingHandCursor)
            po_btn.setStyleSheet(f"font-size:9px;color:{ACCENT2};background:transparent;"
                                  f"border:none;text-decoration:underline;padding:0 4px;")
            po_btn.setToolTip("Replace this minifig with its individual parts (BrickLink subsets)")
            bl2.addWidget(po_btn)
            po_val_lbl = QLabel("⊞ …")
            po_val_lbl.setStyleSheet(f"font-size:9px;color:{TEXT_DIM};background:transparent;padding:0 2px;")
            po_val_lbl.setToolTip("Estimated part-out value (sum of median sold prices)")
            bl2.addWidget(po_val_lbl)

        # Condition toggle — fixed widths to align under Price+Qty+Status columns
        # Price(58)+sp(6)+Qty(38)+sp(6) = 108px spacer, then cond(28) under Status(30)
        _spacer_w = QWidget(); _spacer_w.setFixedWidth(102); _spacer_w.setStyleSheet("background:transparent;")
        bl2.addWidget(_spacer_w)
        cond_btn, _refresh_cond_fn = _make_cond_toggle({"condition": cond_val})
        cond_btn.setFixedWidth(30)  # match status column width
        bl2.addWidget(cond_btn)

        # Store row data with direct widget refs for reliable editing
        row_data = dict(r)
        row_data.setdefault("condition", "U")
        row_data.setdefault("remark", "")
        row_data.setdefault("comment", "")
        row_data["medium_price"] = r.get("price")  # preserve original medium price
        # Keep source image and bbox for click-to-preview
        row_data.setdefault("source_image", r.get("source_image", ""))
        row_data.setdefault("bbox", r.get("bbox"))
        row_data["_bg"] = bg  # store for selection restore
        # Snapshot Brickognize's color so it can always be restored later.
        #
        # Why this matters (2026 fix):
        # Previously we only stored one color field (`color_id`) which might already be the
        # scanner's own sampled color (or a recolor override). The “BG Color” button
        # therefore couldn't truly restore Brickognize's attribution — it just restored
        # whatever was in `color_id` at row creation time.
        #
        # scan-heads now preserves BOTH sources:
        # - brickognize_color_* (Brickognize-inferred)
        # - scan_color_*        (our core-pixel sampler)
        # and `color_id` is the merged final display color.
        scan_cid  = r.get("brickognize_color_id", r.get("color_id"))
        scan_name = r.get("brickognize_color_name", r.get("color_name"))
        scan_rgb  = (BL_COLORS.get(scan_cid, ("?", (128,128,128)))[1]
                     if scan_cid in BL_COLORS else r.get("color_rgb"))
        row_data["_scan_color_id"]   = scan_cid
        row_data["_scan_color_name"] = scan_name
        row_data["_scan_color_rgb"]  = scan_rgb
        row_data.update({
            "_widget":    row,
            "_qty_lbl":   qtyl,
            "_swatch":    sw,
            "_color_lbl": cnl,
            "_price_lbl": prl,
            "_chk":       chk,
            "_bl_lbl":    bl_img,
            "_pid_lbl":   pidl,
            "_name_lbl":  nl,
            "_confl":     confl,
            "_cond_btn":  cond_btn,
            "_crop_lbl":  cl,
            "_cond_refresh": _refresh_cond_fn,
            "_rmk_lbl":   rmk_lbl,
            "_cmt_lbl":   cmt_lbl,
            "_po_val_lbl": po_val_lbl if is_minifig_row else None,
        })
        self._rows.append(row_data)
        row_idx = len(self._rows) - 1
        bl_img.mouseDoubleClickEvent = make_bl_dbl(bl_img, row_data)  # object refs, not index

        # Wire Colors + Value buttons
        if not is_minifig_row:
            def _make_clr(rd_):
                return lambda: self._clone_top_colors(rd_)
            clr_btn.clicked.connect(_make_clr(row_data))

            def _make_val(rd_):
                return lambda: self._show_best_value_color(rd_)
            val_btn.clicked.connect(_make_val(row_data))

        # Wire Part Out button (QPushButton — won't bubble to row handler)
        if is_minifig_row:
            def _make_po(rd_):
                return lambda: self._part_out_minifig(rd_)
            po_btn.clicked.connect(_make_po(row_data))

        # Wire alt button if it exists
        if r.get("alternatives"):
            def _make_alt_click(rd_):
                def _on_click(event):
                    if event.button() == Qt.LeftButton:
                        self._pick_alternative(rd_)
                return _on_click
            alt_btn.mousePressEvent = _make_alt_click(row_data)

        # Wire cond_btn toggle to row_data (now valid)
        def _make_cond_wire(rd_):
            def _refresh():
                c = rd_.get("condition", "U")
                rd_["_cond_btn"].setText(c)
                rd_["_cond_btn"].setStyleSheet(
                    f"background:{'#1a3a1a' if c=='N' else '#2a1800'};"
                    f"color:{'#5fca7a' if c=='N' else ACCENT2};"
                    f"font-size:10px;font-weight:bold;border-radius:3px;"
                    f"border:1px solid {'#2a5a2a' if c=='N' else '#4a3000'};")
            def _toggle(event):
                if event.button() == Qt.LeftButton:
                    rd_["condition"] = "N" if rd_.get("condition","U") == "U" else "U"
                    _refresh()
            rd_["_cond_btn"].mousePressEvent = _toggle
            _refresh()   # render with actual row_data value
        _make_cond_wire(row_data)

        # Wire remark click → inline edit dialog
        # Wire remark template button
        def _make_tpl_click(rd_):
            def _on_tpl():
                self._show_remark_templates(rd_)
            return _on_tpl
        rmk_tpl_btn.clicked.connect(_make_tpl_click(row_data))

        def _make_rmk_click(rd_):
            def _on_click(event):
                if event.button() == Qt.LeftButton:
                    val, ok = QInputDialog.getText(
                        self, "Remark", "Remark (shown in BrickLink lot):",
                        text=rd_.get("remark","") or "")
                    if ok:
                        rd_["remark"] = val.strip()
                        lbl = rd_["_rmk_lbl"]
                        lbl.setText(val.strip() if val.strip() else "remark…")
                        lbl.setStyleSheet(
                            f"color:{'#aaa' if not val.strip() else TEXT};font-size:10px;"
                            f"font-style:italic;background:transparent;padding:0 4px;")
            return _on_click
        rmk_lbl.mousePressEvent = _make_rmk_click(row_data)

        # Wire comment click → inline edit dialog
        def _make_cmt_click(rd_):
            def _on_click(event):
                if event.button() == Qt.LeftButton:
                    val, ok = QInputDialog.getText(
                        self, "Comment", "Comment (internal note):",
                        text=rd_.get("comment","") or "")
                    if ok:
                        rd_["comment"] = val.strip()
                        lbl = rd_["_cmt_lbl"]
                        lbl.setText(val.strip() if val.strip() else "comment…")
                        lbl.setStyleSheet(
                            f"color:{'#aaa' if not val.strip() else TEXT};font-size:10px;"
                            f"font-style:italic;background:transparent;padding:0 4px;")
            return _on_click
        cmt_lbl.mousePressEvent = _make_cmt_click(row_data)

        # Wire checkbox to selection tracking (checkbox-initiated changes)
        def make_chk_handler(rd_):
            def on_chk(state):
                w = rd_["_widget"]
                if state:
                    self._selected.add(id(rd_))
                    new_bg = ROW_SEL
                else:
                    self._selected.discard(id(rd_))
                    new_bg = rd_.get("_bg", "#2c2c2c")
                _p = w.palette()
                _p.setColor(w.backgroundRole(), QColor(new_bg))
                w.setPalette(_p)
                w.setStyleSheet(f"QWidget#resultRow {{ background:{new_bg}; border-bottom:1px solid {BORDER}; }}")
                self._update_bulk_bar()
            return on_chk
        chk.stateChanged.connect(make_chk_handler(row_data))

        # Unified mouse handler: left-click = toggle selection (shift = range),
        # ctrl-click = open BrickLink, right-click handled by context menu
        pid_for_click  = part_id_val
        col_for_click  = r.get("color_id","")
        type_for_click = self._effective_item_type(r)
        r["item_type"] = type_for_click   # correct in-place
        def make_mouse(rd_ref, pid, color_id, itype):
            # rd_ref is the actual row dict — identity never changes even after sort
            _dbl_pending = [False]

            def _sel_rd(rd, state):
                """Select/deselect a row dict by object identity."""
                if rd.get("_deleted"): return
                rd["_chk"].blockSignals(True)
                rd["_chk"].setChecked(state)
                rd["_chk"].blockSignals(False)
                if state:
                    self._selected.add(id(rd))
                    new_bg = ROW_SEL
                else:
                    self._selected.discard(id(rd))
                    new_bg = rd.get("_bg", "#2c2c2c")
                w = rd["_widget"]
                _p = w.palette()
                _p.setColor(w.backgroundRole(), QColor(new_bg))
                w.setPalette(_p)
                w.setStyleSheet(f"QWidget#resultRow {{ background:{new_bg}; border-bottom:1px solid {BORDER}; }}")

            def on_mouse(event):
                if _dbl_pending[0]:
                    _dbl_pending[0] = False
                    return
                if event.button() == Qt.LeftButton:
                    mods = event.modifiers()
                    # Start drag-select (no modifier = fresh drag)
                    if not (mods & Qt.ShiftModifier) and not (mods & Qt.ControlModifier):
                        self._drag_selecting = True
                        self._drag_start_rd  = rd_ref
                    if mods & Qt.ShiftModifier and self._last_clicked_rd is not None:
                        # Shift+click → select range between anchor and this row
                        # Find positions of anchor and current in current display order
                        active = [r for r in self._rows if not r.get("_deleted")]
                        try:
                            i_anchor = next(i for i,r in enumerate(active) if r is self._last_clicked_rd)
                            i_cur    = next(i for i,r in enumerate(active) if r is rd_ref)
                            lo, hi   = min(i_anchor, i_cur), max(i_anchor, i_cur)
                            for r in active[lo:hi+1]:
                                _sel_rd(r, True)
                        except StopIteration:
                            _sel_rd(rd_ref, True)
                        self._last_clicked_rd = rd_ref
                        self._update_bulk_bar()
                    elif mods & Qt.ControlModifier:
                        # Ctrl+click → toggle without clearing others
                        new_state = not rd_ref["_chk"].isChecked()
                        _sel_rd(rd_ref, new_state)
                        self._last_clicked_rd = rd_ref
                        self._update_bulk_bar()
                    else:
                        # Plain click → select ONLY this row (deselect all others)
                        # This is standard list behaviour — enables easy single selection
                        # while Ctrl+click and Shift+click handle multi-select
                        for rd_ in self._rows:
                            if not rd_.get("_deleted") and id(rd_) in self._selected and rd_ is not rd_ref:
                                _sel_rd(rd_, False)
                        _sel_rd(rd_ref, True)
                        self._last_clicked_rd = rd_ref
                        src_path = rd_ref.get("source_image", "")
                        bbox     = rd_ref.get("bbox")
                        if src_path:
                            self._preview_source_image(src_path, bbox)
                        self._fetch_price_guide(rd_ref)
            def on_double(event):
                if event.button() == Qt.LeftButton:
                    _dbl_pending[0] = True   # suppress the preceding single-click
                    mods = event.modifiers()
                    if mods & Qt.ControlModifier:
                        # Ctrl+double-click → open BrickLink
                        if pid:
                            import webbrowser
                            clean_pid = pid.strip()
                            if itype == "M":
                                url = f"https://www.bricklink.com/v2/catalog/catalogitem.page?M={clean_pid}"
                            else:
                                url = f"https://www.bricklink.com/v2/catalog/catalogitem.page?P={clean_pid}"
                                if color_id: url += f"&idColor={color_id}"
                            webbrowser.open(url)
                    else:
                        src_path  = rd_ref.get("source_image", "")
                        bbox      = rd_ref.get("bbox")
                        crop_path = rd_ref.get("crop_image", "")
                        if src_path or crop_path:
                            # Double-click → open full-res enlargement popup
                            self._show_enlarged(src_path, bbox, crop_path or None)
                        elif pid:
                            import webbrowser
                            clean_pid = pid.strip()
                            if itype == "M":
                                url = f"https://www.bricklink.com/v2/catalog/catalogitem.page?M={clean_pid}"
                            else:
                                url = f"https://www.bricklink.com/v2/catalog/catalogitem.page?P={clean_pid}"
                                if color_id: url += f"&idColor={color_id}"
                            webbrowser.open(url)
            def on_move(event):
                if not self._drag_selecting or self._drag_start_rd is None:
                    return
                # Find which row is under the cursor in the container
                global_pos = event.globalPos()
                active = [r for r in self._rows if not r.get("_deleted")]
                # Determine drag range between _drag_start_rd and the row under cursor
                start_i = next((i for i,r in enumerate(active) if r is self._drag_start_rd), None)
                hover_i = None
                for i, r in enumerate(active):
                    w = r["_widget"]
                    tl = w.mapToGlobal(w.rect().topLeft())
                    br = w.mapToGlobal(w.rect().bottomRight())
                    if tl.y() <= global_pos.y() <= br.y():
                        hover_i = i
                        break
                if start_i is None or hover_i is None:
                    return
                lo, hi = min(start_i, hover_i), max(start_i, hover_i)
                # Select rows in range, deselect those outside
                for i, r in enumerate(active):
                    _sel_rd(r, lo <= i <= hi)
                self._last_clicked_rd = rd_ref
                self._update_bulk_bar()

            def on_release(event):
                if event.button() == Qt.LeftButton:
                    self._drag_selecting = False
                    self._drag_start_rd  = None

            return on_mouse, on_move, on_release, on_double
        row.setCursor(Qt.PointingHandCursor)
        row.setToolTip("Click: select  |  Ctrl+click: add to selection  |  Shift+click: range select  |  Double-click: full-res view  |  Ctrl+double-click: BrickLink  |  Right-click: edit")
        on_mouse, on_move, on_release, on_dbl = make_mouse(row_data, pid_for_click, col_for_click, type_for_click)
        # Attach to row AND both sub-widgets — child widgets eat parent mouse events
        for _w in (row, top_w, bot_w):
            _w.mousePressEvent       = on_mouse
            _w.mouseMoveEvent        = on_move
            _w.mouseReleaseEvent     = on_release
            _w.mouseDoubleClickEvent = on_dbl

        # Part ID label click → inline edit
        def make_pid_click(idx):
            def on_pid_click(event):
                if event.button() == Qt.LeftButton:
                    self._edit_part_id(idx)
            return on_pid_click
        pidl.mousePressEvent = make_pid_click(row_idx)

        # Color swatch click → color override picker
        def make_sw_click(idx):
            def on_sw_click(event):
                if event.button() == Qt.LeftButton:
                    self._override_color(idx)
            return on_sw_click
        sw.mousePressEvent = make_sw_click(row_idx)

        # Right-click context menu
        row.setContextMenuPolicy(Qt.CustomContextMenu)
        def make_ctx(rd_ref_):
            def show_ctx(pos):
                from PyQt5.QtWidgets import (QMenu, QInputDialog, QDialog,
                    QVBoxLayout, QListWidget, QListWidgetItem,
                    QDialogButtonBox, QLineEdit)
                rd   = rd_ref_
                pid  = rd.get("part_id","?")
                med  = rd.get("medium_price") or rd.get("price")
                sel_rows = self._selected_rows()
                sel_n = len(sel_rows)

                menu = QMenu()
                if sel_n > 1 and any(r is rd for r in sel_rows):
                    # Multi-select: only show “(selected)” bulk actions to avoid duplicate entries.
                    hdr = menu.addAction(f"— Apply to selected ({sel_n}) —")
                    hdr.setEnabled(False)
                    a_qty_sel   = menu.addAction("✏  Edit quantity (selected)")
                    a_color_sel = menu.addAction("🎨  Pick color (selected)")
                    a_price_sel = menu.addAction("✏  Edit price (selected)")
                    a_med_sel   = menu.addAction("💲  Set medium (selected)")
                    a_med15_sel = menu.addAction("💲  Set medium +15% (selected)")
                    a_rmk_sel   = menu.addAction("📋  Add remark (selected)")
                    a_iph_sel   = menu.addAction("📱  Recheck with latest iPhone photo (selected)")
                    a_del_sel   = menu.addAction("🗑  Delete (selected)")
                    menu.addSeparator()
                    # In multi-select mode we intentionally skip the single-row variants
                    # (a_qty, a_color, etc.) to avoid visual duplicates.
                    a_qty = a_color = a_price = a_med = a_med15 = a_pg = a_rescan = a_rescan_wide = a_iph = None
                else:
                    a_qty_sel = a_color_sel = a_price_sel = a_med_sel = a_med15_sel = a_rmk_sel = a_iph_sel = a_del_sel = None
                    a_qty   = menu.addAction("✏  Edit quantity")
                    a_color = menu.addAction("🎨  Pick color")
                menu.addSeparator()
                a_del   = menu.addAction("🗑  Delete row")
                a_dup   = menu.addAction("⧉  Duplicate row")
                a_rmk   = menu.addAction("📋  Add remark")
                menu.addSeparator()
                a_price = menu.addAction("✏  Edit price manually")
                a_med   = menu.addAction(f"💲  Set medium{f'  (${med:.2f})' if med else ''}")
                a_med15 = menu.addAction(f"💲  Set medium +15%{f'  (${med*1.15:.2f})' if med else ''}")
                menu.addSeparator()
                _pg     = rd.get("_price_guide")
                _med_u  = _pg.get("avg_u") if _pg else None
                _med_n  = _pg.get("avg_n") if _pg else None
                _med_lbl = ""
                if _med_u:   _med_lbl = f"  │  Used avg ${float(_med_u):.2f}"
                elif _med_n: _med_lbl = f"  │  New avg ${float(_med_n):.2f}"
                a_pg    = menu.addAction(f"📊  Price guide{_med_lbl}")
                menu.addSeparator()
                a_rescan      = menu.addAction("🪓  Split attempt (touching/overlapping parts)")
                a_rescan_wide = menu.addAction("🪓+  Split attempt (wider crop)")
                if not med: a_med.setEnabled(False); a_med15.setEnabled(False)
                a_iph = menu.addAction("📱  Recheck with latest iPhone photo")
                # Alt navigation (only meaningful after an alternative was applied)
                menu.addSeparator()
                a_back = menu.addAction("↩  Back (previous ID)")
                a_fwd  = menu.addAction("↪  Forward")
                a_back.setEnabled(bool(rd.get("_alt_back")))
                a_fwd.setEnabled(bool(rd.get("_alt_fwd")))

                # Fix popup position — clamp to screen so it never flies off edge
                from PyQt5.QtWidgets import QApplication as _QApp
                _gpos   = rd["_widget"].mapToGlobal(pos)
                _screen = _QApp.primaryScreen().availableGeometry()
                _msize  = menu.sizeHint()
                _gx = min(_gpos.x(), _screen.right()  - _msize.width())
                _gy = min(_gpos.y(), _screen.bottom() - _msize.height())
                _gx = max(_gx, _screen.left())
                _gy = max(_gy, _screen.top())
                from PyQt5.QtCore import QPoint as _QPoint
                chosen = menu.exec_(_QPoint(_gx, _gy))

                def _set_price(val, is_medium=False):
                    rd["price"] = val
                    if is_medium:
                        rd["medium_price"] = val  # update medium when refetched
                    rd["_price_lbl"].setText(f"${val:.2f}")
                    rd["_price_lbl"].setStyleSheet("color:#7ec8a0;font-size:11px;")

                if chosen == a_qty_sel:
                    self._bulk_edit_qty()
                elif chosen == a_color_sel:
                    self._bulk_override_color()
                elif chosen == a_price_sel:
                    self._bulk_set_price()
                elif chosen == a_med_sel:
                    self._bulk_set_medium()
                elif chosen == a_med15_sel:
                    self._bulk_set_medium15()
                elif chosen == a_rmk_sel:
                    self._bulk_set_remark()
                elif chosen == a_iph_sel:
                    self._iphone_recheck_selected(sel_rows)
                elif chosen == a_del_sel:
                    self._bulk_delete()

                if chosen == a_pg:
                    self._fetch_price_guide(rd)

                elif chosen == a_back:
                    self._alt_back(rd)
                elif chosen == a_fwd:
                    self._alt_forward(rd)

                elif chosen == a_rescan:
                    self._split_attempt(rd)
                elif chosen == a_rescan_wide:
                    self._rescan_region(rd, expand=0.5, require_min_parts=2, use_watershed=True)
                elif chosen == a_iph:
                    self._iphone_recheck_row(rd)

                elif chosen == a_qty:
                    cur = rd.get("qty",1)
                    val, ok = QInputDialog.getInt(self,"Edit Qty",f"Qty for {pid}:",cur,1,9999)
                    if ok:
                        rd["qty"] = val
                        rd["_qty_lbl"].setText(str(val))
                        self._log(f"✏  qty → {val}","info")

                elif chosen == a_color:
                    known_ids  = rd.get("known_color_ids", [])
                    known_list_ctx = [(cid, BL_COLORS[cid][0], BL_COLORS[cid][1])
                                      for cid in known_ids if cid in BL_COLORS]
                    other_list_ctx = [(cid, name, rgb) for cid,(name,rgb) in BL_COLORS_SORTED
                                      if cid not in set(known_ids)]
                    # Show only known colors; all others available via "Show all"
                    color_list = known_list_ctx if known_list_ctx else (known_list_ctx + other_list_ctx)
                    _ctx_all = [False]
                    header = (f"⭐ {len(known_list_ctx)} known colors for this part"
                              if known_list_ctx else "All BrickLink colors")
                    dlg = QDialog(self); dlg.setWindowTitle(f"Pick color — {pid}")
                    dlg.setStyleSheet("background:#1a2433;color:#cdd6e0;"); dlg.resize(300,400)
                    vl = QVBoxLayout(dlg)
                    from PyQt5.QtWidgets import QLabel as _QL2
                    vl.addWidget(_QL2(header, styleSheet="color:#7ec8a0;font-size:10px;padding:2px;"))
                    search = QLineEdit(); search.setPlaceholderText("Filter...")
                    search.setStyleSheet("background:#0d1b2a;color:#cdd6e0;padding:4px;border-radius:4px;")
                    vl.addWidget(search)
                    lw = QListWidget(); lw.setStyleSheet("background:#0d1b2a;color:#cdd6e0;border:none;")
                    from PyQt5.QtGui import QColor as QC2
                    def _pop(filt="",_cl=color_list):
                        lw.clear()
                        for cid,cname,rgb in _cl:
                            if filt.lower() not in cname.lower(): continue
                            item = QListWidgetItem(f"  {cname}")
                            item.setBackground(QC2(*rgb))
                            item.setForeground(QC2(20,20,20) if sum(rgb)>400 else QC2(240,240,240))
                            item.setData(32,(cid,cname,rgb)); lw.addItem(item)
                        cur_id = rd.get("color_id")
                        for i in range(lw.count()):
                            if lw.item(i).data(32)[0]==cur_id: lw.setCurrentRow(i); break
                    _pop(); search.textChanged.connect(_pop); vl.addWidget(lw)
                    bb = QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
                    bb.setStyleSheet("background:{};".format(CARD_BG))
                    bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
                    vl.addWidget(bb); lw.itemDoubleClicked.connect(lambda _: dlg.accept())
                    if dlg.exec_()==QDialog.Accepted and lw.currentItem():
                        cid,cname,rgb = lw.currentItem().data(32)
                        rd["color_id"]=cid; rd["color_name"]=cname; rd["color_rgb"]=rgb
                        rd["_swatch"].setStyleSheet(f"background:rgb{tuple(rgb)};border-radius:3px;border:1px solid #444;")
                        rd["_color_lbl"].setText(cname[:13])
                        self._log(f"🎨  color → {cname} (id {cid})","info")
                        # Refresh BL reference image for new color
                        bl_lbl = rd.get("_bl_lbl")
                        if bl_lbl:
                            bl_lbl.setPixmap(QPixmap())
                            bl_lbl.setText("…")
                            bl_lbl._itype = rd.get("item_type","P")
                            pid_for_img = rd.get("part_id","")
                            threading.Thread(target=self._load_bl_img_for,
                                args=(pid_for_img, cid, rd.get("item_type","P"), bl_lbl),
                                daemon=True).start()
                        rd["price"] = None
                        rd["medium_price"] = None
                        rd["_price_lbl"].setText("—")
                        rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
                        self._fetch_price_for_row(rd)
                        self._fetch_price_guide(rd)

                elif chosen == a_med and med:   _set_price(med, is_medium=True); self._log(f"💲  → ${med:.2f} (medium)","info")
                elif chosen == a_med15 and med:  _set_price(round(med*1.15,2)); self._log(f"💲  → ${med*1.15:.2f} (+15%)","info")
                elif chosen == a_price:
                    cur = rd.get("price") or 0.0
                    val,ok = QInputDialog.getDouble(self,"Edit Price",f"Price for {pid}:",cur,0,9999,2)
                    if ok: _set_price(val); self._log(f"💲  → ${val:.2f}","info")
                elif chosen == a_rescan:
                    self._rescan_region(rd)
                elif chosen == a_rescan_wide:
                    self._rescan_region(rd, expand=0.5)
                elif chosen == a_rmk:
                    self._show_remark_templates(rd)
                elif chosen == a_dup:
                    new_r = {k: v for k, v in rd.items()
                             if not k.startswith("_")}  # strip widget refs
                    saved_scroll = self.results_scroll.verticalScrollBar().value()

                    self._add_result_row(new_r)
                    new_rd = self._rows[-1]
                    src_idx = next((i for i,r in enumerate(self._rows[:-1]) if r is rd), None)
                    if src_idx is not None:
                        self._rows.pop()
                        self._rows.insert(src_idx + 1, new_rd)
                        layout = self.results_list_layout
                        widget = new_rd.get("_widget")
                        if widget:
                            layout.removeWidget(widget)
                            src_widget = rd.get("_widget")
                            src_pos = -1
                            for i in range(layout.count()):
                                if layout.itemAt(i) and layout.itemAt(i).widget() is src_widget:
                                    src_pos = i
                                    break
                            if src_pos >= 0:
                                layout.insertWidget(src_pos + 1, widget)

                    # Restore scroll position
                    QTimer.singleShot(0, lambda v=saved_scroll:
                        self.results_scroll.verticalScrollBar().setValue(v))

                    self._part_count = sum(1 for x in self._rows if not x.get("_deleted"))
                    self.results_count.setText(f"{self._part_count} parts")
                    self._log(f"⧉  Duplicated after current row", "info")
                elif chosen == a_del:
                    rd["_deleted"]=True; rd["_widget"].hide()
                    self._part_count = sum(1 for x in self._rows if not x.get("_deleted"))
                    self.results_count.setText(f"{self._part_count} parts")
                    self._log(f"🗑  Row deleted","info")
            return show_ctx
        row.customContextMenuRequested.connect(make_ctx(row_data))

        # Auto-fetch price if missing (minifigs always need this; parts may also miss)
        if identified and not r.get("price"):
            QTimer.singleShot(50, lambda rd=row_data: self._fetch_price_for_row(rd))
        # Auto-fetch part-out value for minifig rows
        if is_minifig_row and identified:
            QTimer.singleShot(200, lambda rd=row_data: self._fetch_po_value(rd))

        self.results_list_layout.insertWidget(self.results_list_layout.count()-1, row)
        self.results_count.setText(f"{self._part_count} parts")
        if hasattr(self, "total_value_label"):
            self._update_total_value()

    # ── Preview ───────────────────────────────────────────────────────────────
    # ── Preview ───────────────────────────────────────────────────────────────
    # ── Price Guide ──────────────────────────────────────────────────────────

    def _pg_set_condition(self, cond):
        self._pg_condition = cond
        self.pg_toggle_new.setStyleSheet(
            f"background:{ACCENT2 if cond=='N' else CARD_BG};color:{TEXT};font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #333;padding:0 4px;")
        self.pg_toggle_used.setStyleSheet(
            f"background:{ACCENT2 if cond=='U' else CARD_BG};color:{TEXT};font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #333;padding:0 4px;")
        if self._pg_current_rd:
            self._fetch_price_guide(self._pg_current_rd)

    def _pg_set_guide(self, guide):
        self._pg_guide = guide
        self.pg_toggle_sold.setStyleSheet(
            f"background:{ACCENT2 if guide=='sold' else CARD_BG};color:{TEXT};font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #333;padding:0 4px;")
        self.pg_toggle_stock.setStyleSheet(
            f"background:{ACCENT2 if guide=='stock' else CARD_BG};color:{TEXT};font-size:11px;font-weight:bold;border-radius:4px;border:1px solid #333;padding:0 4px;")
        if self._pg_current_rd:
            self._fetch_price_guide(self._pg_current_rd)

    def _pg_close(self):
        """Collapse the price guide panel. Console drag won't reopen it."""
        sizes = self._pg_splitter.sizes()
        self._pg_saved_size = sizes[1] if sizes[1] > 20 else 180
        # Lock price guide to zero height so console handle can't pull it open
        self._pg_panel_widget.setMaximumHeight(0)
        self._pg_splitter.setSizes([sizes[0] + sizes[1], 0, sizes[2]])

    def _pg_open(self):
        """Reopen the price guide panel."""
        self._pg_panel_widget.setMaximumHeight(16777215)  # reset Qt default
        sizes = self._pg_splitter.sizes()
        saved = getattr(self, "_pg_saved_size", 180)
        total = sizes[0] + sizes[1]
        # Save scroll position — price guide opening can cause scroll jump
        sv = self.results_scroll.verticalScrollBar().value()
        self._pg_splitter.setSizes([total - saved, saved, sizes[2]])
        QTimer.singleShot(0, lambda v=sv: self.results_scroll.verticalScrollBar().setValue(v))

    def _fetch_price_guide(self, rd):
        """Fetch full price guide data for a row and populate the price guide panel."""
        pid      = rd.get("part_id")
        color_id = rd.get("color_id", 0)
        currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"
        itype    = self._effective_item_type(rd)
        if not pid:
            self._clear_price_guide()
            return
        self._pg_current_rd = rd
        short = pid if len(pid) <= 16 else pid[:14] + "…"
        self.pg_part_label.setText(f"{short}")
        for w in [self.pg_min_u, self.pg_avg_u, self.pg_qty_avg_u,
                  self.pg_max_u, self.pg_lots_u, self.pg_units_u,
                  self.pg_min_n, self.pg_avg_n, self.pg_qty_avg_n,
                  self.pg_max_n, self.pg_lots_n, self.pg_units_n]:
            w.setText("…")

        def fetch():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if not line or line.startswith("#") or "=" not in line: continue
                        k, v = line.split("=", 1)
                        v = v.split("#")[0].strip().strip('"').strip("'")
                        env[k.strip()] = v
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    self._price_guide_ready.emit(rd, {}); return
                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])
                results = {}
                # Fetch both sold + stock for both new + used in parallel
                import concurrent.futures as cf
                def _fetch_one(guide, condition):
                    if itype == "M":
                        url = f"https://api.bricklink.com/api/store/v1/items/minifig/{pid}/price"
                        params = {"guide_type": guide, "new_or_used": condition, "currency_code": currency}
                    else:
                        url = f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                        params = {"guide_type": guide, "new_or_used": condition,
                                  "currency_code": currency, "color_id": color_id}
                    try:
                        r = req.get(url, params=params, auth=auth, timeout=8)
                        if r.status_code == 200:
                            return (guide, condition), r.json().get("data", {})
                    except Exception:
                        pass
                    return (guide, condition), {}
                with cf.ThreadPoolExecutor(max_workers=4) as pool:
                    futs = [pool.submit(_fetch_one, g, c)
                            for g in ["sold","stock"] for c in ["U","N"]]
                    for f in cf.as_completed(futs):
                        key, data = f.result()
                        results[key] = data
                self._price_guide_ready.emit(rd, results)
            except Exception as e:
                self._price_guide_ready.emit(rd, {})

        threading.Thread(target=fetch, daemon=True).start()

    def _on_price_guide_ready(self, rd, results):
        """Populate price guide panel — both Used and New columns simultaneously."""
        if rd is not self._pg_current_rd:
            return  # stale result for a different row
        guide    = self._pg_guide
        currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"

        def _fmt(val):
            if val is None or str(val) == "0.0000" or float(val or 0) == 0:
                return "—"
            return f"{float(val):.2f}"

        def _fill_col(cond, min_w, avg_w, qty_w, max_w, lots_w, units_w):
            data = results.get((guide, cond), {})
            min_w.setText(_fmt(data.get("min_price")))
            avg_w.setText(_fmt(data.get("avg_price")))
            qty_w.setText(_fmt(data.get("qty_avg_price")))
            max_w.setText(_fmt(data.get("max_price")))
            lots_w.setText(str(data.get("total_lots")     or "—"))
            units_w.setText(str(data.get("total_quantity") or "—"))
            # Highlight qty_avg in green if available
            has_val = _fmt(data.get("qty_avg_price")) != "—"
            qty_w.setStyleSheet(
                f"font-size:10px;font-weight:bold;color:{'#27ae60' if has_val else '#555'};")

        _fill_col("U", self.pg_min_u, self.pg_avg_u, self.pg_qty_avg_u,
                       self.pg_max_u, self.pg_lots_u, self.pg_units_u)
        _fill_col("N", self.pg_min_n, self.pg_avg_n, self.pg_qty_avg_n,
                       self.pg_max_n, self.pg_lots_n, self.pg_units_n)

        # Store Used qty_avg for apply buttons (primary use case)
        data_u = results.get((guide, "U"), {})
        avg_u  = data_u.get("qty_avg_price") or data_u.get("avg_price")
        self._pg_last_avg = float(avg_u) if avg_u and float(avg_u) > 0 else None
        self.pg_apply_btn.setEnabled(self._pg_last_avg is not None)
        self.pg_apply_med_btn.setEnabled(self._pg_last_avg is not None)

        # Reopen if user manually clicks a row (only if currently closed)
        if self._pg_splitter.sizes()[1] < 20:
            self._pg_open()

    # ── Inline Part ID Edit ───────────────────────────────────────────────────
    def _edit_part_id(self, idx):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QLabel
        rd  = self._rows[idx]
        old_pid = rd.get("part_id", "") or ""
        dlg = QDialog(self)
        dlg.setWindowTitle("Override Part ID")
        dlg.setFixedWidth(340)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
        vl = QVBoxLayout(dlg); vl.setSpacing(8)
        vl.addWidget(QLabel(f"Current: <b>{old_pid}</b>  —  enter correct BrickLink part ID"))
        le = QLineEdit(old_pid)
        le.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};padding:4px;font-size:13px;")
        le.selectAll()
        vl.addWidget(le)
        # Type selector
        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("Type:"))
        type_part = self._btn("Part", CARD_BG, lambda: None, w=71)
        type_mini = self._btn("Minifig", CARD_BG, lambda: None, w=84)
        type_part.setFixedHeight(24); type_mini.setFixedHeight(24)
        type_part.setCheckable(True); type_mini.setCheckable(True)
        current_type = rd.get("item_type", "P")
        type_part.setChecked(current_type == "P")
        type_mini.setChecked(current_type == "M")
        def _sel_type(t):
            type_part.setChecked(t=="P"); type_mini.setChecked(t=="M")
            type_part.setStyleSheet(f"background:{ACCENT2 if t=='P' else CARD_BG};color:{TEXT};font-size:11px;border-radius:4px;border:1px solid #333;")
            type_mini.setStyleSheet(f"background:{ACCENT2 if t=='M' else CARD_BG};color:{TEXT};font-size:11px;border-radius:4px;border:1px solid #333;")
        type_part.clicked.connect(lambda: _sel_type("P"))
        type_mini.clicked.connect(lambda: _sel_type("M"))
        _sel_type(current_type)
        type_row.addWidget(type_part); type_row.addWidget(type_mini); type_row.addStretch()
        vl.addLayout(type_row)
        # Buttons
        btns = QHBoxLayout()
        ok_btn  = self._btn("✓ Apply", SUCCESS, dlg.accept, w=104)
        cxl_btn = self._btn("Cancel", CARD_BG, dlg.reject, w=91)
        ok_btn.setFixedHeight(28); cxl_btn.setFixedHeight(28)
        btns.addStretch(); btns.addWidget(cxl_btn); btns.addWidget(ok_btn)
        vl.addLayout(btns)
        le.returnPressed.connect(dlg.accept)
        if dlg.exec_() != QDialog.Accepted:
            return
        new_pid = le.text().strip()
        if not new_pid or new_pid == old_pid:
            return
        new_type = "M" if type_mini.isChecked() else "P"
        rd["part_id"]   = new_pid
        rd["item_type"] = new_type
        # Update part ID label
        pid_txt = ("👤 " if new_type=="M" else "") + new_pid
        rd["_pid_lbl"].setText(pid_txt[:12] + "…" if len(pid_txt)>12 else pid_txt)
        rd["_pid_lbl"].setStyleSheet(f"color:{ACCENT2};font-weight:bold;font-size:12px;")
        # Update name label to show it was overridden
        rd["_name_lbl"].setText("(overridden)")
        rd["_name_lbl"].setStyleSheet(f"color:{WARNING};font-size:11px;font-style:italic;")
        # Reload BL image
        rd["_bl_lbl"].setText("…")
        threading.Thread(target=self._load_bl_img_for,
            args=(new_pid, rd.get("color_id",0), new_type, rd["_bl_lbl"]),
            daemon=True).start()
        # Re-fetch price
        self._fetch_price_for_row(rd)
        # Re-fetch price guide
        self._fetch_price_guide(rd)
        self._log(f"✏  Row {idx+1}: part ID overridden {old_pid} → {new_pid} ({new_type})", "info")

    # ── Color Override ────────────────────────────────────────────────────────
    def _override_color(self, idx):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QListWidget, QListWidgetItem
        rd = self._rows[idx]
        if rd.get("item_type") == "M":
            return  # minifigs have no color
        # Import color table from scan-heads
        try:
            import importlib.util, sys
            spec = importlib.util.spec_from_file_location("scan_heads", "scan-heads.py")
            sh   = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(sh)
            BL_COLORS = sh.BRICKLINK_COLORS
        except Exception:
            from scan_gui_colors import BRICKLINK_COLORS as BL_COLORS  # fallback
            BL_COLORS = {}

        dlg = QDialog(self)
        dlg.setWindowTitle("Override Color")
        dlg.setFixedSize(300, 420)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
        vl = QVBoxLayout(dlg); vl.setSpacing(6)
        vl.addWidget(QLabel(f"Part: <b>{rd.get('part_id','?')}</b>  —  select color:"))
        # Header row: title + "Show all" toggle
        hdr_row = QHBoxLayout()
        known_ids  = rd.get("known_color_ids", [])
        known_set  = set(known_ids)
        known_list_ov = [(cid, BL_COLORS[cid][0], BL_COLORS[cid][1])
                         for cid in known_ids if cid in BL_COLORS]
        other_list_ov = [(cid, name, rgb) for cid,(name,rgb) in BL_COLORS_SORTED
                         if cid not in known_set]
        if known_list_ov:
            known_lbl = QLabel(f"⭐ = {len(known_list_ov)} known BL colors  (all colors shown)")
            known_lbl.setStyleSheet(f"font-size:10px;color:{ACCENT2};")
            hdr_row.addWidget(known_lbl)
        else:
            known_lbl = QLabel("No known colors from BL API — showing all")
            known_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
            hdr_row.addWidget(known_lbl)
        hdr_row.addStretch()
        vl.addLayout(hdr_row)

        search = QLineEdit(); search.setPlaceholderText("Filter colors…")
        search.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};padding:4px;")
        vl.addWidget(search)
        lst = QListWidget()
        lst.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};font-size:12px;")

        # Always show all colors — known ones marked with ⭐ at top
        def _populate(filt=""):
            lst.clear()
            from PyQt5.QtGui import QColor as QC
            cur_id = rd.get("color_id")
            known_ids_set = {cid for cid, _, _ in known_list_ov}
            for cid, cname, rgb in known_list_ov + other_list_ov:
                if filt and filt.lower() not in cname.lower(): continue
                star = "⭐ " if cid in known_ids_set else "   "
                tick = "✓ " if cid == cur_id else "  "
                item = QListWidgetItem(f"{star}{tick}{cname}  (id {cid})")
                item.setData(Qt.UserRole, (cid, cname, rgb))
                item.setBackground(QC(*rgb))
                lum = 0.299*rgb[0]+0.587*rgb[1]+0.114*rgb[2]
                item.setForeground(QC(0,0,0) if lum>128 else QC(255,255,255))
                lst.addItem(item)
            for i in range(lst.count()):
                d = lst.item(i).data(Qt.UserRole)
                if d and d[0] == cur_id:
                    lst.setCurrentRow(i); break
        _populate()
        search.textChanged.connect(_populate)
        vl.addWidget(lst)
        btns = QHBoxLayout()
        ok_btn  = self._btn("✓ Apply", SUCCESS, dlg.accept, w=104)
        cxl_btn = self._btn("Cancel", CARD_BG, dlg.reject, w=91)
        ok_btn.setFixedHeight(28); cxl_btn.setFixedHeight(28)
        btns.addStretch(); btns.addWidget(cxl_btn); btns.addWidget(ok_btn)
        vl.addLayout(btns)
        lst.itemDoubleClicked.connect(lambda: dlg.accept())
        if dlg.exec_() != QDialog.Accepted:
            return
        sel = lst.currentItem()
        if not sel: return
        cid, cname, rgb = sel.data(Qt.UserRole)
        rd["color_id"]   = cid
        rd["color_name"] = cname
        rd["color_rgb"]  = rgb
        # Update swatch
        rd["_swatch"].setStyleSheet(f"background:rgb{tuple(rgb)};border-radius:3px;border:1px solid #444;")
        rd["_color_lbl"].setText(cname[:13])
        rd["_color_lbl"].setStyleSheet(f"color:{TEXT};font-size:11px;")
        # Re-fetch price with new color
        self._fetch_price_for_row(rd)
        self._fetch_price_guide(rd)
        # Reload BL image with new color
        rd["_bl_lbl"].setText("…")
        threading.Thread(target=self._load_bl_img_for,
            args=(rd.get("part_id",""), cid, rd.get("item_type","P"), rd["_bl_lbl"]),
            daemon=True).start()
        _ridx = next((i+1 for i,r in enumerate(self._rows) if r is rd), "?")
        self._log(f"🎨  Row {_ridx}: color → {cname} (id {cid})", "info")

    # ── Lot Merging ───────────────────────────────────────────────────────────
    def _merge_duplicate_lots(self, silent=False):
        """Merge rows with identical part_id + color_id into one, summing qty."""
        seen = {}   # key → first rd
        to_delete = []
        for rd in self._rows:
            if rd.get("_deleted"): continue
            pid = rd.get("part_id")
            if not pid: continue
            key = (pid, rd.get("color_id", 0), rd.get("item_type", "P"))
            if key in seen:
                first = seen[key]
                first["qty"] = first.get("qty", 1) + rd.get("qty", 1)
                first["_qty_lbl"].setText(str(first["qty"]))
                to_delete.append(rd)
            else:
                seen[key] = rd
        for rd in to_delete:
            rd["_deleted"] = True
            rd["_widget"].setVisible(False)
        merged = len(to_delete)
        if merged:
            self._part_count = sum(1 for r in self._rows if not r.get("_deleted"))
            self.results_count.setText(f"{self._part_count} parts")
            self._update_total_value()
            if not silent:
                self._log(f"🔀  Merged {merged} duplicate lot(s)", "success")
            else:
                self._log(f"🔀  Auto-merged {merged} duplicate lot(s)", "info")
        else:
            self._log("🔀  No duplicate lots found", "info")
        return merged

    def _show_best_value_color(self, rd):
        """
        Find the single most valuable color for this part and show it as a popup.
        Value score = avg_sold_price / max(stock_qty, 1)
        This is a discovery tool — teaches you about rare colors worth hunting for.
        """
        pid  = rd.get("part_id", "")
        pname = rd.get("part_name", pid)
        if not pid: return
        self._log(f"💎  Finding most valuable color for {pid}...", "info")

        def fetch():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1); env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    self.log_message.emit("⚠  Missing credentials", "warning"); return

                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"], bl["TOKEN"], bl["TOKEN_SECRET"])

                r = req.get(f"https://api.bricklink.com/api/store/v1/items/part/{pid}/colors",
                            auth=auth, timeout=10)
                if r.status_code != 200:
                    self.log_message.emit(f"⚠  Colors API error {r.status_code}", "warning"); return

                known   = r.json().get("data", [])
                current = rd.get("color_id", 0)

                scored = []
                for c in known:
                    cid = c.get("color_id")
                    try:
                        sold  = req.get(f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                                        f"?color_id={cid}&guide_type=sold&new_or_used=U", auth=auth, timeout=6)
                        stock = req.get(f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                                        f"?color_id={cid}&guide_type=stock&new_or_used=U", auth=auth, timeout=6)
                        avg_price = 0.0; stock_qty = 9999
                        if sold.status_code  == 200: avg_price = float(sold.json().get("data",{}).get("avg_price",0) or 0)
                        if stock.status_code == 200: stock_qty = int(stock.json().get("data",{}).get("total_quantity",9999) or 9999)
                        if avg_price > 0:
                            score = avg_price / max(stock_qty, 1) * 1000
                            cname = BL_COLORS[cid][0] if cid in BL_COLORS else f"Color {cid}"
                            crgb  = BL_COLORS[cid][1] if cid in BL_COLORS else (128,128,128)
                            is_current = (cid == current)
                            scored.append((score, avg_price, stock_qty, cid, cname, crgb, is_current))
                    except Exception:
                        pass

                if not scored:
                    self.log_message.emit(f"⚠  No value data found for {pid}", "warning"); return

                scored.sort(reverse=True)
                best = scored[0]
                # Also find current color rank for context
                cur_rank = next((i+1 for i,s in enumerate(scored) if s[6]), None)
                self._value_result_ready.emit(pid, pname, best, cur_rank, len(scored))

            except Exception as e:
                self.log_message.emit(f"⚠  Value lookup failed: {e}", "warning")

        threading.Thread(target=fetch, daemon=True).start()

    def _on_value_result(self, pid, pname, best, cur_rank, total):
        """Show popup with the most valuable color discovery."""
        score, avg_price, stock_qty, cid, cname, crgb, is_current = best
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
        from PyQt5.QtGui import QColor, QFont
        from PyQt5.QtCore import Qt

        dlg = QDialog(self)
        dlg.setWindowTitle(f"💎  Most Valuable Color — {pid}")
        dlg.setStyleSheet("background:#0d1b2a;color:#cdd6e0;")
        dlg.setFixedWidth(340)
        vl = QVBoxLayout(dlg); vl.setSpacing(12); vl.setContentsMargins(20,20,20,20)

        # Part name
        lbl_part = QLabel(pname[:48])
        lbl_part.setStyleSheet("color:#7a9fff;font-size:11px;")
        vl.addWidget(lbl_part)

        # Color swatch + name
        row_w = QWidget(); rh = QHBoxLayout(row_w); rh.setContentsMargins(0,0,0,0); rh.setSpacing(10)
        swatch = QLabel(); swatch.setFixedSize(28,28)
        swatch.setStyleSheet(f"background:rgb{tuple(crgb)};border-radius:4px;border:1px solid #444;")
        rh.addWidget(swatch)
        lbl_color = QLabel(cname)
        lbl_color.setStyleSheet("color:#ffd700;font-size:16px;font-weight:bold;")
        rh.addWidget(lbl_color); rh.addStretch()
        vl.addWidget(row_w)

        # Price + stock
        rarity = "🔴 Very rare" if stock_qty <= 5 else "🟠 Rare" if stock_qty <= 20 else "🟡 Scarce" if stock_qty <= 50 else "🟢 Available"
        lbl_price = QLabel(f"Avg sold price:   ${avg_price:.2f}")
        lbl_price.setStyleSheet("font-size:13px;color:#7ec8a0;font-weight:bold;")
        lbl_stock = QLabel(f"Stock in market:  {stock_qty} units  {rarity}")
        lbl_stock.setStyleSheet("font-size:11px;color:#cdd6e0;")
        vl.addWidget(lbl_price); vl.addWidget(lbl_stock)

        # Context — where does current color rank?
        if cur_rank:
            ctx = QLabel(f"Your scanned color ranks #{cur_rank} of {total} known colors by value")
            ctx.setStyleSheet("font-size:10px;color:#888;font-style:italic;")
            ctx.setWordWrap(True); vl.addWidget(ctx)

        # Hint
        if not is_current:
            hint = QLabel(f"💡 If you have {pid} in {cname}, it's worth listing!")
            hint.setStyleSheet("font-size:11px;color:#ffd700;font-style:italic;")
            hint.setWordWrap(True); vl.addWidget(hint)

        btn = QPushButton("Got it"); btn.clicked.connect(dlg.accept)
        btn.setStyleSheet(f"background:#1a3a4a;color:#cdd6e0;border-radius:4px;padding:6px 20px;")
        vl.addWidget(btn, alignment=Qt.AlignCenter)

        # Position near center of screen
        from PyQt5.QtWidgets import QApplication as _QApp
        screen = _QApp.primaryScreen().availableGeometry()
        dlg.move(screen.center().x() - 170, screen.center().y() - 120)
        dlg.exec_()

    def _clone_value_colors(self, rd):
        """
        Clone row into top 5 most valuable colors for this part.
        Value score = avg_sold_price × (1 / max(stock_qty, 1))
        High price + low stock = genuinely rare & valuable.
        """
        pid = rd.get("part_id", "")
        if not pid: return
        self._log(f"💎  Finding most valuable colors for {pid}...", "info")

        def fetch_and_clone():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    self.log_message.emit("⚠  Missing credentials", "warning"); return

                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])

                # Get known colors
                r = req.get(f"https://api.bricklink.com/api/store/v1/items/part/{pid}/colors",
                            auth=auth, timeout=10)
                if r.status_code != 200:
                    self.log_message.emit(f"⚠  Colors API error {r.status_code}", "warning"); return

                known = r.json().get("data", [])
                current_cid = rd.get("color_id", 0)
                candidates = [c for c in known if c.get("color_id") != current_cid]
                if not candidates:
                    self.log_message.emit(f"⚠  No other known colors for {pid}", "warning"); return

                # For each color: fetch avg sold price + stock qty → value score
                scored = []
                for c in candidates:
                    cid = c.get("color_id")
                    try:
                        sold = req.get(
                            f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                            f"?color_id={cid}&guide_type=sold&new_or_used=U",
                            auth=auth, timeout=6)
                        stock = req.get(
                            f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                            f"?color_id={cid}&guide_type=stock&new_or_used=U",
                            auth=auth, timeout=6)
                        avg_price  = 0.0
                        stock_qty  = 9999
                        if sold.status_code == 200:
                            d = sold.json().get("data", {})
                            avg_price = float(d.get("avg_price", 0) or 0)
                        if stock.status_code == 200:
                            d2 = stock.json().get("data", {})
                            stock_qty = int(d2.get("total_quantity", 9999) or 9999)
                        # Value = price / scarcity — rare + expensive = high score
                        value_score = avg_price / max(stock_qty, 1) * 1000
                        # Must have some price data to be relevant
                        if avg_price > 0:
                            scored.append((cid, avg_price, stock_qty, value_score))
                    except Exception:
                        pass

                if not scored:
                    self.log_message.emit(f"⚠  No value data for {pid}", "warning"); return

                scored.sort(key=lambda x: x[3], reverse=True)
                top5 = scored[:5]

                clones = []
                import copy
                for cid, avg_price, stock_qty, score in top5:
                    cname, crgb = ("—", (128,128,128))
                    if cid in BL_COLORS:
                        cname = BL_COLORS[cid][0]; crgb = BL_COLORS[cid][1]
                    clone = copy.copy(rd)
                    clone["color_id"]   = cid
                    clone["color_name"] = cname
                    clone["color_rgb"]  = crgb
                    clone["qty"]        = 1
                    clone["price"]      = None
                    clone["from_color_clone"] = True
                    clone["_avg_price"] = avg_price
                    clone["_stock_qty"] = stock_qty
                    clones.append(clone)

                self.log_message.emit(
                    f"💎  {pid} → {len(clones)} valuable color clones", "success")
                self._value_colors_ready.emit(rd, clones)

            except Exception as e:
                self.log_message.emit(f"⚠  Value colors failed: {e}", "warning")

        threading.Thread(target=fetch_and_clone, daemon=True).start()

    def _clone_top_colors(self, rd):
        """
        Fetch the top 9 colors by total sales quantity for this part from BrickLink,
        excluding the color already on this row, then clone the row for each.
        Uses GET /items/part/{id}/colors — returns known colors with sales data.
        Falls back to price guide endpoint per color to rank by qty_avg.
        """
        pid = rd.get("part_id", "")
        if not pid:
            return
        self._log(f"🎨  Fetching top colors for {pid}...", "info")

        def fetch_and_clone():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    self.log_message.emit("⚠  Missing credentials for color clone", "warning")
                    return

                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])

                # Step 1: get known colors for this part
                url = f"https://api.bricklink.com/api/store/v1/items/part/{pid}/colors"
                r = req.get(url, auth=auth, timeout=10)
                if r.status_code != 200:
                    self.log_message.emit(f"⚠  Colors API error {r.status_code}", "warning")
                    return

                known = r.json().get("data", [])  # list of {color_id, color_name, ...}
                current_cid = rd.get("color_id", 0)
                candidates = [c for c in known if c.get("color_id") != current_cid]

                if not candidates:
                    self.log_message.emit(f"⚠  No other known colors for {pid}", "warning")
                    return

                # Step 2: fetch sales qty for each candidate color (last 6 months sold)
                # Use price guide endpoint — qty_avg_price * unit_quantity approximates sales
                color_sales = []
                for c in candidates:
                    cid = c.get("color_id")
                    try:
                        pg_url = (f"https://api.bricklink.com/api/store/v1/items/part/{pid}/price"
                                  f"?color_id={cid}&guide_type=sold&new_or_used=U")
                        pg = req.get(pg_url, auth=auth, timeout=6)
                        if pg.status_code == 200:
                            d = pg.json().get("data", {})
                            total_qty = d.get("total_quantity", 0) or 0
                            color_sales.append((cid, total_qty))
                        else:
                            color_sales.append((cid, 0))
                    except Exception:
                        color_sales.append((cid, 0))

                # Sort by total sales qty descending, take top 9
                color_sales.sort(key=lambda x: x[1], reverse=True)
                top9 = color_sales[:9]

                if not top9:
                    self.log_message.emit(f"⚠  No sales data found for {pid}", "warning")
                    return

                # Build clone rows
                clones = []
                for cid, qty_sold in top9:
                    if qty_sold == 0:
                        continue  # skip colors with zero recorded sales
                    cname, crgb = ("—", (128,128,128))
                    if cid in BL_COLORS:
                        cname = BL_COLORS[cid][0]
                        crgb  = BL_COLORS[cid][1]
                    import copy
                    clone = copy.copy(rd)
                    clone["color_id"]     = cid
                    clone["color_name"]   = cname
                    clone["color_rgb"]    = crgb
                    clone["qty"]          = 1
                    clone["price"]        = None
                    clone["from_color_clone"] = True
                    clone["_sales_qty"]   = qty_sold
                    clones.append(clone)

                if not clones:
                    self.log_message.emit("⚠  All top colors have zero sales — skipping", "warning")
                    return

                self.log_message.emit(
                    f"🎨  {pid} → {len(clones)} color clones added", "success")
                self._clone_colors_ready.emit(rd, clones)

            except Exception as e:
                self.log_message.emit(f"⚠  Color clone failed: {e}", "warning")

        threading.Thread(target=fetch_and_clone, daemon=True).start()

    def _on_clone_colors_ready(self, rd, clones):
        self._save_undo_snapshot("clone colors")
        """Main thread: append cloned rows for each top color."""
        for clone in clones:
            clone["index"] = self._part_count + 1
            self._part_count += 1
            self._add_result_row(clone)

        self._update_total_value()
        self.results_count.setText(f"{self._part_count} parts")
        self._log(f"🎨  {len(clones)} color rows added — set qty for each", "success")

    # Fields that can be undone — mutable data only, no widget refs
    _UNDO_FIELDS = ("qty","price","medium_price","color_id","color_name","color_rgb",
                    "condition","remark","comment","part_id","part_name","item_type",
                    "_deleted","known_color_ids","color_method",
                    "_pre_scan_color_id","_pre_scan_color_name","_pre_scan_color_rgb")

    def _cycle_theme(self):
        """Cycle through Dark → Light → LEGO themes and reapply stylesheet."""
        order = ["dark", "light", "lego"]
        labels = {"dark": "🌙", "light": "☀", "lego": "🟡"}
        names  = {"dark": "Dark", "light": "Light", "lego": "LEGO"}
        cur = getattr(self, "_current_theme", "dark")
        nxt = order[(order.index(cur) + 1) % len(order)]
        self._current_theme = nxt

        # Apply new globals
        _apply_theme(nxt)
        global DARK_BG, PANEL_BG, CARD_BG, BORDER, ACCENT, ACCENT2
        global TEXT, TEXT_DIM, SUCCESS, WARNING, CONSOLE_BG, ROW_ALT, ROW_SEL

        # Save preference
        try:
            cfg = {}
            if CFG_PATH.exists():
                try: cfg = json.loads(CFG_PATH.read_text())
                except Exception: pass
            cfg["theme"] = nxt
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception: pass

        # Update button label
        self.theme_btn.setText(labels[nxt])

        # Rebuild the application stylesheet
        self._apply_app_stylesheet()
        self._log(f"🎨  Theme → {names[nxt]}", "info")

    def _apply_app_stylesheet(self):
        """Reapply the full app stylesheet with current theme colors."""
        from PyQt5.QtWidgets import QApplication
        t = THEMES[self._current_theme]
        bg    = t["DARK_BG"]
        panel = t["PANEL_BG"]
        card  = t["CARD_BG"]
        bdr   = t["BORDER"]
        txt   = t["TEXT"]
        dim   = t["TEXT_DIM"]
        acc   = t["ACCENT"]
        acc2  = t["ACCENT2"]
        ok    = t["SUCCESS"]
        warn  = t["WARNING"]
        con   = t["CONSOLE_BG"]

        QApplication.instance().setStyleSheet(f"""
            QWidget {{
                background: {bg};
                color: {txt};
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }}
            QScrollArea, QScrollBar {{
                background: {bg};
                border: none;
            }}
            QScrollBar:vertical {{
                background: {panel};
                width: 8px;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background: {bdr};
                border-radius: 4px;
                min-height: 20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
            QPushButton {{
                background: {card};
                color: {txt};
                border: 1px solid {bdr};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
            }}
            QPushButton:hover {{ background: {bdr}; }}
            QPushButton:disabled {{ color: {dim}; }}
            QComboBox {{
                background: {card};
                color: {txt};
                border: 1px solid {bdr};
                border-radius: 3px;
                padding: 2px 4px;
            }}
            QComboBox QAbstractItemView {{
                background: {panel};
                color: {txt};
                selection-background-color: {acc2};
            }}
            QSpinBox, QLineEdit, QTextEdit {{
                background: {card};
                color: {txt};
                border: 1px solid {bdr};
                border-radius: 3px;
                padding: 2px 4px;
            }}
            QCheckBox {{ color: {txt}; }}
            QLabel {{ background: transparent; color: {txt}; }}
            QSplitter::handle {{ background: {bdr}; }}
            QMenu {{
                background: {card};
                color: {txt};
                border: 1px solid {bdr};
            }}
            QMenu::item:selected {{ background: {acc2}; color: #000; }}
            QMenu::separator {{ height: 1px; background: {bdr}; margin: 3px 6px; }}
            QListWidget {{
                background: {card};
                color: {txt};
                border: 1px solid {bdr};
            }}
            QListWidget::item:selected {{ background: {acc2}; color: #000; }}
            QDialog {{ background: {bg}; color: {txt}; }}
            QToolTip {{
                background-color: #1a1a1a;
                color: #f0f0f0;
                border: 1px solid #555555;
                padding: 4px 6px;
                font-size: 11px;
                border-radius: 3px;
            }}
        """)

    def _set_batch_scan_mode(self, label):
        """Turn SCAN button orange with batch label so user knows they're in batch mode."""
        self.scan_btn.setText(f"▶  SCAN  —  {label}")
        self.scan_btn.setStyleSheet("""
            QPushButton { background: #7a4400; color: white; font-size: 16px;
                          font-weight: bold; border-radius: 10px; border: none; }
            QPushButton:hover { background: #a05800; }
        """)

    def _reset_scan_btn(self):
        """Restore SCAN button to default green."""
        self.scan_btn.setText("▶  SCAN")
        self.scan_btn.setStyleSheet(f"""
            QPushButton {{background:{ACCENT};color:white;font-size:22px;font-weight:bold;
                          border-radius:10px;border:none;}}
            QPushButton:hover {{background:#ff6b7a;}}
            QPushButton:disabled {{background:#444;color:#777;}}
        """)

    def _toggle_auto_merge(self):
        self._auto_merge_on = not getattr(self, "_auto_merge_on", True)
        if self._auto_merge_on:
            self.auto_merge_btn.setText("🔀 Auto-merge ✓")
            self.auto_merge_btn.setStyleSheet(
                "background:#1a4a1a;color:white;font-size:13px;font-weight:bold;"
                "border-radius:6px;border:none;")
            self.auto_merge_btn.setToolTip("Auto-merge ON — click to disable")
        else:
            self.auto_merge_btn.setText("🔀 Auto-merge ✗")
            self.auto_merge_btn.setStyleSheet(
                "background:#2a2a2a;color:#666;font-size:13px;font-weight:bold;"
                "border-radius:6px;border:1px solid #444;")
            self.auto_merge_btn.setToolTip("Auto-merge OFF — click to enable")

    def _on_detected_count(self, n):
        self.results_count.setText(f"{self._part_count} parts  |  🔍 {n} on mat")

    def _scan_or_cancel(self):
        """SCAN button toggles: starts scan or cancels if already scanning."""
        if self.scan_btn.text().startswith("⏳") or "CANCEL" in self.scan_btn.text():
            self._cancel_scan()
        else:
            self._start_scan()

    def _toggle_results_fullscreen(self):
        self._results_fullscreen = not self._results_fullscreen
        w0 = self._main_splitter.widget(0)
        if self._results_fullscreen:
            w0.setMinimumWidth(0)
            w0.setMaximumWidth(0)
            self._main_splitter.setSizes([0, 99999])
            sizes = self._pg_splitter.sizes()
            self._pg_splitter.setSizes([sum(sizes), 0, 0, 0])
            self.fullscreen_btn.setText("✕⛶")
            self.fullscreen_btn.setToolTip("Exit fullscreen (click again)")
        else:
            w0.setMinimumWidth(220)
            # Usability: don't re-impose a narrow max width when exiting fullscreen.
            w0.setMaximumWidth(9999)
            self._main_splitter.setSizes([440, 1100])
            self._pg_splitter.setSizes([480, 180, 100, 0])
            self.fullscreen_btn.setText("⛶")
            self.fullscreen_btn.setToolTip("Toggle fullscreen results panel")

    def _save_undo_snapshot(self, action=""):
        """Save mutable field values for all rows — one snapshot per user operation."""
        snap = {
            "action": action,
            "count":  len(self._rows),
            "fields": [
                {f: rd.get(f) for f in self._UNDO_FIELDS}
                for rd in self._rows
            ],
        }
        self._undo_stack.append(snap)
        if len(self._undo_stack) > getattr(self, "_UNDO_MAX", 25):
            self._undo_stack = self._undo_stack[-getattr(self, "_UNDO_MAX", 25):]
        tip = f"↩ Undo: {snap.get('action')}" if snap.get("action") else "↩ Undo"
        self._bulk_undo.setToolTip(tip)
        self._bulk_undo.setEnabled(True)

    def _undo_last(self):
        """Restore mutable fields in-place. Rows added since snapshot are hidden."""
        if not getattr(self, "_undo_stack", None):
            self._log("Nothing to undo", "info")
            return
        snap = self._undo_stack.pop()
        snap_fields = snap["fields"]
        snap_count  = snap["count"]

        # Rows that existed at snapshot time — restore their fields in-place
        for i, rd in enumerate(self._rows[:snap_count]):
            if i >= len(snap_fields): break
            sf = snap_fields[i]
            for f in self._UNDO_FIELDS:
                if f in sf:
                    rd[f] = sf[f]
            # Sync widgets to restored values
            if rd.get("_qty_lbl"):
                rd["_qty_lbl"].setText(str(rd.get("qty", 1)))
            if rd.get("_price_lbl"):
                p = rd.get("price")
                rd["_price_lbl"].setText(f"${p:.2f}" if p else "—")
            if rd.get("_color_lbl"):
                rd["_color_lbl"].setText((rd.get("color_name") or "—")[:13])
            if rd.get("_swatch"):
                rgb = rd.get("color_rgb", (128,128,128))
                rd["_swatch"].setStyleSheet(
                    f"background:rgb{tuple(rgb)};border-radius:3px;border:1px solid #444;")
            if rd.get("_cond_btn"):
                c = rd.get("condition","U")
                rd["_cond_btn"].setText(c)
            if rd.get("_remark_lbl"):
                v = rd.get("remark","") or ""
                rd["_remark_lbl"].setText(v if v else "remark…")
            # Show/hide based on _deleted flag
            deleted = rd.get("_deleted", False)
            if rd.get("_widget"):
                rd["_widget"].setVisible(not deleted)

        # Rows added AFTER the snapshot — hide them (they didn't exist before)
        for rd in self._rows[snap_count:]:
            if rd.get("_widget"):
                rd["_widget"].setVisible(False)
            rd["_deleted"] = True

        self._part_count = sum(1 for r in self._rows if not r.get("_deleted"))
        self.results_count.setText(f"{self._part_count} parts")
        self._update_total_value()
        if self._undo_stack:
            prev = self._undo_stack[-1]
            tip = f"↩ Undo: {prev.get('action')}" if prev.get("action") else "↩ Undo"
            self._bulk_undo.setToolTip(tip)
            self._bulk_undo.setEnabled(True)
        else:
            self._bulk_undo.setToolTip("↩ Undo (nothing more to undo)")
            self._bulk_undo.setEnabled(False)
        self._log(f"↩  Undid: {snap.get('action')}", "info")


    # ── Weight Scale Methods ─────────────────────────────────────────────────

    def _wt_connect_scale(self):
        """Try to connect to USB scale via serial port."""
        try:
            import serial.tools.list_ports as lp
            ports = [p.device for p in lp.comports()]
        except ImportError:
            self._wt_status.setText("pyserial not installed")
            self._wt_status.setStyleSheet(f"font-size:10px;color:{ACCENT};")
            self._log("⚖  pip install pyserial to use USB scale", "warning")
            return

        if not ports:
            self._wt_status.setText("No serial ports found")
            return

        # Try each port looking for a scale
        import serial
        for port in ports:
            try:
                s = serial.Serial(port, baudrate=9600, timeout=1)
                line = s.readline().decode("ascii", errors="ignore").strip()
                if any(c.isdigit() for c in line):
                    self._wt_serial = s
                    self._wt_status.setText(f"✓ {port}")
                    self._wt_status.setStyleSheet(f"font-size:10px;color:{SUCCESS};")
                    self._wt_con_btn.setText("Disconnect")
                    self._wt_con_btn.clicked.disconnect()
                    self._wt_con_btn.clicked.connect(self._wt_disconnect_scale)
                    self._wt_start_reading()
                    self._log(f"⚖  Scale connected on {port}", "success")
                    return
                s.close()
            except Exception:
                pass
        self._wt_status.setText("No scale found on any port")
        self._log("⚖  No scale found — check USB connection", "warning")

    def _wt_disconnect_scale(self):
        if self._wt_serial:
            try: self._wt_serial.close()
            except Exception: pass
            self._wt_serial = None
        self._wt_status.setText("Disconnected")
        self._wt_status.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        self._wt_con_btn.setText("Connect")
        self._wt_con_btn.clicked.disconnect()
        self._wt_con_btn.clicked.connect(self._wt_connect_scale)
        self._wt_live.setText("— g")

    def _wt_start_reading(self):
        """Background thread: read weight from serial, emit to GUI."""
        def read_loop():
            import serial
            while self._wt_serial and self._wt_serial.is_open:
                try:
                    line = self._wt_serial.readline().decode("ascii", errors="ignore").strip()
                    # Parse grams from common scale formats: "123.4 g", "ST,GS,+  123.4g", etc.
                    import re
                    m = re.search(r"([+-]?\d+[.]?\d*)\s*g", line, re.IGNORECASE)
                    if not m:
                        m = re.search(r"([+-]?\d+\.?\d*)", line)
                    if m:
                        g = float(m.group(1))
                        self._wt_raw_signal.emit(g)
                except Exception:
                    pass
        threading.Thread(target=read_loop, daemon=True).start()

    def _wt_on_raw(self, g):
        """Main thread: update weight display and calculate qty."""
        self._wt_raw_g = g
        net = max(0.0, g - self._wt_tare_val)
        self._wt_live.setText(f"{net:.1f} g")
        self._wt_calculate(net)

    def _wt_tare(self):
        self._wt_tare_val = self._wt_raw_g
        self._wt_live.setText("0.0 g")
        self._log(f"⚖  Tare set ({self._wt_raw_g:.1f}g)", "info")

    def _wt_on_unit_ready(self, unit_g):
        """Main thread: unit weight fetched — update label and recalculate."""
        self._wt_unit_lbl.setText(f"unit: {unit_g:.3f}g")
        self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{SUCCESS};")
        self._wt_recalculate_from_inputs()

    def _wt_recalculate_from_inputs(self):
        if self._wt_raw_g > 0:
            net = max(0.0, self._wt_raw_g - self._wt_tare_val)
        else:
            try:
                net = float(self._wt_manual.text().replace(",", "."))
            except (ValueError, AttributeError):
                net = 0.0
        self._wt_calculate(net)

    def _wt_manual_changed(self, text):
        try:
            g = float(text.replace(",", "."))
            self._wt_live.setText(f"{g:.1f} g (manual)")
            # Only calculate if we have a part selected
            manual_pid = getattr(self, "_wt_pid_edit", None)
            has_pid = (manual_pid and manual_pid.text().strip()) or self.weight_part_id.currentData()
            if has_pid:
                self._wt_calculate(g)
        except ValueError:
            self._wt_result.setText("Qty: —")
            if hasattr(self, "_wt_apply_btn"): self._wt_apply_btn.setEnabled(False)

    def _wt_calculate(self, net_g):
        manual_pid = getattr(self, "_wt_pid_edit", None)
        pid = (manual_pid.text().strip() if manual_pid and manual_pid.text().strip()
               else self.weight_part_id.currentData())
        if not pid or net_g <= 0:
            self._wt_result.setText("Qty: —")
            self._wt_apply_btn.setEnabled(False)
            return
        # Check if user typed a unit weight manually in the unit label
        unit_g = None
        try:
            unit_txt = self._wt_unit_lbl.text().replace("unit:", "").replace("g","").strip()
            if unit_txt and unit_txt not in ("—", "unknown"):
                unit_g = float(unit_txt)
        except (ValueError, AttributeError):
            pass
        if not unit_g:
            unit_g = self._wt_get_unit_weight(pid)
        if not unit_g:
            self._wt_unit_lbl.setText("unit: — (enter g/part)")
            self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{WARNING};")
            self._wt_unit_lbl.setToolTip("No weight data found — click and type unit weight in grams")
            self._wt_unit_lbl.setCursor(Qt.PointingHandCursor)
            # Make unit label editable on click
            def _edit_unit(e, pid_=pid):
                from PyQt5.QtWidgets import QInputDialog
                val, ok = QInputDialog.getDouble(self, "Unit Weight",
                    f"Weight per {pid_} (grams):", 1.0, 0.01, 999.0, 3)
                if ok:
                    if not hasattr(self, "_wt_weight_cache"):
                        self._wt_weight_cache = {}
                    self._wt_weight_cache[pid_.lower()] = val
                    self._wt_unit_lbl.setText(f"unit: {val:.2f}g")
                    self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{SUCCESS};")
                    self._wt_calculate(net_g)
            self._wt_unit_lbl.mousePressEvent = _edit_unit
            self._wt_result.setText("Qty: ?  — click unit weight")
            self._wt_apply_btn.setEnabled(False)
            return
        self._wt_unit_lbl.setText(f"unit: {unit_g:.2f}g")
        self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")
        qty = max(1, round(net_g / unit_g))
        self._wt_qty_result = qty
        self._wt_result.setText(f"Qty: {qty}")
        self._wt_result.setStyleSheet(f"font-size:13px;font-weight:bold;color:{SUCCESS};")
        self._wt_apply_btn.setEnabled(True)

    def _wt_part_selected(self, index):
        """When user picks a part from combo, fetch unit weight via BrickLink API."""
        pid = self.weight_part_id.currentData()
        if not pid:
            return
        # Already cached
        if hasattr(self, "_wt_weight_cache") and pid.lower() in self._wt_weight_cache:
            self._wt_recalculate_from_inputs()
            return
        self._wt_unit_lbl.setText("unit: fetching…")
        self._wt_unit_lbl.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};")

        def fetch():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    return
                from requests_oauthlib import OAuth1
                import requests
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])
                r = requests.get(
                    f"https://api.bricklink.com/api/store/v1/items/part/{pid}",
                    auth=auth, timeout=6)
                if r.status_code == 200:
                    data = r.json().get("data", {})
                    w = data.get("weight")
                    if w and float(w) > 0:
                        if not hasattr(self, "_wt_weight_cache"):
                            self._wt_weight_cache = {}
                        self._wt_weight_cache[pid.lower()] = float(w)
                        self.log_message.emit(f"⚖  {pid}: {float(w):.3f}g/part", "info")
                        self._wt_unit_ready.emit(float(w))
                        return
            except Exception:
                pass
            self.log_message.emit(f"⚖  No weight for {pid} — click unit label to enter manually", "info")

        threading.Thread(target=fetch, daemon=True).start()

    def _wt_get_unit_weight(self, part_id):
        """Get part weight from manual cache only — no file I/O, no network on main thread."""
        if not hasattr(self, "_wt_weight_cache"):
            self._wt_weight_cache = {}
        pid = part_id.lower().strip()
        if pid in self._wt_weight_cache:
            return self._wt_weight_cache[pid]
        # Check material from parts.csv — rubber parts weigh differently but no weight data available
        # User must enter unit weight manually via click on unit label
        return None

    def _wt_apply_qty(self):
        """Apply calculated qty to all selected rows with matching part_id."""
        qty = getattr(self, "_wt_qty_result", 0)
        if not qty or qty <= 0: return
        manual_pid = getattr(self, "_wt_pid_edit", None)
        pid = (manual_pid.text().strip() if manual_pid and manual_pid.text().strip()
               else self.weight_part_id.currentData())
        if not pid: return
        sel = self._selected_rows()
        targets = [r for r in sel if r.get("part_id","").lower() == pid.lower()] if sel else                   [r for r in self._rows if r.get("part_id","").lower() == pid.lower()
                   and not r.get("_deleted")]
        for rd in targets:
            rd["qty"] = qty
            if rd.get("_qty_lbl"): rd["_qty_lbl"].setText(str(qty))
        self._update_total_value()
        self._log(f"⚖  {pid} qty set to {qty} ({len(targets)} row(s))", "success")


    # ── Remark Templates ─────────────────────────────────────────────────────

    _CUSTOM_PRESETS_FILE = Path("custom_presets.json")

    def _load_custom_presets(self):
        """Load user-added comment presets from disk."""
        try:
            if self._CUSTOM_PRESETS_FILE.exists():
                import json as _j
                return _j.loads(self._CUSTOM_PRESETS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
        return []

    def _save_custom_presets(self, presets):
        """Save user-added comment presets to disk."""
        try:
            import json as _j
            self._CUSTOM_PRESETS_FILE.write_text(
                _j.dumps(presets, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    _REMARK_TEMPLATES = [
        ("Condition",  [
            ("Like new",              "Like new"),
            ("Slight discoloration",  "Slight discoloration"),
            ("Discolored",            "Discolored"),
            ("Heavy discoloration",   "Heavy discoloration"),
            ("Scratched",             "Scratched"),
            ("Print wear",            "Print wear"),
            ("Sticker residue",       "Sticker residue"),
            ("Sticker applied",       "Sticker applied"),
            ("Cracked",               "Cracked"),
            ("Cracked right side",    "Cracked right side"),
            ("Cracked left side",     "Cracked left side"),
            ("Broken clip",           "Broken clip"),
        ]),
        ("Lot notes",  [
            ("Mixed condition",       "Mixed condition"),
            ("Qty approximate",       "Qty approximate"),
            ("Weighed quantity",      "Weighed quantity"),
            ("New from set",          "New from set"),
            ("From bulk lot",         "From bulk lot"),
        ]),
        ("Extras",  [
            ("Cleaned",               "Cleaned"),
            ("Uncleaned",             "Uncleaned"),
            ("Tested",                "Tested"),
        ]),
    ]

    def _show_remark_templates(self, rd):
        """Pop a quick template picker near the remark label."""
        from PyQt5.QtWidgets import QMenu, QAction, QWidgetAction, QLabel as _QL2
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};}}"
            f"QMenu::item{{padding:5px 18px 5px 10px;}}"
            f"QMenu::item:selected{{background:{ACCENT2};color:#000;}}"
            f"QMenu::separator{{height:1px;background:{BORDER};margin:3px 6px;}}"
        )

        def _apply(text):
            rd["comment"] = text
            lbl = rd.get("_cmt_lbl")
            if lbl:
                lbl.setText(text)
                lbl.setStyleSheet(f"color:{TEXT};font-size:10px;font-style:italic;"
                                   f"background:transparent;padding:0 4px;")

        # Custom remark at top
        custom_act  = menu.addAction("✏  Custom comment…")
        add_preset  = menu.addAction("➕  Save as preset…")
        mgr_act     = menu.addAction("🗂  Manage presets…")
        menu.addSeparator()

        for group_name, items in self._REMARK_TEMPLATES:
            hdr = menu.addAction(f"── {group_name} ──")
            hdr.setEnabled(False)
            for label, value in items:
                act = menu.addAction(label)
                act.triggered.connect(lambda _, v=value: _apply(v))
            menu.addSeparator()

        # User custom presets section
        custom_presets = self._load_custom_presets()
        if custom_presets:
            hdr2 = menu.addAction("── My presets ──")
            hdr2.setEnabled(False)
            for cp in custom_presets:
                act = menu.addAction(cp)
                act.triggered.connect(lambda _, v=cp: _apply(v))
            menu.addSeparator()

        # Clear option
        clear_act = menu.addAction("✕  Clear comment")

        # Show near cursor
        from PyQt5.QtGui import QCursor
        chosen = menu.exec_(QCursor.pos())

        if chosen == custom_act:
            from PyQt5.QtWidgets import QInputDialog
            cur = rd.get("comment", "") or ""
            val, ok = QInputDialog.getText(self, "Custom Comment", "Comment text:", text=cur)
            if ok:
                _apply(val.strip())
        elif chosen == add_preset:
            from PyQt5.QtWidgets import QInputDialog
            cur = rd.get("comment", "") or ""
            val, ok = QInputDialog.getText(self, "Save Preset",
                "Preset text to save:", text=cur)
            if ok and val.strip():
                presets = self._load_custom_presets()
                if val.strip() not in presets:
                    presets.append(val.strip())
                    self._save_custom_presets(presets)
                    self._log(f"📋  Preset saved: {val.strip()}", "success")
        elif chosen == mgr_act:
            self._manage_presets()
        elif chosen == clear_act:
            _apply("")
            lbl = rd.get("_cmt_lbl")
            if lbl:
                lbl.setText("comment…")
                lbl.setStyleSheet(f"color:#aaa;font-size:10px;font-style:italic;"
                                   f"background:transparent;padding:0 4px;")

    def _manage_presets(self):
        """Dialog to view and delete custom presets."""
        from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                      QListWidget, QPushButton, QLabel)
        presets = self._load_custom_presets()
        dlg = QDialog(self)
        dlg.setWindowTitle("Manage Custom Presets")
        dlg.setMinimumWidth(340)
        dlg.setStyleSheet(f"background:{CARD_BG};color:{TEXT};")
        vl = QVBoxLayout(dlg)
        vl.addWidget(QLabel("Your saved presets:"))
        lst = QListWidget()
        lst.setStyleSheet(f"background:{DARK_BG};color:{TEXT};border:1px solid {BORDER};")
        for p in presets:
            lst.addItem(p)
        vl.addWidget(lst)
        btns = QHBoxLayout()
        del_btn = QPushButton("🗑 Delete selected")
        del_btn.setStyleSheet(f"background:{ACCENT};color:#000;padding:4px 8px;")
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(f"background:{CARD_BG};color:{TEXT};padding:4px 8px;")
        btns.addWidget(del_btn); btns.addWidget(close_btn)
        vl.addLayout(btns)
        def _del():
            row = lst.currentRow()
            if row >= 0:
                presets.pop(row)
                lst.takeItem(row)
                self._save_custom_presets(presets)
        del_btn.clicked.connect(_del)
        close_btn.clicked.connect(dlg.accept)
        dlg.exec_()

    # ── Session History ───────────────────────────────────────────────────────

    def _show_sessions(self):
        """Show recent scan sessions from the reports/ folder."""
        from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                      QLabel, QListWidget, QListWidgetItem,
                                      QPushButton)
        from PyQt5.QtCore import Qt

        reports_dir = Path("reports")
        if not reports_dir.exists():
            self._log("No reports folder found yet", "warning"); return

        # Find all sessions — each subdir with a scan-results.json
        sessions = []
        for d in sorted(reports_dir.iterdir(), reverse=True):
            json_p = d / "scan-results.json"
            if not json_p.exists(): continue
            try:
                data = json.loads(json_p.read_text(encoding="utf-8"))
                n_parts = len(data)
                # Try to compute total value
                def _qty(row, default=1):
                    try:
                        v = row.get("qty", default)
                        if v is None or v == "":
                            return int(default)
                        return int(v)
                    except Exception:
                        return int(default)
                total = sum(
                    (r.get("price") or r.get("medium_price") or 0) * _qty(r, 1)
                    for r in data
                )
                # Parse date from folder name e.g. scan-20260314-141917
                name = d.name
                date_str = ""
                import re
                m = re.search(r"(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})", name)
                if m:
                    date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}  {m.group(4)}:{m.group(5)}"
                sessions.append((d, date_str or name, n_parts, total, data))
            except Exception:
                continue

        if not sessions:
            self._log("No scan sessions found in reports/", "info"); return

        dlg = QDialog(self)
        dlg.setWindowTitle("🕐  Recent Sessions")
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
        dlg.resize(460, 380)
        vl = QVBoxLayout(dlg); vl.setSpacing(8); vl.setContentsMargins(14,14,14,14)

        hdr = QLabel("Select a session to load into the results panel:")
        hdr.setStyleSheet(f"font-size:11px;color:{TEXT_DIM};")
        vl.addWidget(hdr)

        lst = QListWidget()
        lst.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};"
                           f"font-size:11px;")
        lst.setAlternatingRowColors(True)

        for d, date_str, n_parts, total, data in sessions[:20]:
            total_str = f"  ${total:.2f}" if total > 0 else ""
            item = QListWidgetItem(f"  {date_str}    {n_parts} parts{total_str}")
            item.setData(Qt.UserRole, (d, data))
            lst.addItem(item)

        vl.addWidget(lst)

        note = QLabel("⚠  Loading a session will ADD rows to current results.")
        note.setStyleSheet(f"font-size:10px;color:{ACCENT2};font-style:italic;")
        vl.addWidget(note)

        btn_row = QHBoxLayout()
        load_btn  = self._btn("📂  Load session", SUCCESS,  dlg.accept,  w=140)
        clear_btn = self._btn("🗑  Clear & load", ACCENT,   lambda: (setattr(dlg, '_clear', True), dlg.accept()), w=120)
        cxl_btn   = self._btn("Cancel",           CARD_BG,  dlg.reject,  w=80)
        load_btn.setFixedHeight(28); clear_btn.setFixedHeight(28); cxl_btn.setFixedHeight(28)
        btn_row.addWidget(clear_btn); btn_row.addStretch()
        btn_row.addWidget(cxl_btn); btn_row.addWidget(load_btn)
        vl.addLayout(btn_row)

        lst.itemDoubleClicked.connect(lambda _: dlg.accept())

        if dlg.exec_() != QDialog.Accepted: return
        sel = lst.currentItem()
        if not sel: return

        d, data = sel.data(Qt.UserRole)
        if getattr(dlg, '_clear', False):
            self._clear_results()

        # Load rows from JSON
        loaded = 0
        for r in data:
            if not r.get("part_id"): continue
            cid = r.get("color_id", 0)
            if cid and cid in BL_COLORS:
                r["color_name"] = BL_COLORS[cid][0]
                r["color_rgb"]  = BL_COLORS[cid][1]
            r["index"] = self._part_count + 1
            self._part_count += 1
            self._add_result_row(r)
            loaded += 1

        self._update_total_value()
        self.results_count.setText(f"{self._part_count} parts")
        xml_p = d / "scan-results.xml"
        if xml_p.exists():
            self.last_xml_path = str(xml_p)
            self.copy_btn.setEnabled(True)
            self.clip_btn.setEnabled(True)
            self.merge_btn.setEnabled(True)
            self.upload_btn.setEnabled(True)
            self.bs_btn.setEnabled(True)
            self.merge_r2.setEnabled(True)
        self._log(f"🕐  Loaded {loaded} parts from {d.name}", "success")

    def _save_grid_prefs(self):
        try:
            cfg = {}
            if CFG_PATH.exists():
                try: cfg = json.loads(CFG_PATH.read_text())
                except Exception: pass
            cfg["grid_enabled"] = self.grid_check.isChecked()
            cfg["grid_cols"]    = self.cols_spin.value()
            cfg["grid_rows"]    = self.rows_spin.value()
            CFG_PATH.write_text(json.dumps(cfg, indent=2))
        except Exception: pass

    def _load_grid_prefs(self):
        try:
            if not CFG_PATH.exists(): return
            cfg = json.loads(CFG_PATH.read_text())
            if "grid_enabled" in cfg:
                self.grid_check.setChecked(bool(cfg["grid_enabled"]))
            if "grid_cols" in cfg:
                self.cols_spin.setValue(int(cfg["grid_cols"]))
            if "grid_rows" in cfg:
                self.rows_spin.setValue(int(cfg["grid_rows"]))
        except Exception: pass

    def _fetch_po_value(self, rd):
        """Background: fetch subsets + median prices to estimate part-out value."""
        pid = rd.get("part_id", "")
        if not pid: return
        currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"

        def fetch():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4: return
                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"], bl["TOKEN"], bl["TOKEN_SECRET"])
                # Step 1: get subsets
                r = req.get(f"https://api.bricklink.com/api/store/v1/items/minifig/{pid}/subsets",
                            auth=auth, timeout=10)
                if r.status_code != 200: return
                total = 0.0
                for match in r.json().get("data", []):
                    for entry in match.get("entries", []):
                        item = entry.get("item", {})
                        if item.get("type") not in ("PART", "MINIFIG"): continue
                        sub_pid = item.get("no", "")
                        cid     = entry.get("color_id", 0)
                        qty     = entry.get("quantity", 1)
                        itype   = "minifig" if item.get("type") == "MINIFIG" else "part"
                        # Step 2: get median price for each part
                        try:
                            pr = req.get(
                                f"https://api.bricklink.com/api/store/v1/items/{itype}/{sub_pid}/price",
                                params={"guide_type":"sold","new_or_used":"U",
                                        "currency_code":currency,"color_id":cid},
                                auth=auth, timeout=6)
                            if pr.status_code == 200:
                                avg = pr.json().get("data", {}).get("avg_price") or                                       pr.json().get("data", {}).get("qty_avg_price")
                                if avg and float(avg) > 0:
                                    total += float(avg) * qty
                        except Exception:
                            pass
                if total > 0:
                    self._po_value_ready.emit(rd, total)
            except Exception:
                pass

        threading.Thread(target=fetch, daemon=True).start()

    def _on_po_value_ready(self, rd, total):
        """Main thread: update part-out value label on the minifig row."""
        lbl = rd.get("_po_val_lbl")
        if not lbl: return
        median = rd.get("medium_price") or rd.get("price") or 0
        diff   = total - median
        color  = SUCCESS if diff > 0 else ACCENT
        arrow  = "▲" if diff > 0 else "▼"
        lbl.setText(f"⊞ ${total:.2f} {arrow}{abs(diff):.2f}")
        lbl.setStyleSheet(f"font-size:9px;color:{color};background:transparent;padding:0 2px;font-weight:bold;")
        lbl.setToolTip(
            f"Part-out value: ${total:.2f}\n"
            f"Minifig median: ${median:.2f}\n"
            f"Difference: {arrow}${abs(diff):.2f} — "
            f"{'Part out is MORE profitable' if diff > 0 else 'Sell whole fig is better'}"
        )

    def _rescan_region(self, rd, expand=0.0, require_min_parts: int = 1, use_watershed: bool = False):
        """Re-scan a box region to split merged parts."""
        source    = rd.get("source_image", "")
        bbox      = rd.get("bbox")
        crop_path = rd.get("crop_image", "")

        scan_path = None
        # Always cut from source when we have bbox (allows expanding beyond crop)
        if source and Path(source).exists() and bbox:
            try:
                from PIL import Image as _PIL
                img = _PIL.open(source)
                x1, y1, x2, y2 = bbox
                w_src, h_src = img.size
                bw, bh = x2 - x1, y2 - y1
                # Base margin + optional expansion
                base_margin = 20
                extra_x = int(bw * expand)
                extra_y = int(bh * expand)
                margin_x = base_margin + extra_x
                margin_y = base_margin + extra_y
                region = img.crop((max(0, x1-margin_x), max(0, y1-margin_y),
                                   min(w_src, x2+margin_x), min(h_src, y2+margin_y)))
                import tempfile
                tmp = tempfile.NamedTemporaryFile(suffix=".jpg", prefix="rescan-", delete=False)
                region.save(tmp.name, quality=95)
                tmp.close()
                scan_path = tmp.name
                if expand > 0:
                    self._log(f"🔍+  Expanding crop by {int(expand*100)}% — was {bw}×{bh}px", "info")
            except Exception as e:
                self._log(f"⚠  Re-scan crop failed: {e}", "warning"); return
        elif crop_path and Path(crop_path).exists() and expand == 0:
            scan_path = crop_path
        else:
            self._log("⚠  No source image with bbox for re-scan — try setting background color first", "warning"); return

        if require_min_parts > 1:
            label = f"🪓  Split attempt (need ≥{require_min_parts} parts) for {rd.get('part_id','?')}"
        else:
            label = f"🔍  Re-scanning region for {rd.get('part_id','?')}"
        if expand > 0:
            label = "🪓+  " + label
        self._log(label + "...", "info")

        # Find position of this row to insert results after it
        try:
            insert_idx = next(i for i,r in enumerate(self._rows) if r is rd)
        except StopIteration:
            return

        # Find layout position
        layout = self.results_list_layout
        layout_pos = -1
        src_widget = rd.get("_widget")
        for i in range(layout.count()):
            item = layout.itemAt(i)
            if item and item.widget() is src_widget:
                layout_pos = i + 1  # insert AFTER this row
                break

        def run_rescan():
            try:
                import subprocess, json
                settings = {
                    "gap":        4,   # tight gap to split merged blobs
                    "padding":    self.padding_spin.value(),
                    "bg_color":   getattr(self, "_bg_rgb", None),
                    "shadow_color": getattr(self, "_shadow_rgb", None),
                    "confidence": self.conf_slider.value() / 100.0,
                    "currency":   self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD",
                }
                rescan_gap = 4 if expand == 0 else max(8, settings["gap"])
                args = [_python_exe(), "scan-heads.py", scan_path,
                        "--mode", "all",
                        "--confidence", str(settings["confidence"]),
                        "--gap", str(rescan_gap),
                        "--padding", str(settings["padding"]),
                        "--currency", settings["currency"],
                        "--concurrency", "4",
                ]
                if use_watershed:
                    args += ["--watershed"]
                if settings["bg_color"]:
                    r2,g2,b2 = settings["bg_color"]
                    args += ["--bg-color", f"{r2},{g2},{b2}"]
                if settings["shadow_color"]:
                    r2,g2,b2 = settings["shadow_color"]
                    args += ["--shadow-color", f"{r2},{g2},{b2}"]

                env = os.environ.copy(); env["PYTHONIOENCODING"] = "utf-8"
                proc = subprocess.run(args, capture_output=True, text=True,
                                      encoding="utf-8", errors="replace",
                                      env=env, cwd=str(Path.cwd()),
                                      **_hidden_popen_kwargs())

                # Parse JSON results — find output dir from stdout
                import re as _re
                out_match = _re.search(r"Output: (.+)", proc.stdout)
                if not out_match:
                    self.log_message.emit("⚠  Re-scan: no output dir in stdout", "warning"); return
                json_p = Path(out_match.group(1).strip()) / "scan-results.json"

                if not json_p.exists():
                    self.log_message.emit(f"⚠  Re-scan JSON not found: {json_p}", "warning"); return

                entries = json.loads(json_p.read_text())
                if len(entries) <= 1:
                    self.log_message.emit(f"🔍  Re-scan found {len(entries)} part(s) — no improvement", "info")
                    return

                self.log_message.emit(f"🔍  Re-scan found {len(entries)} parts — replacing row", "success")
                # Tag each entry with the rescan source image
                for entry in entries:
                    if not entry.get("source_image"):
                        entry["source_image"] = scan_path
                self._rescan_ready.emit(rd, entries)

            except Exception as e:
                import traceback
                self.log_message.emit(f"⚠  Re-scan failed: {e}", "warning")

        # Run in background thread
        import threading
        t = threading.Thread(target=run_rescan, daemon=True)
        t.start()

        # Store insert position for _add_result_row to use
        self._rescan_target_rd   = rd
        self._rescan_layout_pos  = layout_pos
        self._rescan_insert_idx  = insert_idx

    def _on_rescan_ready(self, rd, entries):
        """Main thread: replace merged row with individual re-scanned parts."""
        try:
            insert_idx = next(i for i,r in enumerate(self._rows) if r is rd)
        except StopIteration:
            return

        layout = self.results_list_layout
        layout_pos = -1
        src_widget = rd.get("_widget")
        for i in range(layout.count()):
            item = layout.itemAt(i)
            if item and item.widget() is src_widget:
                layout_pos = i
                break

        sv = self.results_scroll.verticalScrollBar().value()

        # Remove original row
        if src_widget:
            src_widget.setParent(None)
            src_widget.deleteLater()
        self._rows.pop(insert_idx)
        self._selected.discard(id(rd))
        self._part_count -= 1

        # Build set of existing part_id+color_id combos to avoid duplicates
        existing = {(r.get("part_id",""), r.get("color_id",0))
                    for r in self._rows if not r.get("_deleted") and r.get("part_id")}

        inserted = 0
        skipped  = 0
        for i, entry in enumerate(entries):
            pid  = entry.get("part_id", "")
            cid  = entry.get("color_id", 0)
            key  = (pid, cid)
            if key in existing:
                skipped += 1
                continue
            existing.add(key)
            # Use the crop from the rescan scan_path as source image so row shows it
            if not entry.get("crop_image") and not entry.get("source_image"):
                entry["source_image"] = rd.get("source_image", "")
            entry["index"] = self._part_count + 1
            self._part_count += 1
            self._add_result_row(entry)
            new_rd = self._rows[-1]
            new_widget = new_rd.get("_widget")
            if new_widget and layout_pos >= 0:
                layout.removeWidget(new_widget)
                layout.insertWidget(layout_pos + inserted, new_widget)
            self._rows.pop()
            self._rows.insert(insert_idx + inserted, new_rd)
            inserted += 1

        QTimer.singleShot(0, lambda v=sv: self.results_scroll.verticalScrollBar().setValue(v))
        self._update_total_value()
        self.results_count.setText(f"{self._part_count} parts")
        msg = f"🔍  Re-scan: {inserted} new parts inserted"
        if skipped: msg += f", {skipped} skipped (already in results)"
        self._log(msg, "success")

    def _part_out_minifig(self, rd):
        """Replace a minifig row with its individual parts via BrickLink subsets API."""
        pid = rd.get("part_id", "")
        if not pid: return
        self._log(f"⊞  Parting out {pid}...", "info")

        def fetch_and_expand():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    self.log_message.emit("⚠  Missing credentials for part-out", "warning"); return
                from requests_oauthlib import OAuth1
                import requests as req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"], bl["TOKEN"], bl["TOKEN_SECRET"])
                url = f"https://api.bricklink.com/api/store/v1/items/minifig/{pid}/subsets"
                r = req.get(url, auth=auth, timeout=10)
                if r.status_code != 200:
                    self.log_message.emit(f"⚠  Part-out error {r.status_code} for {pid}", "warning"); return
                matches = r.json().get("data", [])
                parts = []
                for match in matches:
                    for entry in match.get("entries", []):
                        item = entry.get("item", {})
                        if item.get("type") not in ("PART", "MINIFIG"): continue
                        cid = entry.get("color_id", 0)
                        cname, crgb = ("—", (128,128,128))
                        if cid and cid in BL_COLORS:
                            cname = BL_COLORS[cid][0]; crgb = BL_COLORS[cid][1]
                        parts.append({
                            "part_id":    item.get("no", ""),
                            "part_name":  item.get("name", ""),
                            "item_type":  "M" if item.get("type") == "MINIFIG" else "P",
                            "color_id":   cid, "color_name": cname, "color_rgb": crgb,
                            "confidence": 1.0, "qty": entry.get("quantity", 1),
                            "condition":  rd.get("condition", "U"),
                            "color_method": "exact", "known_color_ids": [cid],
                            "source_image": rd.get("source_image", ""),
                            "crop_image": "", "thumb_url": "", "bbox": None,
                            "from_part_out": True,
                        })
                if not parts:
                    self.log_message.emit(f"⚠  No parts found for {pid}", "warning"); return
                self.log_message.emit(f"⊞  {pid} → {len(parts)} parts", "success")
                self._part_out_ready.emit(rd, parts)
            except Exception as e:
                self.log_message.emit(f"⚠  Part-out failed: {e}", "warning")

        threading.Thread(target=fetch_and_expand, daemon=True).start()

    def _on_part_out_ready(self, rd, parts):
        """Main thread: replace minifig row in-place with individual part rows."""
        self._save_undo_snapshot("part out")
        try:
            insert_idx = next(i for i,r in enumerate(self._rows) if r is rd)
        except StopIteration:
            return

        # Find the widget's position in the layout
        layout = self.results_list_layout
        layout_pos = -1
        src_widget = rd.get("_widget")
        for i in range(layout.count()):
            item = layout.itemAt(i)
            if item and item.widget() is src_widget:
                layout_pos = i
                break

        # Save scroll position
        sv = self.results_scroll.verticalScrollBar().value()

        # Remove the minifig row
        if src_widget:
            src_widget.setParent(None)
            src_widget.deleteLater()
        self._rows.pop(insert_idx)
        self._selected.discard(id(rd))

        # Add part rows and insert them at the right layout position
        for i, part in enumerate(parts):
            part["index"] = self._part_count + 1
            self._part_count += 1
            self._add_result_row(part)
            new_rd = self._rows[-1]
            new_widget = new_rd.get("_widget")
            if new_widget and layout_pos >= 0:
                # Remove from end where _add_result_row placed it
                layout.removeWidget(new_widget)
                # Insert at original minifig position + offset
                layout.insertWidget(layout_pos + i, new_widget)
            # Move in self._rows too
            self._rows.pop()
            self._rows.insert(insert_idx + i, new_rd)

        # Restore scroll
        QTimer.singleShot(0, lambda v=sv: self.results_scroll.verticalScrollBar().setValue(v))
        self._update_total_value()
        self.results_count.setText(f"{self._part_count} parts")
        self._log(f"⊞  Parted out — {len(parts)} parts added", "success")

    def _pick_alternative(self, rd):
        """Show a dialog with top Brickognize alternatives and let user pick one.
        Fetches alternate_no live from BrickLink to get full mold variant list."""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QListWidgetItem

        pid = rd.get("part_id", "")
        alts = list(rd.get("alternatives") or [])

        # Fetch alternate_no from BrickLink live — covers mold variants the name can't tell us
        if pid:
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY":["CONSUMER_KEY"],"CONSUMER_SECRET":["CONSUMER_SECRET"],
                      "TOKEN":["TOKEN","ACCESS_TOKEN"],"TOKEN_SECRET":["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) == 4:
                    from requests_oauthlib import OAuth1
                    import requests as _req
                    auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                                  bl["TOKEN"], bl["TOKEN_SECRET"])
                    r2 = _req.get(f"https://api.bricklink.com/api/store/v1/items/part/{pid}",
                                  auth=auth, timeout=6)
                    if r2.status_code == 200:
                        data = r2.json().get("data", {})
                        alt_no = data.get("alternate_no", "")
                        existing_ids = {a.get("part_id","").lower() for a in alts} |                                        {a.get("id","").lower() for a in alts}
                        if alt_no:
                            for alt_pid in alt_no.replace(" ","").split(","):
                                if alt_pid and alt_pid.lower() not in existing_ids:
                                    # Fetch name for this variant
                                    try:
                                        r3 = _req.get(
                                            f"https://api.bricklink.com/api/store/v1/items/part/{alt_pid}",
                                            auth=auth, timeout=4)
                                        name = r3.json().get("data",{}).get("name", alt_pid) if r3.status_code==200 else alt_pid
                                    except Exception:
                                        name = alt_pid
                                    alts.insert(0, {"part_id": alt_pid, "id": alt_pid,
                                                    "name": name, "score": 0.0,
                                                    "is_mold_variant": True})
                                    existing_ids.add(alt_pid.lower())
                        # Also grab BrickLink's own name for current part
                        if not rd.get("part_name") or rd["part_name"] == "Unknown":
                            rd["part_name"] = data.get("name", rd.get("part_name",""))
            except Exception as e:
                self._log(f"⚠  BL catalog fetch: {e}", "warning")

        # Update rd with enriched alts
        rd["alternatives"] = alts

        if not alts:
            self._log("No alternatives available for this row", "info")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Alternatives for {rd.get('part_id','?')}")
        dlg.setFixedSize(420, 320)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
        vl = QVBoxLayout(dlg); vl.setSpacing(6); vl.setContentsMargins(10,10,10,10)

        from PyQt5.QtWidgets import QLabel as _QL2
        vl.addWidget(_QL2(f"Current: <b>{rd.get('part_id','?')}</b> — {rd.get('part_name','?')} "
                          f"({rd.get('confidence',0):.0%})",
                          styleSheet=f"color:{ACCENT2};font-size:11px;"))

        lst = QListWidget()
        lst.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};font-size:12px;")
        from PyQt5.QtGui import QColor as QC
        has_mold = any(a.get("is_mold_variant") for a in alts)
        if has_mold:
            sep = QListWidgetItem("── 🔀 Mold variants (BrickLink catalog) ──")
            sep.setFlags(Qt.NoItemFlags)
            sep.setForeground(QC(200, 160, 40))
            lst.addItem(sep)
            for alt in alts:
                if not alt.get("is_mold_variant"): continue
                item = QListWidgetItem(f"  🔀 {alt.get('id','')}   {alt.get('name','')[:40]}")
                item.setData(Qt.UserRole, alt)
                item.setForeground(QC(200, 160, 40))
                lst.addItem(item)
            sep2 = QListWidgetItem("── 🔍 Brickognize alternatives ──")
            sep2.setFlags(Qt.NoItemFlags)
            sep2.setForeground(QC(120, 120, 120))
            lst.addItem(sep2)

        for alt in alts:
            if alt.get("is_mold_variant"): continue
            if alt.get("_dual_cam_alt"):
                # Dual-cam alternatives get their own section header
                sep_dual = QListWidgetItem("── 📷 Other camera's result ──")
                sep_dual.setFlags(Qt.NoItemFlags)
                sep_dual.setForeground(QC(120, 159, 255))
                lst.addItem(sep_dual)
                score = alt.get("score", 0)
                cname = f"  [{alt.get('color_name','')}]" if alt.get("color_name") else ""
                item = QListWidgetItem(f"  📷 {alt.get('id','')}   {alt.get('name','')[:35]}{cname}   ({score:.0%})")
                item.setData(Qt.UserRole, alt)
                item.setForeground(QC(120, 159, 255))
                lst.addItem(item)
                continue
            score = alt.get("score", 0)
            cc = SUCCESS if score >= 0.7 else WARNING if score >= 0.4 else ACCENT
            item = QListWidgetItem(f"  {alt.get('id','')}   {alt.get('name','')[:40]}   ({score:.0%})")
            item.setData(Qt.UserRole, alt)
            item.setForeground(QC(cc))
            lst.addItem(item)
        vl.addWidget(lst)

        btns = QHBoxLayout()
        back_btn = self._btn("↩ Back", "#1a1a2e", lambda: (self._alt_back(rd), dlg.accept()), w=86)
        back_btn.setToolTip("Go back to previous selected alternative for this row")
        back_btn.setEnabled(bool(rd.get("_alt_back")))
        ok_btn  = self._btn("✓ Use this", SUCCESS, dlg.accept, w=117)
        cxl_btn = self._btn("Cancel", CARD_BG, dlg.reject, w=91)
        ok_btn.setFixedHeight(28); cxl_btn.setFixedHeight(28)
        btns.addWidget(back_btn)
        btns.addStretch(); btns.addWidget(cxl_btn); btns.addWidget(ok_btn)
        vl.addLayout(btns)
        lst.itemDoubleClicked.connect(lambda: dlg.accept())

        if dlg.exec_() != QDialog.Accepted:
            return
        sel = lst.currentItem()
        if not sel: return
        alt = sel.data(Qt.UserRole)

        # Apply the selected alternative
        new_pid   = alt.get("id", "")
        new_name  = alt.get("name", "")
        new_itype = alt.get("type", "P")
        new_conf  = alt.get("score", 0)
        if not new_pid: return

        # Save current state so user can go back/forward
        self._alt_nav_push(rd, action="pick alternative")

        rd["part_id"]    = new_pid
        rd["part_name"]  = new_name
        rd["item_type"]  = self._effective_item_type({"part_id": new_pid, "item_type": new_itype})
        rd["confidence"] = new_conf
        rd["_pid_lbl"].setText(("👤 " if rd["item_type"]=="M" else "") + new_pid)
        rd["_pid_lbl"].setStyleSheet(f"color:{ACCENT2};font-weight:bold;font-size:12px;")
        name_short = new_name[:38]+"…" if len(new_name)>38 else new_name
        rd["_name_lbl"].setText(name_short)
        rd["_name_lbl"].setToolTip(new_name)
        cc = SUCCESS if new_conf>=0.7 else WARNING if new_conf>=0.4 else ACCENT
        # If dual-cam alt, also apply its color
        if alt.get("_dual_cam_alt") and alt.get("color_id") is not None:
            rd["color_id"]   = alt["color_id"]
            rd["color_name"] = alt.get("color_name", rd.get("color_name"))
            if rd.get("color_name") and rd.get("_color_lbl"):
                cn = rd["color_name"][:13]
                rd["_color_lbl"].setText(cn)
            if rd.get("_swatch") and alt.get("color_id") and alt["color_id"] in BL_COLORS:
                rgb = BL_COLORS[alt["color_id"]][1]
                r2,g2,b2 = rgb
                rd["_swatch"].setStyleSheet(
                    f"background:rgb({r2},{g2},{b2});border-radius:3px;"
                    f"border:1px solid rgba(255,255,255,0.15);")
        # Update confidence label
        confl_w = rd.get("_confl")
        if confl_w:
            confl_w.setText(f"{new_conf:.0%}")
            confl_w.setStyleSheet(f"color:{cc};font-weight:bold;font-size:12px;")
        # Reload BL image, price, price guide
        rd["_bl_lbl"].setText("…")
        threading.Thread(target=self._load_bl_img_for,
            args=(new_pid, rd.get("color_id",0), rd["item_type"], rd["_bl_lbl"]),
            daemon=True).start()
        self._fetch_price_for_row(rd)
        self._fetch_price_guide(rd)
        self._log(f"↩  Row updated: {new_pid} — {new_name} ({new_conf:.0%})", "success")

    def _bulk_set_condition(self, cond):
        self._save_undo_snapshot("bulk set condition")
        """Set Used/New condition for all selected rows."""
        rows = self._selected_rows()
        if not rows: return
        for rd in rows:
            rd["condition"] = cond
            if "_cond_btn" in rd and rd["_cond_btn"]:
                rd["_cond_refresh"]() if "_cond_refresh" in rd else None
                # Re-trigger the wire directly
                c = rd.get("condition","U")
                rd["_cond_btn"].setText(c)
                rd["_cond_btn"].setStyleSheet(
                    f"background:{'#1a3a1a' if c=='N' else '#2a1800'};"
                    f"color:{'#5fca7a' if c=='N' else ACCENT2};"
                    f"font-size:10px;font-weight:bold;border-radius:3px;"
                    f"border:1px solid {'#2a5a2a' if c=='N' else '#4a3000'};")
        label = "New" if cond == "N" else "Used"
        self._log(f"  {len(rows)} rows → {label}", "success")

    def _bulk_set_remark(self):
        self._save_undo_snapshot("bulk set remark")
        """Set remark text for all selected rows."""
        rows = self._selected_rows()
        if not rows: return
        val, ok = QInputDialog.getText(self, "Bulk Remark",
            f"Set remark for {len(rows)} rows (leave blank to clear):",
            text="")
        if not ok: return
        val = val.strip()
        for rd in rows:
            rd["remark"] = val
            if "_rmk_lbl" in rd and rd["_rmk_lbl"]:
                rd["_rmk_lbl"].setText(val if val else "remark…")
                rd["_rmk_lbl"].setStyleSheet(
                    f"color:{'#aaa' if not val else TEXT};font-size:10px;"
                    f"font-style:italic;background:transparent;padding:0 4px;")
        self._log(f"✏  {len(rows)} rows → remark: {val!r}", "info")

    def _bulk_aggregate_selected(self):
        self._save_undo_snapshot("bulk merge")
        """Merge selected rows that share the same part_id+color_id, summing qty."""
        sel = self._selected_rows()
        if not sel:
            self._log("⚠  No rows selected for aggregation", "warning")
            return
        seen = {}   # key → first rd
        to_delete = []
        for rd in sel:
            key = (rd.get("part_id"), rd.get("color_id", 0), rd.get("item_type", "P"))
            if key in seen:
                first = seen[key]
                first["qty"] = first.get("qty", 1) + rd.get("qty", 1)
                first["_qty_lbl"].setText(str(first["qty"]))
                to_delete.append(rd)
            else:
                seen[key] = rd
        for rd in to_delete:
            rd["_deleted"] = True
            rd["_widget"].setVisible(False)
            self._selected.discard(id(rd))
        if to_delete:
            self._log(f"Σ  Aggregated {len(to_delete)} duplicate(s) in selection", "success")
        else:
            self._log("Σ  No duplicates in selection", "info")
        self._update_bulk_bar()

    def _bulk_set_price(self):
        self._save_undo_snapshot("bulk set price")
        """Ask for a fixed price and apply it to all selected rows."""
        rows = self._selected_rows()
        if not rows: return
        from PyQt5.QtWidgets import QInputDialog
        val, ok = QInputDialog.getDouble(self, "Bulk Set Price",
            f"Set price for {len(rows)} selected rows ($):",
            1.00, 0.01, 9999.99, 2)
        if not ok: return
        for rd in rows:
            rd["price"] = val
            rd["_price_lbl"].setText(f"${val:.2f}")
            rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;")
        self._log(f"✏  {len(rows)} rows → ${val:.2f}", "success")
        self._update_total_value()

    def _bulk_override_color(self):
        self._save_undo_snapshot("bulk set color")
        """Open a color picker and apply the chosen color to all selected rows."""
        sel_rows = [r for r in self._selected_rows() if r.get("item_type") != "M"]
        if not sel_rows:
            self._log("⚠  No eligible rows (parts only — minifigs have no color)", "warning")
            return

        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QListWidget, QListWidgetItem
        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location("scan_heads", "scan-heads.py")
            sh   = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(sh)
            BL_COLORS = sh.BRICKLINK_COLORS
        except Exception:
            BL_COLORS = {}

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Set Color — {len(sel_rows)} rows")
        dlg.setFixedSize(320, 440)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")
        vl = QVBoxLayout(dlg); vl.setSpacing(6)
        vl.addWidget(QLabel(f"Choose color for <b>{len(sel_rows)}</b> selected rows:"))
        search = QLineEdit(); search.setPlaceholderText("Filter colors…")
        search.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};padding:4px;")
        vl.addWidget(search)
        lst = QListWidget()
        lst.setStyleSheet(f"background:{CARD_BG};color:{TEXT};border:1px solid {BORDER};font-size:12px;")
        all_known_bulk = set()
        for rd_ in sel_rows:
            all_known_bulk.update(rd_.get("known_color_ids", []))
        known_bulk = [(cid, BL_COLORS[cid][0], BL_COLORS[cid][1])
                      for cid in sorted(all_known_bulk) if cid in BL_COLORS]
        other_bulk = [(cid, name, rgb) for cid,(name,rgb) in BL_COLORS_SORTED
                      if cid not in all_known_bulk]

        def _populate(filt=""):
            lst.clear()
            from PyQt5.QtGui import QColor as QC
            if known_bulk:
                sep = QListWidgetItem(f"── ⭐ Known colors ({len(known_bulk)}) ──")
                sep.setFlags(Qt.NoItemFlags); sep.setForeground(QC(200,160,40))
                lst.addItem(sep)
            for cid, cname, rgb in known_bulk:
                if filt and filt.lower() not in cname.lower(): continue
                item = QListWidgetItem(f"  {cname}  (id {cid})")
                item.setData(Qt.UserRole, (cid, cname, rgb))
                item.setBackground(QC(*rgb))
                lum = 0.299*rgb[0]+0.587*rgb[1]+0.114*rgb[2]
                item.setForeground(QC(0,0,0) if lum > 128 else QC(255,255,255))
                lst.addItem(item)
            sep2 = QListWidgetItem("── All other colors ──")
            sep2.setFlags(Qt.NoItemFlags); sep2.setForeground(QC(120,120,120))
            lst.addItem(sep2)
            for cid, cname, rgb in other_bulk:
                if filt and filt.lower() not in cname.lower(): continue
                item = QListWidgetItem(f"  {cname}  (id {cid})")
                item.setData(Qt.UserRole, (cid, cname, rgb))
                item.setBackground(QC(*rgb))
                lum = 0.299*rgb[0]+0.587*rgb[1]+0.114*rgb[2]
                item.setForeground(QC(0,0,0) if lum > 128 else QC(255,255,255))
                lst.addItem(item)
        _populate()
        search.textChanged.connect(_populate)
        vl.addWidget(lst)
        btns = QHBoxLayout()
        ok_btn  = self._btn("✓ Apply to all selected", SUCCESS, dlg.accept, w=208)
        cxl_btn = self._btn("Cancel", CARD_BG, dlg.reject, w=91)
        ok_btn.setFixedHeight(28); cxl_btn.setFixedHeight(28)
        btns.addStretch(); btns.addWidget(cxl_btn); btns.addWidget(ok_btn)
        vl.addLayout(btns)
        lst.itemDoubleClicked.connect(lambda: dlg.accept())
        if dlg.exec_() != QDialog.Accepted: return
        sel = lst.currentItem()
        if not sel: return
        cid, cname, rgb = sel.data(Qt.UserRole)

        for rd in sel_rows:
            rd["color_id"]   = cid
            rd["color_name"] = cname
            rd["color_rgb"]  = rgb
            rd["_swatch"].setStyleSheet(f"background:rgb{tuple(rgb)};border-radius:3px;border:1px solid #444;")
            rd["_color_lbl"].setText(cname[:13])
            rd["_color_lbl"].setStyleSheet(f"color:{TEXT};font-size:11px;")
            rd["_bl_lbl"].setText("…")
            threading.Thread(target=self._load_bl_img_for,
                args=(rd.get("part_id",""), cid, rd.get("item_type","P"), rd["_bl_lbl"]),
                daemon=True).start()
        # Reset stale price and re-fetch for new color
        for rd in sel_rows:
            rd["price"] = None
            rd["medium_price"] = None
            rd["_price_lbl"].setText("—")
            rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
            self._fetch_price_for_row(rd)
        self._log(f"🎨  {len(sel_rows)} rows → {cname} (id {cid})", "success")

    def _bulk_apply_scan_color(self):
        """Apply (or revert) Brickognize scan color for all selected rows.
        First call → restores _scan_color_*  (BG Color mode).
        Second call on same rows → reverts to whatever color was set before (toggle).
        """
        sel_rows = self._selected_rows()
        if not sel_rows:
            self._log("⚠  No rows selected", "warning"); return
        self._save_undo_snapshot("bulk BG color")

        applied = reverted = 0
        for rd in sel_rows:
            scan_cid  = rd.get("_scan_color_id")
            scan_name = rd.get("_scan_color_name")
            scan_rgb  = rd.get("_scan_color_rgb")
            if scan_cid is None:
                continue  # no scan color stored (manually-added row)

            cur_cid = rd.get("color_id")

            if cur_cid != scan_cid:
                # Save current color so the user can revert back
                rd["_pre_scan_color_id"]   = cur_cid
                rd["_pre_scan_color_name"] = rd.get("color_name")
                rd["_pre_scan_color_rgb"]  = rd.get("color_rgb")
                # Apply scan color
                rd["color_id"]   = scan_cid
                rd["color_name"] = scan_name
                rd["color_rgb"]  = scan_rgb or (128, 128, 128)
                applied += 1
            else:
                # Already at scan color — revert to pre-scan color if available
                prev_cid  = rd.get("_pre_scan_color_id")
                prev_name = rd.get("_pre_scan_color_name")
                prev_rgb  = rd.get("_pre_scan_color_rgb")
                if prev_cid is not None:
                    rd["color_id"]   = prev_cid
                    rd["color_name"] = prev_name
                    rd["color_rgb"]  = prev_rgb or (128, 128, 128)
                    reverted += 1

            # Refresh swatch + label
            rgb = rd.get("color_rgb") or (128, 128, 128)
            r2, g2, b2 = rgb
            rd["_swatch"].setStyleSheet(
                f"background:rgb({r2},{g2},{b2});border-radius:3px;"
                f"border:1px solid rgba(255,255,255,0.15);")
            cn = (rd.get("color_name") or "—")
            cn = cn[:13] if len(cn) > 13 else cn
            rd["_color_lbl"].setText(cn)
            # Re-fetch price for new color
            rd["_price_lbl"].setText("…")
            rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
            self._fetch_price_for_row(rd)

        msg = []
        if applied:  msg.append(f"{applied} → scan color")
        if reverted: msg.append(f"{reverted} → reverted")
        self._log(f"🔬  BG Color: {', '.join(msg) or 'no change'}", "success")

    def _bulk_pixel_recolor(self):
        """Re-sample the crop image pixel color for selected rows using Lab ΔE against the
        full BrickLink color table — completely ignores what Brickognize said.
        Respects known_color_ids if available (narrows candidates), otherwise full table.
        """
        sel_rows = self._selected_rows()
        if not sel_rows:
            self._log("⚠  No rows selected", "warning"); return

        # Check at least one has a crop image
        if not any(rd.get("crop_image") for rd in sel_rows):
            self._log("⚠  No crop images available for selected rows", "warning"); return

        self._save_undo_snapshot("bulk recolor")

        import numpy as np
        from PIL import Image as _PIL_Image

        def _rgb_to_lab(rgb):
            """sRGB → CIE Lab (D65)."""
            r, g, b = [x / 255.0 for x in rgb]
            def linearize(c):
                return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
            r, g, b = linearize(r), linearize(g), linearize(b)
            X = r*0.4124564 + g*0.3575761 + b*0.1804375
            Y = r*0.2126729 + g*0.7151522 + b*0.0721750
            Z = r*0.0193339 + g*0.1191920 + b*0.9503041
            X /= 0.95047; Z /= 1.08883
            def f(t): return t**0.3333 if t > 0.008856 else 7.787*t + 16/116
            L = 116*f(Y) - 16; a = 500*(f(X)-f(Y)); b_ = 200*(f(Y)-f(Z))
            return (L, a, b_)

        def _sample_rgb(crop_path):
            """Sample dominant non-background color from crop image."""
            try:
                arr = np.array(_PIL_Image.open(crop_path).convert("RGB").resize((100,100)))
                h, w = arr.shape[:2]
                def clean(region):
                    px = region.reshape(-1,3).astype(float)
                    mx = px.max(axis=1); mn = px.min(axis=1)
                    sat = np.where(mx>0, (mx-mn)/mx, 0)
                    not_white    = ~((px[:,0]>210)&(px[:,1]>210)&(px[:,2]>210))
                    not_black    = ~(mx < 25)
                    not_specular = ~((sat<0.10)&(mx>190))
                    not_grey     = ~(sat<0.07)
                    return px[not_white & not_black & not_specular & not_grey]
                best = np.empty((0,3))
                for frac in [0.50, 0.65, 0.80, 1.00]:
                    y0=int(h*(1-frac)/2); y1=int(h*(1+frac)/2)
                    x0=int(w*(1-frac)/2); x1=int(w*(1+frac)/2)
                    px = clean(arr[y0:y1, x0:x1])
                    if len(px) >= 20: best = px; break
                    if len(px) > len(best): best = px
                if len(best) < 4:
                    px = arr.reshape(-1,3).astype(float)
                    mask = ~((px[:,0]>215)&(px[:,1]>215)&(px[:,2]>215))
                    best = px[mask] if mask.sum()>=4 else px
                BINS=32; scale=256.0/BINS
                bins = np.floor(best/scale).astype(int).clip(0,BINS-1)
                idx = bins[:,0]*BINS*BINS + bins[:,1]*BINS + bins[:,2]
                counts = np.bincount(idx, minlength=BINS**3)
                b = int(counts.argmax())
                return (int((b//(BINS*BINS)+0.5)*scale),
                        int(((b//BINS)%BINS+0.5)*scale),
                        int((b%BINS+0.5)*scale))
            except Exception as e:
                return None

        # Build Lab cache once for all BL colors
        lab_cache = {cid: _rgb_to_lab(rgb) for cid,(name,rgb) in BL_COLORS.items()}

        changed = skipped = 0
        for rd in sel_rows:
            crop_path = rd.get("crop_image", "")
            if not crop_path or not Path(crop_path).exists():
                skipped += 1; continue

            sampled = _sample_rgb(crop_path)
            if sampled is None:
                skipped += 1; continue

            s_lab = _rgb_to_lab(sampled)

            # Luminance guard: if image is dark, only consider dark BL colors
            try:
                arr_chk = np.array(_PIL_Image.open(crop_path).convert("RGB").resize((60,60)))
                med_br  = float(np.median(arr_chk.max(axis=2)))
            except Exception:
                med_br = 128.0
            dark_part = med_br < 60

            # Candidate pool: known_color_ids if available, else full table
            known = rd.get("known_color_ids") or []
            if known:
                candidates = [(cid, BL_COLORS[cid][0], BL_COLORS[cid][1])
                              for cid in known if cid in BL_COLORS]
            else:
                candidates = [(cid, name, rgb) for cid,(name,rgb) in BL_COLORS.items()]

            # Apply luminance filter when part is clearly dark
            if dark_part:
                dark_cands = [(cid,n,rgb) for cid,n,rgb in candidates
                              if lab_cache.get(cid,(100,))[0] < 40]
                if dark_cands:
                    candidates = dark_cands

            best_cid = best_name = best_rgb = None
            best_dist = float("inf")
            for cid, name, rgb in candidates:
                ref = lab_cache.get(cid)
                if ref is None: continue
                d = ((s_lab[0]-ref[0])**2+(s_lab[1]-ref[1])**2+(s_lab[2]-ref[2])**2)**0.5
                if d < best_dist:
                    best_dist = d; best_cid = cid; best_name = name; best_rgb = rgb

            if best_cid is None:
                skipped += 1; continue

            # Apply
            rd["color_id"]   = best_cid
            rd["color_name"] = best_name
            rd["color_rgb"]  = best_rgb

            # Refresh swatch + label
            r2,g2,b2 = best_rgb
            rd["_swatch"].setStyleSheet(
                f"background:rgb({r2},{g2},{b2});border-radius:3px;"
                f"border:1px solid rgba(255,255,255,0.15);")
            cn = best_name[:13] if len(best_name)>13 else best_name
            rd["_color_lbl"].setText(cn)
            rd["_price_lbl"].setText("…")
            rd["_price_lbl"].setStyleSheet(f"color:{TEXT_DIM};font-size:11px;")
            self._fetch_price_for_row(rd)
            changed += 1

        self._log(f"🎯  Recolor: {changed} updated, {skipped} skipped (no crop)", "success")

    # ── Direct BrickLink Upload ───────────────────────────────────────────────
    def _upload_to_bricklink(self):
        """Copy XML to clipboard and open the BrickLink Upload Inventory page."""
        from PyQt5.QtWidgets import QApplication
        xml = self._build_xml_from_rows()   # uses per-row condition
        if not xml:
            self._log("⚠  No rows to upload", "warning"); return
        QApplication.clipboard().setText(xml)
        n = sum(1 for r in self._rows if not r.get("_deleted") and r.get("part_id"))
        self._log(f"📋  {n} lots copied to clipboard — opening BrickLink upload page...", "success")
        self._log("    Paste the XML into the box and click Upload.", "info")
        import webbrowser
        webbrowser.open("https://www.bricklink.com/invXML.asp")


    def _clear_price_guide(self):
        self.pg_part_label.setText("—")
        self._pg_current_rd = None
        self._pg_last_avg = None
        self.pg_apply_btn.setEnabled(False)
        self.pg_apply_med_btn.setEnabled(False)
        for w in [self.pg_min_u, self.pg_avg_u, self.pg_qty_avg_u,
                  self.pg_max_u, self.pg_lots_u, self.pg_units_u,
                  self.pg_min_n, self.pg_avg_n, self.pg_qty_avg_n,
                  self.pg_max_n, self.pg_lots_n, self.pg_units_n]:
            w.setText("—")
            w.setStyleSheet(f"font-size:10px;color:{TEXT};font-weight:bold;")

    def _pg_apply_avg(self):
        """Apply price guide avg to currently selected rows."""
        avg = getattr(self, "_pg_last_avg", None)
        if avg is None: return
        targets = [self._rows[i] for i in self._selected if not self._rows[i].get("_deleted")]
        if not targets and self._pg_current_rd:
            targets = [self._pg_current_rd]
        for rd in targets:
            rd["price"] = round(avg, 2)
            rd["_price_lbl"].setText(f"{avg:.2f}")
            rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;font-weight:bold;")

    def _pg_apply_to_all(self):
        """Apply price guide avg to every row with the same part+color."""
        avg = getattr(self, "_pg_last_avg", None)
        rd0 = self._pg_current_rd
        if avg is None or rd0 is None: return
        pid   = rd0.get("part_id")
        color = rd0.get("color_id")
        count = 0
        for rd in self._rows:
            if rd.get("_deleted"): continue
            if rd.get("part_id") == pid and rd.get("color_id") == color:
                rd["price"] = round(avg, 2)
                rd["_price_lbl"].setText(f"{avg:.2f}")
                rd["_price_lbl"].setStyleSheet(f"color:{SUCCESS};font-size:11px;font-weight:bold;")
                count += 1
        self._log(f"💲  Applied {avg:.2f} to {count} row(s) of {pid}", "success")

    def _preview_source_image(self, source_path, bbox):
        """Show source image in preview panel with bounding box highlighted."""
        if not source_path or not Path(source_path).exists():
            return
        try:
            from PIL import Image as PILImage, ImageDraw
            img = PILImage.open(source_path).convert("RGB")
            if bbox:
                x1, y1, x2, y2 = bbox
                draw = ImageDraw.Draw(img)
                # Bright highlight box with thick border
                pad = 6
                for t in range(4):
                    draw.rectangle([x1-pad-t, y1-pad-t, x2+pad+t, y2+pad+t],
                                   outline=(255, 80, 0))
                # Semi-transparent fill hint — just a bright outline is enough
            # Scale to fit preview
            import io
            buf = io.BytesIO()
            # Resize to reasonable preview size preserving aspect
            w, h = img.size
            max_dim = 1200
            if max(w, h) > max_dim:
                scale = max_dim / max(w, h)
                img = img.resize((int(w*scale), int(h*scale)), PILImage.LANCZOS)
            img.save(buf, "JPEG", quality=90)
            buf.seek(0)
            from PyQt5.QtGui import QPixmap
            from PyQt5.QtCore import QByteArray
            pm = QPixmap()
            pm.loadFromData(QByteArray(buf.read()))
            if not pm.isNull():
                lw = max(self.preview_label.width(), 300)
                lh = max(self.preview_label.height(), 220)
                self.preview_label.setPixmap(pm.scaled(lw, lh, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                self.preview_label.setText("")
                self._live_paused = True
                self._update_preview_toggle()
                # Store for enlarge
                self._last_preview_source = source_path
                self._last_preview_bbox = bbox
        except Exception as e:
            self._log(f"⚠  Preview error: {e}", "warning")

    def _show_enlarged(self, source_path, bbox, crop_path=None):
        """Open a resizable popup showing the full-res source image or crop."""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QScrollArea, QPushButton
        from PyQt5.QtGui import QPixmap
        from PyQt5.QtCore import QByteArray
        import io

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Full resolution — {Path(source_path).name if source_path else 'crop'}")
        dlg.resize(1000, 750)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT};")

        vl = QVBoxLayout(dlg)
        vl.setContentsMargins(8, 8, 8, 8)

        # Toggle: source image vs crop
        btn_row = QHBoxLayout()
        self._enlarge_show_source = True

        img_label = QLabel()
        img_label.setAlignment(Qt.AlignCenter)
        img_label.setStyleSheet(f"background:{CARD_BG};")

        _current_pm = [None]   # mutable container so resizeEvent can access it

        def _show(show_source):
            try:
                from PIL import Image as PILImage, ImageDraw
                if show_source and source_path and Path(source_path).exists():
                    img = PILImage.open(source_path).convert("RGB")
                    if bbox:
                        x1, y1, x2, y2 = bbox
                        draw = ImageDraw.Draw(img)
                        pad = 6
                        for t in range(5):
                            draw.rectangle([x1-pad-t, y1-pad-t, x2+pad+t, y2+pad+t],
                                           outline=(255, 80, 0))
                elif crop_path and Path(crop_path).exists():
                    img = PILImage.open(crop_path).convert("RGB")
                else:
                    img_label.setText("No image available")
                    return
                buf = io.BytesIO()
                img.save(buf, "JPEG", quality=95)
                buf.seek(0)
                pm = QPixmap()
                pm.loadFromData(QByteArray(buf.read()))
                if not pm.isNull():
                    _current_pm[0] = pm
                    _fit_image()
                    img_label.setText("")
            except Exception as e:
                img_label.setText(f"Error: {e}")

        def _fit_image():
            pm = _current_pm[0]
            if pm is None: return
            w = max(scroll.width()  - 4, 100)
            h = max(scroll.height() - 4, 100)
            img_label.setPixmap(pm.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        def _on_resize(event):
            _fit_image()
            QDialog.resizeEvent(dlg, event)
        dlg.resizeEvent = _on_resize

        src_btn = self._btn("📷 Source image", CARD_BG, lambda: (_show(True), src_btn.__setattr__("_active", True)))
        src_btn.setFixedHeight(26)
        crop_btn_w = self._btn("🔍 Crop only", CARD_BG, lambda: _show(False))
        crop_btn_w.setFixedHeight(26)
        if not crop_path:
            crop_btn_w.setEnabled(False)
        btn_row.addWidget(src_btn)
        btn_row.addWidget(crop_btn_w)
        btn_row.addStretch()
        close_btn = self._btn("✕ Close", "#5a2020", dlg.accept)
        close_btn.setFixedHeight(26)
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

        scroll = QScrollArea()
        scroll.setWidget(img_label)
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background:{CARD_BG};border:none;")
        vl.addWidget(scroll)

        _show(True)   # load image; will fit after dialog is shown
        dlg.show()    # show first so scroll has real dimensions
        QTimer.singleShot(50, _fit_image)  # re-fit once laid out
        dlg.exec_()

    def _set_preview(self, path):
        self._last_preview_path = path   # remember for grid redraw
        try:
            px = QPixmap(path)
            if not px.isNull():
                lw = max(self.preview_label.width(),  300)
                lh = max(self.preview_label.height(), 220)
                pix = px.scaled(lw, lh, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                # Draw grid overlay on Last image too (same as live feed)
                if getattr(self, "grid_check", None) and self.grid_check.isChecked():
                    cols = self.cols_spin.value()
                    rows = self.rows_spin.value()
                    from PyQt5.QtGui import QPainter, QPen, QFont as QF
                    painter = QPainter(pix)
                    painter.setRenderHint(QPainter.Antialiasing, False)
                    pen = QPen(QColor(255, 200, 0, 200)); pen.setWidth(1)
                    painter.setPen(pen)
                    pw, ph = pix.width(), pix.height()
                    cell_w = pw / cols; cell_h = ph / rows
                    for c in range(1, cols):
                        painter.drawLine(int(c * cell_w), 0, int(c * cell_w), ph)
                    for r in range(1, rows):
                        painter.drawLine(0, int(r * cell_h), pw, int(r * cell_h))
                    pen2 = QPen(QColor(255, 200, 0, 160)); pen2.setWidth(2)
                    painter.setPen(pen2)
                    painter.drawRect(0, 0, pw-1, ph-1)
                    painter.setPen(QPen(QColor(255, 220, 0, 220)))
                    font = QF(); font.setPointSize(7); font.setBold(True)
                    painter.setFont(font)
                    for r in range(rows):
                        for c in range(cols):
                            painter.drawText(int(c*cell_w)+3, int(r*cell_h)+10, str(r*cols+c+1))
                    painter.end()
                self.preview_label.setPixmap(pix)
                self.preview_label.setText("")
                self._live_paused = True  # stay on Last until user clicks Live
                self._update_preview_toggle()
        except Exception as e:
            self.preview_label.setText(f"Preview error: {e}")

    def _resume_live_preview(self):
        self._live_paused = False

    def _update_preview_toggle(self):
        """Sync button appearance to current _live_paused state."""
        _btn_base = "font-size:11px;border-radius:4px;border:1px solid #333;padding:0 6px;"
        if self._live_paused:
            self.live_btn.setStyleSheet(f"background:{CARD_BG};color:{TEXT};{_btn_base}")
            self.snap_btn.setStyleSheet(f"background:{SUCCESS};color:#000;{_btn_base}")
        else:
            self.live_btn.setStyleSheet(f"background:{SUCCESS};color:#000;{_btn_base}")
            self.snap_btn.setStyleSheet(f"background:{CARD_BG};color:{TEXT};{_btn_base}")

    def _fetch_bl_store_total(self):
        """Fetch BrickLink store total lots/qty/value from API and update the store label."""
        if getattr(self, "_bl_store_fetching", False):
            return
        self._bl_store_fetching = True
        if hasattr(self, "bl_store_label"):
            self.bl_store_label.setText("  🏪 …  ")

        def _fetch():
            try:
                env = {}
                if Path(".env").exists():
                    for line in Path(".env").read_text().splitlines():
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            env[k.strip()] = v.strip().strip('"').strip("'")
                km = {"CONSUMER_KEY": ["CONSUMER_KEY"], "CONSUMER_SECRET": ["CONSUMER_SECRET"],
                      "TOKEN": ["TOKEN", "ACCESS_TOKEN"], "TOKEN_SECRET": ["TOKEN_SECRET"]}
                bl = {}
                for canon, aliases in km.items():
                    for a in aliases:
                        if a in env: bl[canon] = env[a]; break
                if len(bl) < 4:
                    if hasattr(self, "bl_store_label"):
                        self.bl_store_label.setText("  🏪 No API keys  ")
                    return
                from requests_oauthlib import OAuth1
                import requests as _req
                auth = OAuth1(bl["CONSUMER_KEY"], bl["CONSUMER_SECRET"],
                              bl["TOKEN"], bl["TOKEN_SECRET"])
                currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"
                resp = _req.get("https://api.bricklink.com/api/store/v1/inventories",
                                auth=auth, timeout=15)
                if resp.status_code != 200:
                    if hasattr(self, "bl_store_label"):
                        self.bl_store_label.setText(f"  🏪 Error {resp.status_code}  ")
                    return
                items = resp.json().get("data", [])
                lots  = len(items)
                qty   = sum(i.get("quantity", 0) for i in items)
                val   = sum(float(i.get("unit_price", 0) or 0) * int(i.get("quantity", 0) or 0)
                            for i in items)
                txt = f"  🏪 {lots} lots · {qty} pcs · ${val:.2f} {currency}  "
                if hasattr(self, "bl_store_label"):
                    self.bl_store_label.setText(txt)
                    self.bl_store_label.setStyleSheet(
                        f"font-size:11px;color:{SUCCESS};font-weight:bold;padding:0 8px;"
                        f"background:#1a1a1a;border-radius:4px;border:1px solid {BORDER};"
                        f"min-width:160px;")
            except Exception as e:
                import traceback
                err_msg = str(e)
                self.log_message.emit(f"⚠  BL store total error: {err_msg}", "warning")
                self.log_message.emit(traceback.format_exc(), "warning")
                if hasattr(self, "bl_store_label"):
                    self.bl_store_label.setText(f"  🏪 {err_msg[:30]}  ")
            finally:
                self._bl_store_fetching = False

        threading.Thread(target=_fetch, daemon=True).start()

    def _update_total_value(self):
        """Recompute and display total inventory value — always visible bottom-right."""
        currency = self.currency_combo.currentText() if hasattr(self, "currency_combo") else "CAD"
        all_rows = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        sel      = self._selected_rows()

        def _qty(row, default=1):
            """Robust qty getter (rows sometimes carry qty=None during edits/undo)."""
            try:
                v = row.get("qty", default)
                if v is None or v == "":
                    return int(default)
                return int(v)
            except Exception:
                return int(default)

        all_lots  = len(all_rows)
        all_qty   = sum(_qty(r, 1) for r in all_rows)
        all_total = sum((r.get("price") or 0) * _qty(r, 1) for r in all_rows)
        priced    = sum(1 for r in all_rows if r.get("price"))

        _base_style = (f"font-size:11px;font-weight:bold;padding:0 8px;"
                       f"background:#1a1a1a;border-radius:4px;border:1px solid {BORDER};"
                       f"min-width:320px;")

        if not all_rows:
            self.total_value_label.setText("  No inventory  ")
            self.total_value_label.setStyleSheet(_base_style + f"color:{TEXT_DIM};")
            return

        unpriced     = all_lots - priced
        unpriced_str = f"  ⚠ {unpriced} unpriced" if unpriced else ""

        if sel:
            sel_lots  = len(sel)
            sel_qty   = sum(_qty(r, 1) for r in sel)
            sel_total = sum((r.get("price") or 0) * _qty(r, 1) for r in sel)
            self.total_value_label.setText(
                f"  ☑ {sel_lots} lots · {sel_qty} pcs · ${sel_total:.2f}  │  "
                f"Total: {all_lots} lots · {all_qty} pcs · ${all_total:.2f} {currency}"
                f"{unpriced_str}  ")
            self.total_value_label.setStyleSheet(_base_style + f"color:{ACCENT2};")
        else:
            self.total_value_label.setText(
                f"  Inventory: {all_lots} lots · {all_qty} pcs · "
                f"${all_total:.2f} {currency}{unpriced_str}  ")
            color = ACCENT2 if not unpriced else WARNING
            self.total_value_label.setStyleSheet(_base_style + f"color:{color};")

    def _redraw_preview(self):
        """Redraw current Last preview — called when grid settings change."""
        p = getattr(self, "_last_preview_path", None)
        if p and self._live_paused:
            self._set_preview(p)

    def _show_live(self):
        self._live_paused = False
        self._update_preview_toggle()
        # Restart camera/stream if not already running
        if not self._live_running and not self._stream_active:
            if self._use_http_stream and self._http_stream_url:
                self._toggle_stream()
            else:
                self._start_live_camera()

    def _show_last_snap(self):
        self._live_paused = True
        self._update_preview_toggle()
        if hasattr(self, '_last_snap_path') and self._last_snap_path:
            self._set_preview(self._last_snap_path)

    def _save_snap_path(self, path):
        self._last_snap_path = path

    # ── Settings ──────────────────────────────────────────────────────────────
    def _get_settings(self):
        cam2_idx = -1
        if hasattr(self, "dual_cam_chk") and self.dual_cam_chk.isChecked():
            cam2_idx = self.camera2_combo.currentData() if hasattr(self, "camera2_combo") else -1
        return {
            "mode":       "all",
            "confidence": self.conf_slider.value()/100,
            "currency":   self.currency_combo.currentText(),
            "color":      self.color_combo.currentData() or "",
            "qty":        self.qty_spin.value(),
            "gap":        self.gap_spin.value(),
            "padding":    self.padding_spin.value(),
            "studs":      getattr(self, "_det_mode", "standard") == "studs",
            "geometric":  getattr(self, "_det_mode", "standard") == "geometric",
            "brightness_bias": self.brightness_bias.value(),
            "grid":       self.grid_check.isChecked(),
            "cols":       self.cols_spin.value(),
            "rows":       self.rows_spin.value(),
            "bg_color":    getattr(self, "_bg_rgb", None),
            "shadow_color": getattr(self, "_shadow_rgb", None),
            "dual_cam":   hasattr(self, "dual_cam_chk") and self.dual_cam_chk.isChecked() and cam2_idx >= 0,
            "cam2_idx":   cam2_idx,
            # NEW (2026): optional “HD scan” override (higher detection max-side)
            "max_side":   getattr(self, "_max_side_override", 0) or 0,
        }

    # ── Scan ──────────────────────────────────────────────────────────────────
    def _cancel_scan(self):
        """Kill the running scan-heads subprocess and clear any batch queue."""
        proc = getattr(self, "_scan_proc", None)
        if proc and proc.poll() is None:
            proc.kill()
            self._log("✕  Scan cancelled", "warning")
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
        # Clear batch queue so cancel stops the whole batch
        if getattr(self, "_batch_queue", []):
            self._log(f"✕  Batch cancelled ({len(self._batch_queue)} remaining images skipped)", "warning")
        self._batch_queue = []
        self._batch_total = 0
        self.batch_btn.setText("📂 Batch")
        self.folder_btn.setText("📁 Folder")
        self._reset_scan_btn()
        self.scan_btn.setEnabled(True)
        self.scan_btn.setText("▶  SCAN")
        if hasattr(self, "cancel_btn"): self.cancel_btn.setEnabled(False)
        self.scan_btn.setStyleSheet(self.scan_btn.styleSheet().replace("#c0392b","#1a5a1a").replace("■  CANCEL","▶  SCAN"))
        self.scan_btn.setText("▶  SCAN")
        self.scan_btn.setStyleSheet("""
            QPushButton { background: #1a3a1a; color: white; font-size: 22px;
            font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background: #2a5a2a; }
        """)
        self.load_img_btn.setEnabled(True)

    def _load_image_file(self):
        self._reset_scan_btn()
        self.batch_btn.setText("📂 Batch")
        self.folder_btn.setText("📁 Folder")
        """Open a file picker, show the image in preview, and arm it for scanning.
        Does NOT start the scan — adjust grid/settings first, then hit SCAN."""
        from PyQt5.QtWidgets import QFileDialog
        start = str(Path(self._iphone_photo_dir)) if getattr(self, "_iphone_photo_dir", "") else ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Select image to scan", start,
            "Images (*.jpg *.jpeg *.png *.bmp *.tiff *.webp)")
        if not path:
            return
        self._forced_image_path = path
        self._log(f"📂  Image loaded: {Path(path).name} — adjust settings then hit SCAN", "success")
        self._set_preview(path)
        # Show filename on the button so user knows it's armed
        name = Path(path).name
        short = name[:22] + "…" if len(name) > 24 else name
        self.load_img_btn.setText(f"📂  {short}")
        self.load_img_btn.setToolTip("Loaded: " + path + "\nClick to change image")

    def _start_capture_countdown(self):
        """Start 3-second countdown then grab current frame — no scan, no API."""
        if not (self._last_frame is not None or self._stream_active or self._live_running):
            self._log("⚠  No live feed active — connect camera or stream first", "warning")
            return
        self._capture_countdown = 3
        self.capture_btn.setEnabled(False)
        self._tick_capture_countdown()

    def _tick_capture_countdown(self):
        n = self._capture_countdown
        if n > 0:
            self.capture_btn.setText(f"📸  {n}...")
            self._capture_countdown -= 1
            QTimer.singleShot(1000, self._tick_capture_countdown)
        else:
            self._do_capture()

    def _do_capture(self):
        """Grab the current frame and save it — no Brickognize, no BL API."""
        from datetime import datetime
        frame = self._last_frame
        if frame is None:
            self._log("⚠  No frame available to capture", "warning")
            self.capture_btn.setText("📸 Capture")
            self.capture_btn.setEnabled(True)
            return
        captures_dir = Path("scans") / "captures"
        captures_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = captures_dir / f"capture-{ts}.jpg"
        try:
            import cv2
            # Apply crop if configured
            cfg_path = Path("station.cfg")
            if cfg_path.exists():
                import json as _json
                cfg = _json.loads(cfg_path.read_text())
                cam_idx = cfg.get("active_camera", cfg.get("camera", 0))
                cam_profiles = cfg.get("cameras", {})
                cam_profile = cam_profiles.get(str(cam_idx), {})
                crop = cam_profile.get("crop") or cfg.get("crop")
                if crop:
                    x1, y1, x2, y2 = crop
                    frame = frame[y1:y2, x1:x2]
            cv2.imwrite(str(out_path), frame, [cv2.IMWRITE_JPEG_QUALITY, 97])
            h, w = frame.shape[:2]
            self._log(f"📸  Captured: {out_path.name}  ({w}×{h}px)", "success")
            self._set_preview(str(out_path))
            # Show the captures folder path once
            if not getattr(self, "_captures_folder_logged", False):
                self._captures_folder_logged = True
                self._log(f"   📁 Saves to: {captures_dir.resolve()}", "info")
                self._log(f"   → Use 📁 Folder to batch-scan when done shooting", "info")
        except Exception as e:
            self._log(f"⚠  Capture failed: {e}", "error")
        finally:
            self.capture_btn.setText("📸 Capture")
            self.capture_btn.setEnabled(True)

    def _load_batch_files(self):
        """Pick multiple image files — queued and scanned sequentially."""
        from PyQt5.QtWidgets import QFileDialog
        start = str(Path(self._iphone_photo_dir)) if getattr(self, "_iphone_photo_dir", "") else ""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select images to batch scan", start,
            "Images (*.jpg *.jpeg *.png *.bmp *.tiff *.webp)")
        if not paths:
            return
        paths = sorted(paths)
        self._batch_queue = paths
        self._batch_total = len(paths)
        self._log(f"📂  Batch loaded: {len(paths)} images — hit SCAN to start", "success")
        self.batch_btn.setText(f"📂 {len(paths)} files")
        self.batch_btn.setToolTip("\n".join(Path(p).name for p in paths))
        self._set_batch_scan_mode(f"BATCH  {len(paths)} images")
        self._stop_live_camera()
        self._live_paused = True
        self._update_preview_toggle()
        self._set_preview(paths[0])

    def _load_scan_folder(self):
        """Pick a folder — all images in it are queued and scanned sequentially."""
        from PyQt5.QtWidgets import QFileDialog
        start = str(Path(self._iphone_photo_dir)) if getattr(self, "_iphone_photo_dir", "") else ""
        folder = QFileDialog.getExistingDirectory(self, "Select folder to scan", start)
        if not folder:
            return
        exts = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}
        paths = sorted(p for p in Path(folder).iterdir()
                       if p.suffix.lower() in exts)
        if not paths:
            self._log(f"📁  No images found in {folder}", "warning")
            return
        self._batch_queue = [str(p) for p in paths]
        self._batch_total = len(self._batch_queue)
        self._log(f"📁  Folder loaded: {len(self._batch_queue)} images from {Path(folder).name} — hit SCAN to start", "success")
        self.folder_btn.setText(f"📁 {len(self._batch_queue)} files")
        self.folder_btn.setToolTip(folder)
        self._set_batch_scan_mode(f"FOLDER  {len(self._batch_queue)} images  ·  {Path(folder).name}")
        self._stop_live_camera()
        self._live_paused = True
        self._update_preview_toggle()
        self._set_preview(self._batch_queue[0])

    def _start_scan(self):
        # Grab and immediately clear forced path — batch queue always wins
        forced_image_path = getattr(self, "_forced_image_path", None)
        self._forced_image_path = None
        self._show_last_snap()  # switch preview to Last when scan begins
        if self.worker and self.worker.isRunning(): return
        self._log("\n" + "─"*50, "info")
        self._log(f"🕐  {datetime.now().strftime('%H:%M:%S')} — Starting scan...", "info")
        self.scan_btn.setEnabled(True)   # keep enabled so cancel click works
        self.scan_btn.setText("■  CLICK TO CANCEL")
        if hasattr(self, "cancel_btn"): self.cancel_btn.setEnabled(True)
        self.scan_btn.setStyleSheet("""
            QPushButton { background: #5a1a1a; color: white; font-size: 22px;
            font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background: #7a2020; }
        """)
        self.load_img_btn.setEnabled(False)
        self._scan_proc = None
        self.copy_btn.setEnabled(False)
        self.clip_btn.setEnabled(False)
        self.html_btn.setEnabled(False)
        self.merge_btn.setEnabled(False)
        self.upload_btn.setEnabled(False)
        self.status_label.setText("Scanning...")
        self.last_xml_path = None; self.last_html_path = None
        if hasattr(self, "bs_btn"): self.bs_btn.setEnabled(False)

        # Batch queue takes priority over single forced image
        batch_queue = getattr(self, "_batch_queue", [])
        batch_total = getattr(self, "_batch_total", 0)
        if batch_queue:
            forced_image_path = batch_queue[0]
            self._batch_queue = batch_queue[1:]
            remaining = len(self._batch_queue)
            done = batch_total - remaining
            self.scan_btn.setText(f"⏳  Image {done}/{batch_total}")
            self._log(f"📂  Batch scan {done}/{batch_total}: {Path(forced_image_path).name}", "info")
            self._set_preview(forced_image_path)
            # Only clear results on the very first image — accumulate for rest
            if done == 1:
                self._clear_results()
        else:
            self.load_img_btn.setText("📂 Image")
            self._clear_results()
        settings = self._get_settings()
        if forced_image_path:
            settings["forced_image_path"] = forced_image_path
        self.worker = ScanWorker(settings, self)
        self.worker.log.connect(self._log)
        self.worker.detected_count.connect(self._on_detected_count)
        self.worker.preview.connect(self._set_preview)
        self.worker.preview.connect(self._save_snap_path)
        self.worker.step.connect(lambda s: self.status_label.setText(s))
        self.worker.part_found.connect(self._add_result_row)
        self.worker.finished.connect(self._scan_finished)
        self.worker.start()

    def _scan_finished(self, success, xml_path):
        # If batch queue has more images, auto-fire next scan
        if getattr(self, "_batch_queue", []):
            self._log(f"📂  Auto-advancing to next image ({len(self._batch_queue)} remaining)...", "info")
            QTimer.singleShot(400, self._start_scan)
            return
        # Batch complete — reset batch buttons
        if getattr(self, "_batch_total", 0) > 1:
            self._log(f"✅  Batch complete — {self._batch_total} images scanned", "success")
            self._batch_total = 0
            self.batch_btn.setText("📂 Batch")
            self.folder_btn.setText("📁 Folder")
        self.scan_btn.setEnabled(True)
        self.scan_btn.setText("▶  SCAN")
        if hasattr(self, "cancel_btn"): self.cancel_btn.setEnabled(False)
        self.scan_btn.setText("▶  SCAN")
        self.scan_btn.setStyleSheet("""
            QPushButton { background: #1a3a1a; color: white; font-size: 22px;
            font-weight: bold; border-radius: 6px; border: none; }
            QPushButton:hover { background: #2a5a2a; }
        """)
        self.load_img_btn.setEnabled(True)
        self.setEnabled(True)  # re-enable window in case it got stuck
        if success and xml_path:
            self.last_xml_path = xml_path
            # Auto-merge duplicates if enabled
            if getattr(self, "_auto_merge_on", True):
                self._merge_duplicate_lots(silent=True)
            html_p = Path(xml_path).parent / "scan-results.html"
            if html_p.exists():
                self.last_html_path = str(html_p); self.html_btn.setEnabled(True)
            # Populate weight part ID dropdown (only if weight counter widget exists)
            if hasattr(self, "weight_part_id"):
                jp = Path(xml_path).parent / "scan-results.json"
                if jp.exists():
                    try:
                        seen = set()
                        self.weight_part_id.blockSignals(True)
                        self.weight_part_id.clear()
                        for p in json.loads(jp.read_text()):
                            pid = p.get("part_id")
                            if pid and pid not in seen:
                                seen.add(pid)
                                self.weight_part_id.addItem(f"{pid} — {p.get('part_name','')[:20]}", pid)
                        self.weight_part_id.blockSignals(False)
                    except Exception as e:
                        self._log(f"⚠  Weight part list: {e}", "warning")
            self._copy_xml()
            self.copy_btn.setEnabled(True)
            self.clip_btn.setEnabled(True)
            self.merge_btn.setEnabled(True)
            self.upload_btn.setEnabled(True)
            self.bs_btn.setEnabled(True)
            self.merge_r2.setEnabled(True)
            if hasattr(self, "bo_btn"): self.bo_btn.setEnabled(True)
            self.status_label.setText("✓ Done — XML copied to clipboard")
            self._log("📋  XML copied to clipboard.", "success")
        else:
            self.status_label.setText("✗ Scan failed — check console")

    def _build_xml_from_rows(self, condition="U"):
        """Build BrickLink-compatible XML for the Upload Inventory page.
        Format follows: https://www.bricklink.com/help.asp?helpID=207
        """
        active = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        if not active:
            return None
        lines = ["<INVENTORY>"]
        for r in active:
            item_type = "M" if r.get("item_type") == "M" else "P"
            color_id  = r.get("color_id") or 0
            qty       = r.get("qty", 1)
            cond      = r.get("condition") or condition  # per-row, fallback to global
            pid       = r.get("part_id", "")
            price     = r.get("price")
            lines.append("  <ITEM>")
            lines.append(f"    <ITEMTYPE>{item_type}</ITEMTYPE>")
            lines.append(f"    <ITEMID>{pid}</ITEMID>")
            if item_type == "P":
                lines.append(f"    <COLOR>{color_id}</COLOR>")
            lines.append(f"    <QTY>{qty}</QTY>")
            lines.append(f"    <CONDITION>{cond}</CONDITION>")
            if price is not None and price > 0:
                lines.append(f"    <PRICE>{price:.3f}</PRICE>")
            remark = r.get("remark","") or ""
            if remark:
                lines.append(f"    <REMARKS>{remark}</REMARKS>")
            comment = r.get("comment","") or ""
            if comment:
                lines.append(f"    <DESCRIPTION>{comment}</DESCRIPTION>")
            lines.append("  </ITEM>")
        lines.append("</INVENTORY>")
        return "\n".join(lines)

    def _copy_xml_to_clipboard(self):
        """Copy XML to clipboard using per-row condition (set on each row or bulk)."""
        from PyQt5.QtWidgets import QApplication
        xml = self._build_xml_from_rows()
        if not xml:
            self._log("No results to copy", "warning"); return
        QApplication.clipboard().setText(xml)
        n = sum(1 for r in self._rows if not r.get("_deleted") and r.get("part_id"))
        self._log(f"📋  XML copied ({n} items) — paste into BrickLink → Sell → Upload Inventory", "success")

    def _build_bsx(self):
        """Build a BrickStock/BrickStore BSX file using the correct format."""
        active = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        if not active:
            return None
        lines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<!DOCTYPE BrickStockXML>',
            '<BrickStockXML>',
            '<Inventory>',
        ]
        for r in active:
            itype    = "M" if r.get("item_type") == "M" else "P"
            color_id = r.get("color_id") or 0
            qty      = r.get("qty", 1)
            cond     = r.get("condition") or "U"
            pid      = r.get("part_id", "")
            price    = r.get("price") or 0
            name     = r.get("part_name", "") or ""
            lines.append("  <Item>")
            lines.append(f"    <ItemID>{pid}</ItemID>")
            lines.append(f"    <ItemTypeID>{itype}</ItemTypeID>")
            lines.append(f"    <ColorID>{color_id}</ColorID>")
            if name:
                lines.append(f"    <ItemName>{name}</ItemName>")
            lines.append(f"    <Qty>{qty}</Qty>")
            lines.append(f"    <Price>{price:.4f}</Price>")
            lines.append(f"    <Condition>{cond}</Condition>")
            remark = r.get("remark", "") or ""
            if remark:
                lines.append(f"    <Remarks>{remark}</Remarks>")
            comment = r.get("comment", "") or ""
            if comment:
                lines.append(f"    <Comments>{comment}</Comments>")
            lines.append("  </Item>")
        lines += ["</Inventory>", "</BrickStockXML>"]
        return "\n".join(lines)

    def _open_in_brickstore(self):
        """Save as BSX (BrickStore native format) and open — preserves CAD currency."""
        import tempfile
        bsx = self._build_bsx()
        if not bsx:
            self._log("No results to open", "warning"); return
        try:
            tmp = tempfile.NamedTemporaryFile(
                suffix=".bsx", prefix="scanstation-", delete=False)
            tmp.write(bsx.encode("utf-8")); tmp.close()
            tmp_path = tmp.name
        except Exception as e:
            self._log(f"Temp file error: {e}", "error"); return
        # Common BrickStore install paths on Windows
        import glob
        candidates = [
            "C:/Program Files/BrickStore/brickstore.exe",
            "C:/Program Files (x86)/BrickStore/brickstore.exe",
            str(Path.home() / "AppData/Local/Programs/BrickStore/brickstore.exe"),
        ]
        # Also search common install dirs
        candidates += glob.glob("C:/Program Files*/BrickStore*/brickstore.exe")
        bs_path = next((c for c in candidates if Path(c).exists()), None)
        if bs_path:
            subprocess.Popen([bs_path, tmp_path])
            self._log("🧱  Opened in BrickStore", "success")
        else:
            # Fallback: let Windows open with whatever handles .xml
            os.startfile(tmp_path)
            self._log("🧱  BrickStore not found — opened with default app. Set path in code if needed.", "info")

    def _download_xml(self):
        from PyQt5.QtWidgets import QFileDialog
        from datetime import datetime
        xml_content = self._build_xml_from_rows()
        if not xml_content and not self.last_xml_path:
            self._log("No results to export", "warning"); return
        default_name = f"bricklink-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xml"
        dest, _ = QFileDialog.getSaveFileName(self, "Save XML", default_name, "XML Files (*.xml)")
        if dest:
            try:
                if xml_content:
                    Path(dest).write_text(xml_content, encoding="utf-8")
                else:
                    import shutil; shutil.copy2(self.last_xml_path, dest)
                self._log(f"💾  XML saved to {dest}", "success")
            except Exception as e:
                self._log(f"Could not save XML: {e}", "error")

    def _copy_xml(self):
        try:
            xml_content = self._build_xml_from_rows()
            if not xml_content and self.last_xml_path:
                xml_content = Path(self.last_xml_path).read_text(encoding="utf-8")
            if not xml_content:
                self._log("No results to copy", "warning"); return
            QApplication.clipboard().setText(xml_content)
            self._log("📋  XML copied.", "success")
            self.status_label.setText("✓ XML copied — paste into BrickStore")
        except Exception as e:
            self._log(f"Copy failed: {e}", "error")

    # ── BrickOwl export ───────────────────────────────────────────────────────
    _BO_COND_MAP = {"U": "Used", "N": "New"}
    _bo_boid_cache: dict = {}

    def _build_brickowl_csv(self) -> str:
        import csv, io
        active = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        if not active:
            return ""
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(["BOID", "Quantity", "Condition", "My Price", "My Remarks", "My Notes"])
        for r in active:
            pid    = r.get("part_id", "")
            qty    = r.get("qty", 1)
            cond   = self._BO_COND_MAP.get(r.get("condition", "U"), "Used")
            price  = f"{r['price']:.3f}" if r.get("price") else ""
            remark = r.get("remark", "") or ""
            notes  = r.get("comment", "") or ""
            boid   = self._bo_lookup_boid(pid, r.get("color_id", 0), r.get("item_type", "P"))
            writer.writerow([boid or pid, qty, cond, price, remark, notes])
        return buf.getvalue()

    def _bo_lookup_boid(self, pid: str, color_id: int, item_type: str) -> str:
        key = (pid.lower(), color_id, item_type)
        if key in self._bo_boid_cache:
            return self._bo_boid_cache[key]
        try:
            import requests as _req
            bl_type = {"M": "Minifig", "P": "Part", "S": "Set"}.get(item_type, "Part")
            resp = _req.get("https://api.brickowl.com/v1/catalog/id_lookup",
                            params={"key": "guest", "type": bl_type,
                                    "bl_id": pid, "color_id": color_id}, timeout=6)
            if resp.status_code == 200:
                data = resp.json()
                matches = data if isinstance(data, list) else data.get("results", [])
                if matches:
                    boid = str(matches[0].get("boid", ""))
                    self._bo_boid_cache[key] = boid
                    return boid
        except Exception:
            pass
        self._bo_boid_cache[key] = ""
        return ""

    def _export_brickowl_csv(self):
        active = [r for r in self._rows if not r.get("_deleted") and r.get("part_id")]
        if not active:
            self._log("No results to export", "warning"); return
        self._log("🦉  Building BrickOwl CSV — looking up BOIDs...", "info")

        def _do():
            csv_content = self._build_brickowl_csv()
            if not csv_content:
                self.log_message.emit("No results to export", "warning"); return
            self.log_message.emit(f"🦉  BrickOwl CSV ready ({len(active)} rows)", "success")
            from PyQt5.QtWidgets import QFileDialog
            from datetime import datetime
            default = f"brickowl-{datetime.now().strftime('%Y%m%d-%H%M%S')}.csv"
            dest, _ = QFileDialog.getSaveFileName(self, "Save BrickOwl CSV", default, "CSV Files (*.csv)")
            if dest:
                try:
                    Path(dest).write_text(csv_content, encoding="utf-8-sig")
                    self.log_message.emit(f"🦉  Saved to {dest}", "success")
                except Exception as e:
                    self.log_message.emit(f"Could not save: {e}", "error")

        threading.Thread(target=_do, daemon=True).start()

    def _open_html(self):
        if not self.last_html_path: return
        try:
            if platform.system() == "Windows":
                os.startfile(self.last_html_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", self.last_html_path])
            else:
                subprocess.Popen(["xdg-open", self.last_html_path])
        except Exception as e:
            self._log(f"Could not open report: {e}", "error")


# ── Entry ─────────────────────────────────────────────────────────────────────
def exception_hook(exc_type, exc_value, exc_tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, exc_tb)
    # Don't exit — just log it

sys.excepthook = exception_hook

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.Window,          QColor(DARK_BG))
    pal.setColor(QPalette.WindowText,      QColor(TEXT))
    pal.setColor(QPalette.Base,            QColor(PANEL_BG))
    pal.setColor(QPalette.AlternateBase,   QColor(ROW_ALT))
    pal.setColor(QPalette.Highlight,       QColor(ACCENT2))
    pal.setColor(QPalette.HighlightedText, QColor("#000000"))
    pal.setColor(QPalette.Text,            QColor(TEXT))
    pal.setColor(QPalette.Button,          QColor(PANEL_BG))
    pal.setColor(QPalette.ButtonText,      QColor(TEXT))
    pal.setColor(QPalette.Highlight,       QColor(ACCENT))
    pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(pal)

    # Force high-contrast tooltip style — Windows/Fusion ignores widget-level QToolTip
    app.setStyleSheet("""
        QToolTip {
            background-color: #1a1a1a;
            color: #f0f0f0;
            border: 1px solid #555555;
            padding: 4px 6px;
            font-size: 11px;
            border-radius: 3px;
        }
    """)

    win = ScanStation()
    win.show()
    app.exec_()  # don't pass return code to sys.exit


if __name__ == "__main__":
    main()
