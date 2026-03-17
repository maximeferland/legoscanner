#!/usr/bin/env python3
"""
scan-heads.py — Brickognize Head Scanner (skips color detection, outputs Yellow)
==========================================
1. Takes a photo of LEGO parts laid out on a white sheet
2. Auto-detects each part by contrast
3. Crops each part, sends to Brickognize API
4. Detects dominant color and maps to BrickLink color ID
5. Exports XML (BrickStore) + TSV (BrickLink upload page)

Usage:
    py scan-sheet.py photo.jpg
    py scan-sheet.py photo.jpg --output reports/
    py scan-sheet.py photo.jpg --confidence 0.7
    py scan-sheet.py photo.jpg --debug   (saves cropped images for inspection)
"""

import argparse
import asyncio
import aiohttp
import json
import os
import sys
import re
import requests
from requests_oauthlib import OAuth1
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path
from datetime import datetime
import colorsys

try:
    from PIL import Image, ImageFilter
    import numpy as np
except ImportError:
    print("Missing dependencies. Run: pip install Pillow numpy --break-system-packages")
    sys.exit(1)

# ─── BrickLink Color Map (ID → (name, avg_rgb)) ───────────────────────────────
# fmt: off
# ── Rebrickable parts.csv lookup ─────────────────────────────────────────────
_RB_PARTS: dict = {}

def _load_rb_parts_csv():
    """Load parts.csv from script directory into _RB_PARTS. Call once at startup."""
    global _RB_PARTS
    csv_path = Path(__file__).parent / "parts.csv"
    if not csv_path.exists():
        return
    import csv as _csv
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
        print(f"   📚  parts.csv: {count:,} parts loaded")
    except Exception as e:
        print(f"   ⚠  parts.csv: {e}")

def rb_part_name(part_num: str) -> str:
    """Part name from Rebrickable CSV, or ''."""
    return _RB_PARTS.get(part_num.lower().strip(), {}).get("name", "")

def rb_part_material(part_num: str) -> str:
    """Part material from Rebrickable CSV (e.g. 'Plastic', 'Rubber')."""
    return _RB_PARTS.get(part_num.lower().strip(), {}).get("material", "")

# ── BrickLink color table ─────────────────────────────────────────────────────
BRICKLINK_COLORS = {
    # ── Solid — White / Grays / Black ──────────────────────────
       1: ("White",                       (242,243,242)),
      49: ("Very Light Gray",             (225,225,225)),
      99: ("Very Light Bluish Gray",      (228,232,232)),
      86: ("Light Bluish Gray",           (171,175,175)),
       9: ("Light Gray",                  (156,156,156)),
      10: ("Dark Gray",                   (107, 98, 90)),
      85: ("Dark Bluish Gray",            ( 89, 93, 96)),
      11: ("Black",                       ( 33, 33, 33)),
    # ── Solid — Reds ───────────────────────────────────────────
      59: ("Dark Red",                    (106, 14, 21)),
       5: ("Red",                         (177,  0,  6)),
     167: ("Reddish Orange",              (210, 70, 30)),
     231: ("Dark Salmon",                 (210,110, 90)),
      25: ("Salmon",                      (244, 92, 64)),
     220: ("Coral",                       (242,112, 94)),
      26: ("Light Salmon",                (255,188,180)),
      58: ("Sand Red",                    (149, 86, 83)),
    # ── Solid — Browns ─────────────────────────────────────────
     120: ("Dark Brown",                  ( 53, 33,  0)),
     168: ("Umber",                       ( 90, 60, 35)),
       8: ("Brown",                       (117, 78, 36)),
      88: ("Reddish Brown",               (136, 53, 33)),
      91: ("Light Brown",                 (181,132, 83)),
     240: ("Medium Brown",                (155,110, 72)),
     106: ("Fabuland Brown",              (160, 95, 52)),
    # ── Solid — Tans / Nougats ─────────────────────────────────
      69: ("Dark Tan",                    (137,125, 87)),
       2: ("Tan",                         (222,196,152)),
      90: ("Light Nougat",                (255,213,174)),
     241: ("Medium Tan",                  (215,185,142)),
      28: ("Nougat",                      (216,127, 77)),
     150: ("Medium Nougat",               (175,116, 70)),
     225: ("Dark Nougat",                 (145, 80, 40)),
     169: ("Sienna",                      (165, 90, 50)),
    # ── Solid — Oranges / Yellows ──────────────────────────────
     160: ("Fabuland Orange",             (220,140, 80)),
      29: ("Earth Orange",                (168,120, 60)),
      68: ("Dark Orange",                 (157, 82, 28)),
      27: ("Rust",                        (181, 44, 32)),
     165: ("Neon Orange",                 (255,140, 50)),
       4: ("Orange",                      (209,109, 27)),
      31: ("Medium Orange",               (235,158, 50)),
      32: ("Light Orange",                (252,172,120)),
     110: ("Bright Light Orange",         (247,186, 48)),
     172: ("Warm Yellowish Orange",       (240,175, 80)),
      96: ("Very Light Orange",           (255,218,170)),
     161: ("Dark Yellow",                 (200,160, 40)),
     173: ("Ochre Yellow",                (210,170, 60)),
       3: ("Yellow",                      (243,195, 60)),
      33: ("Light Yellow",                (255,240,188)),
     103: ("Bright Light Yellow",         (255,236, 61)),
     236: ("Neon Yellow",                 (225,240, 30)),
     171: ("Lemon",                       (250,240, 70)),
    # ── Solid — Greens ─────────────────────────────────────────
     166: ("Neon Green",                  (170,235, 50)),
      35: ("Light Lime",                  (196,234,166)),
     158: ("Yellowish Green",             (194,224, 80)),
      76: ("Medium Lime",                 (199,210,113)),
      34: ("Lime",                        (187,233, 11)),
     248: ("Fabuland Lime",               (175,225,140)),
     155: ("Olive Green",                 (119,119, 78)),
     242: ("Dark Olive Green",            ( 80, 80, 40)),
      80: ("Dark Green",                  ( 25, 89, 55)),
       6: ("Green",                       ( 35,120, 65)),
      36: ("Bright Green",                ( 75,151, 74)),
      37: ("Medium Green",                (132,182,141)),
      38: ("Light Green",                 (167,228,164)),
      48: ("Sand Green",                  (118,150,125)),
    # ── Solid — Turquoise / Aqua ───────────────────────────────
      39: ("Dark Turquoise",              (  0,143,155)),
      40: ("Light Turquoise",             ( 85,185,175)),
      41: ("Aqua",                        (173,220,212)),
     152: ("Light Aqua",                  (204,242,233)),
    # ── Solid — Blues ──────────────────────────────────────────
      63: ("Dark Blue",                   ( 20, 48,104)),
       7: ("Blue",                        (  0, 87,166)),
     153: ("Dark Azure",                  ( 70,155,195)),
     247: ("Little Robots Blue",          (100,175,220)),
      72: ("Maersk Blue",                 (111,165,210)),
     156: ("Medium Azure",                (104,195,226)),
      87: ("Sky Blue",                    (125,188,227)),
      42: ("Medium Blue",                 ( 98,162,212)),
     105: ("Bright Light Blue",           (159,195,233)),
      62: ("Light Blue",                  (180,210,228)),
      55: ("Sand Blue",                   ( 90,113,132)),
     109: ("Dark Royal Blue",             ( 48, 68,145)),
    # ── Solid — Purples / Violets ──────────────────────────────
      43: ("Violet",                      ( 65, 97,165)),
      97: ("Royal Blue",                  ( 75, 90,180)),
     245: ("Lilac",                       (150,130,195)),
     174: ("Blue Violet",                 ( 92,112,190)),
      73: ("Medium Violet",               (138,124,183)),
     246: ("Light Lilac",                 (200,185,225)),
      44: ("Light Violet",                (180,170,215)),
      89: ("Dark Purple",                 ( 67, 28, 93)),
      24: ("Purple",                      (105, 46,119)),
      93: ("Light Purple",                (180,100,160)),
     157: ("Medium Lavender",             (170,133,196)),
     154: ("Lavender",                    (193,159,217)),
     227: ("Clikits Lavender",            (190,150,210)),
      54: ("Sand Purple",                 (143,114,158)),
    # ── Solid — Pinks / Magentas ───────────────────────────────
      71: ("Magenta",                     (181, 41, 82)),
      47: ("Dark Pink",                   (200,112,128)),
      94: ("Medium Dark Pink",            (208,108,156)),
     104: ("Bright Pink",                 (240,150,190)),
      23: ("Pink",                        (246,169,187)),
      56: ("Rose Pink",                   (220,150,165)),
     175: ("Warm Pink",                   (250,170,185)),
    # ── Trans ──────────────────────────────────────────────────
      12: ("Trans-Clear",                 (236,236,236)),
      13: ("Trans-Brown",                 (120,100, 80)),
     251: ("Trans-Black",                 ( 60, 60, 60)),
      17: ("Trans-Red",                   (195, 62, 50)),
      18: ("Trans-Neon Orange",           (255, 66, 49)),
      98: ("Trans-Orange",                (215,128, 25)),
     164: ("Trans-Light Orange",          (240,180, 80)),
     121: ("Trans-Neon Yellow",           (255,215,  0)),
      19: ("Trans-Yellow",                (245,205, 47)),
      16: ("Trans-Neon Green",            (170,235, 50)),
     108: ("Trans-Bright Green",          ( 60,179,113)),
     221: ("Trans-Light Green",           (180,230,180)),
     226: ("Trans-Light Bright Green",    (200,245,180)),
      20: ("Trans-Green",                 ( 84,194,104)),
      14: ("Trans-Dark Blue",             (  0, 96,175)),
      74: ("Trans-Medium Blue",           ( 80,145,210)),
      15: ("Trans-Light Blue",            (174,233,251)),
     113: ("Trans-Aqua",                  (180,235,230)),
     114: ("Trans-Light Purple",          (210,180,225)),
     234: ("Trans-Medium Purple",         (160,130,200)),
      51: ("Trans-Purple",                (100, 50,150)),
      50: ("Trans-Dark Pink",             (205, 98,152)),
     107: ("Trans-Pink",                  (230,160,200)),
    # ── Chrome ─────────────────────────────────────────────────
      21: ("Chrome Gold",                 (200,165, 70)),
      22: ("Chrome Silver",               (190,195,200)),
      57: ("Chrome Antique Brass",        (155,130, 75)),
     122: ("Chrome Black",                ( 50, 50, 50)),
      52: ("Chrome Blue",                 ( 80,130,200)),
      64: ("Chrome Green",                ( 50,140, 80)),
      82: ("Chrome Pink",                 (220,130,160)),
    # ── Pearl / Metallic ───────────────────────────────────────
      83: ("Pearl White",                 (240,240,240)),
     119: ("Pearl Very Light Gray",       (218,218,218)),
      66: ("Pearl Light Gray",            (171,173,172)),
      95: ("Flat Silver",                 (138,146,141)),
     239: ("Bionicle Silver",             (130,140,150)),
      77: ("Pearl Dark Gray",             (100,104,102)),
     244: ("Pearl Black",                 ( 40, 40, 40)),
      61: ("Pearl Light Gold",            (210,188,120)),
     115: ("Pearl Gold",                  (220,188,100)),
     235: ("Reddish Gold",                (195,155, 80)),
     238: ("Bionicle Gold",               (200,170, 60)),
      81: ("Flat Dark Gold",              (161,135, 57)),
     249: ("Reddish Copper",              (180,100, 70)),
      84: ("Copper",                      (174,122, 73)),
     237: ("Bionicle Copper",             (170,100, 65)),
     255: ("Pearl Brown",                 (150,105, 70)),
     252: ("Pearl Red",                   (185, 60, 55)),
     253: ("Pearl Green",                 ( 60,130, 80)),
     254: ("Pearl Blue",                  ( 60,100,175)),
      78: ("Pearl Sand Blue",             (116,134,157)),
     243: ("Pearl Sand Purple",           (143,114,158)),
    # ── Milky / Glow / Special ─────────────────────────────────
     183: ("Metallic White",              (240,240,240)),
     139: ("Metallic Green",              ( 80,140, 90)),
     145: ("Metallic Sand Blue",          (100,130,155)),
     129: ("Glow in Dark White",          (216,221,212)),
     294: ("Glow in Dark Trans",          (189,198,173)),
      46: ("Glow in Dark Opaque",         (205,215,185)),
}
# fmt: on


# Minifig component base part IDs — Brickognize returns these for printed variants
# e.g. it may return "3626c" for any printed head, "973c00" for any torso body
# BrickLink minifig catalog ID prefixes — used to detect whole minifigs that
# Brickognize returns with type="P" but are actually in BL's Minifigs catalog.
# Rule: if the part_id starts with any of these (case-insensitive), treat as
# item_type="M" for all API calls (price, URL, XML).
MINIFIG_CATALOG_PREFIXES = (
    "ac",
    "adv",
    "agt",
    "alp",
    "arc",
    "atl",
    "bat",
    "cas",
    "cca",
    "cc",
    "cl",
    "col",
    "coldnd",
    "colhp",
    "colmar",
    "collon",
    "coltlm",
    "colnin",
    "cre",
    "cty",
    "dim",
    "dis",
    "dp",
    "edu",
    "elf",
    "fab",
    "fst",
    "frnd",
    "gal",
    "gen",
    "har",
    "hol",
    "hob",
    "hp",
    "hs",
    "ice",
    "idea",
    "ind",
    "jw",
    "lor",
    "mk",
    "min",
    "mof",
    "mvl",
    "nba",
    "nin",
    "njo",
    "njr",
    "ora",
    "pac",
    "pi",
    "pm",
    "potc",
    "prince",
    "pur",
    "rac",
    "res",
    "sc",
    "sh",
    "she",
    "shf",
    "soc",
    "spd",
    "spj",
    "sw",
    "tlbm",
    "tlm",
    "tlnm",
    "toy",
    "twn",
    "vik",
    "wc",
    "ww",
)

MINIFIG_PART_PREFIXES = (
    # ── Body structure ────────────────────────────────────────────────────────
    "3626",   # Head (3626c, 3626cpb*)
    "973",    # Torso + arms assembly (973c*, 973pb*)
    "970",    # Hips + legs assembly (970c*)
    "971",    # Right leg
    "972",    # Left leg
    "981",    # Right arm
    "982",    # Left arm
    "983",    # Hand
    "74261",  # Head short/child variant
    "3901",   # Hand alternative
    # ── Short / special legs ─────────────────────────────────────────────────
    "92250",  # Short leg alternative
    "87819",  # Leg right short
    "41879",  # Short leg variant
    "37364",  # Leg short (Friends)
    "24782",  # Leg short (newer moulds)
    # ── Accessories commonly part of a minifig listing ───────────────────────
    "90541",  # Neckwear / collar
    "x749",   # Cloth cape
    "18964",  # Visor
    "2446",   # Hat / headgear
    "3844",   # Glasses
    "30175",  # Hair
    "hairp",  # Hair parts with hairp prefix
    "hpb",    # Decorated hair
)


def estimate_stud_area(name: str) -> float:
    """
    Parse a BrickLink part name and return an estimated stud-footprint area.
    e.g. "Brick 2 x 4" → 8.0,  "Tile 1 x 2" → 2.0,  "Plate 2 x 2" → 4.0
    Returns 0.0 if unparseable.
    """
    import re
    # Match patterns like "2 x 4", "1X2", "2x2x3" (take first two numbers)
    m = re.search(r"(\d+)\s*[xX×]\s*(\d+)", name)
    if m:
        return float(m.group(1)) * float(m.group(2))
    return 0.0


def box_size_ratio(box_wh: tuple, all_boxes_wh: list) -> float:
    """
    Return how large this box is relative to the median box area.
    A 2x4 brick next to 1x2 tiles should give ratio ≈ 4.0.
    """
    if not all_boxes_wh:
        return 1.0
    areas = [bw * bh for bw, bh in all_boxes_wh]
    median_area = sorted(areas)[len(areas) // 2]
    if median_area == 0:
        return 1.0
    this_area = box_wh[0] * box_wh[1]
    return this_area / median_area


def size_score_penalty(candidate_name: str, size_ratio: float) -> float:
    """
    Return a penalty [0.0 – 0.35] to subtract from a candidate's confidence
    when its implied stud area contradicts the observed crop size ratio.

    Rules:
    - If candidate stud area is unparseable → no penalty (0.0)
    - Expected stud ratio ≈ candidate_studs / median_studs (2.0 studs median assumed)
    - If observed size_ratio vs expected differ by >2× → penalty 0.35
    - Graduated linearly between 1× and 2× mismatch → 0.0 – 0.35
    """
    studs = estimate_stud_area(candidate_name)
    if studs == 0:
        return 0.0
    # Assume median part is ~2 studs (common 1×2 tiles/plates)
    MEDIAN_STUDS = 2.0
    expected_ratio = studs / MEDIAN_STUDS
    if expected_ratio == 0:
        return 0.0
    # How far off is the observed ratio from expected?
    mismatch = max(size_ratio / expected_ratio, expected_ratio / size_ratio)
    # mismatch=1 → no penalty; mismatch=3 → max penalty
    penalty = min(0.35, max(0.0, (mismatch - 1.0) / 2.0 * 0.35))
    return penalty

# Cache for BL-verified item types: {part_id: "M" or "P"}
_bl_item_type_cache: dict = {}

def verify_item_type_with_bl(part_id: str, creds: dict) -> str:
    """Query BrickLink API to definitively determine if an ID is a Minifig or Part.
    Returns "M" or "P". Caches result to avoid repeat calls.
    Falls back to prefix-based guess if no credentials or API error.
    """
    pid = part_id.lower().strip()
    if pid in _bl_item_type_cache:
        return _bl_item_type_cache[pid]
    if not creds:
        result = "M" if is_minifig_catalog(part_id) else "P"
        return result
    try:
        # Try minifig endpoint first
        url = f"https://api.bricklink.com/api/store/v1/items/minifig/{pid}"
        resp = requests.get(url, auth=_get_oauth(creds), timeout=5)
        if resp.status_code == 200:
            _bl_item_type_cache[pid] = "M"
            return "M"
        elif resp.status_code == 404:
            _bl_item_type_cache[pid] = "P"
            return "P"
        # Other error — fall back to prefix guess
    except Exception:
        pass
    result = "M" if is_minifig_catalog(part_id) else "P"
    _bl_item_type_cache[pid] = result
    return result

def is_minifig_catalog(part_id: str) -> bool:
    """Return True if this ID is a BrickLink Minifigs-catalog item (ITEMTYPE=M).
    These are whole/assembled minifigs sold under the Minifigs catalog, not Parts.
    e.g. sw0001, col323, sh0735, hp539, cty0452
    Uses prefix list as fast first check. For unknown prefixes, caller should
    use verify_item_type_with_bl() for a definitive API-based answer.
    """
    pid = part_id.lower().strip()
    if not pid or pid[0].isdigit():
        return False
    return any(pid.startswith(p.lower()) for p in MINIFIG_CATALOG_PREFIXES)


def is_minifig_part(part_id: str) -> bool:
    """Return True if this part ID looks like a minifig body part (sold as Parts).
    Covers numeric prefixes (971, 972…) and specific letter prefixes (hairp*, hpb*).
    Note: sh* is in MINIFIG_CATALOG_PREFIXES — those are whole figures, not parts.
    """
    pid = part_id.lower().strip()
    if pid and not pid[0].isdigit():
        return any(pid.startswith(p.lower()) for p in MINIFIG_PART_PREFIXES
                   if not p[0].isdigit())
    return any(pid.startswith(p.lower()) for p in MINIFIG_PART_PREFIXES
               if p[0].isdigit())


def _rgb_to_lab(rgb):
    """Convert sRGB (0-255 each) to CIE Lab (D65 illuminant)."""
    r, g, b = [x / 255.0 for x in rgb]
    def lin(c): return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = lin(r), lin(g), lin(b)
    X = r*0.4124564 + g*0.3575761 + b*0.1804375
    Y = r*0.2126729 + g*0.7151522 + b*0.0721750
    Z = r*0.0193339 + g*0.1191920 + b*0.9503041
    def f(t): return t ** (1/3) if t > 0.008856 else 7.787 * t + 16/116
    L = 116 * f(Y / 1.00000) - 16
    a = 500 * (f(X / 0.95047) - f(Y / 1.00000))
    b_ = 200 * (f(Y / 1.00000) - f(Z / 1.08883))
    return L, a, b_


# Pre-compute Lab for every reference color — done once at import time
_BL_COLORS_LAB: dict = {}  # populated lazily on first use


def _ensure_lab_cache():
    if not _BL_COLORS_LAB:
        for cid, (name, ref_rgb) in BRICKLINK_COLORS.items():
            _BL_COLORS_LAB[cid] = _rgb_to_lab(ref_rgb)


def rgb_distance(c1, c2):
    """CIE76 ΔE perceptual distance in Lab color space.
    Correctly separates neutral colors (tan vs light gray vs white) that
    the old hue-weighted RGB metric confused.
    """
    L1, a1, b1 = _rgb_to_lab(c1)
    L2, a2, b2 = _rgb_to_lab(c2)
    return ((L1-L2)**2 + (a1-a2)**2 + (b1-b2)**2) ** 0.5


def dominant_color_from_image(img_pil):
    """Match part color to nearest BrickLink color using Lab ΔE.
    Uses sample_part_color_rgb for a consistent inner-crop sample.
    """
    _ensure_lab_cache()
    sampled = sample_part_color_rgb(img_pil)
    sampled_lab = _rgb_to_lab(sampled)
    best_id, best_name, best_dist = 1, "White", float("inf")
    for color_id, (name, _) in BRICKLINK_COLORS.items():
        ref_lab = _BL_COLORS_LAB[color_id]
        d = ((sampled_lab[0]-ref_lab[0])**2 +
             (sampled_lab[1]-ref_lab[1])**2 +
             (sampled_lab[2]-ref_lab[2])**2) ** 0.5
        if d < best_dist:
            best_dist = d
            best_id = color_id
            best_name = name
    return best_id, best_name


def auto_gap(raw_boxes):
    """
    Measure the typical spacing between parts in the photo.
    Strategy: for each blob, find its nearest neighbour (edge-to-edge distance).
    The median of those distances = the inter-part spacing.
    We use half that as the merge threshold — blobs closer than half the
    typical inter-part gap must be fragments of the same part.
    """
    if len(raw_boxes) < 2:
        return 20  # fallback for very sparse sheets

    def edge_dist(a, b):
        ax1, ay1, ax2, ay2 = a
        bx1, by1, bx2, by2 = b
        dx = max(0, max(ax1, bx1) - min(ax2, bx2))
        dy = max(0, max(ay1, by1) - min(ay2, by2))
        return (dx*dx + dy*dy) ** 0.5

    nearest = []
    for i, box in enumerate(raw_boxes):
        dists = [edge_dist(box, raw_boxes[j]) for j in range(len(raw_boxes)) if j != i]
        nearest.append(min(dists))

    # Sort and take median of the lower half (ignores outliers from large empty areas)
    nearest.sort()
    lower_half = nearest[:max(1, len(nearest)//2)]
    median_gap = sorted(lower_half)[len(lower_half)//2]

    # Merge threshold = half the inter-part spacing
    # Clamp between 8px (never over-merge) and 120px (never under-merge)
    merge_gap = int(median_gap * 0.5)
    merge_gap = max(8, min(merge_gap, 120))
    return merge_gap, median_gap


def detect_parts_fixed_grid(img_pil, cols=8, rows=6):
    """
    Divide image into a fixed cols×rows grid — no detection needed.
    Every cell gets one crop, top-left to bottom-right order.
    """
    w, h = img_pil.size
    cell_w = w // cols
    cell_h = h // rows
    boxes = []
    for row in range(rows):
        for col in range(cols):
            x1 = col * cell_w
            y1 = row * cell_h
            x2 = min(w, x1 + cell_w)
            y2 = min(h, y1 + cell_h)
            boxes.append((x1, y1, x2, y2))
    print(f"   Fixed grid: {cols} cols × {rows} rows = {len(boxes)} cells")
    return boxes


def find_dominant_spacing(projection, min_spacing, max_spacing):
    """
    Use FFT to find the dominant periodic spacing in a 1D projection.
    Returns the most likely distance between evenly-spaced parts.
    """
    import numpy as np
    n = len(projection)
    fft = np.abs(np.fft.rfft(projection - projection.mean()))
    freqs = np.fft.rfftfreq(n)

    # Convert frequency range to period range
    # period = 1/freq, but we want periods in pixels
    best_period = None
    best_power = 0
    for i, freq in enumerate(freqs):
        if freq == 0:
            continue
        period = 1.0 / freq
        if min_spacing <= period <= max_spacing:
            if fft[i] > best_power:
                best_power = fft[i]
                best_period = period

    return int(round(best_period)) if best_period else None


def detect_parts_grid(img_pil):
    """
    Grid-based detection for evenly-spaced parts.
    Uses FFT on edge projections to automatically find the grid spacing —
    works regardless of how densely or sparsely parts are spread.
    """
    try:
        import cv2
    except ImportError:
        return None

    import numpy as np
    from scipy.signal import find_peaks

    w, h = img_pil.size
    img_arr = np.array(img_pil.convert("RGB"))

    # Edge detection — parts have sharp edges, cloth/carpet is smooth
    gray = cv2.cvtColor(img_arr, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 1)
    edges = cv2.Canny(blur, 30, 80)

    # Project onto axes
    proj_x = edges.sum(axis=0).astype(float)
    proj_y = edges.sum(axis=1).astype(float)

    # Smooth lightly to merge split-part activity
    for proj, length in [(proj_x, w), (proj_y, h)]:
        k = max(3, length // 80)
        kernel = np.ones(k) / k

    proj_x = np.convolve(proj_x, np.ones(max(3, w//80)) / max(3, w//80), mode='same')
    proj_y = np.convolve(proj_y, np.ones(max(3, h//80)) / max(3, h//80), mode='same')

    # Part size bounds: 3% to 30% of image
    min_part = int(min(w, h) * 0.03)
    max_part = int(min(w, h) * 0.30)

    # FFT to find dominant grid spacing automatically
    col_spacing = find_dominant_spacing(proj_x, min_part, max_part)
    row_spacing = find_dominant_spacing(proj_y, min_part, max_part)

    if not col_spacing or not row_spacing:
        print("   Grid: could not determine spacing from FFT")
        return None

    print(f"   Grid spacing: {col_spacing}px cols × {row_spacing}px rows (auto-detected)")

    # Find peaks using the auto-detected spacing as minimum distance
    col_peaks, _ = find_peaks(proj_x, distance=int(col_spacing * 0.6), height=proj_x.max() * 0.12)
    row_peaks, _ = find_peaks(proj_y, distance=int(row_spacing * 0.6), height=proj_y.max() * 0.12)

    print(f"   Grid: {len(col_peaks)} columns × {len(row_peaks)} rows = {len(col_peaks)*len(row_peaks)} cells")

    if len(col_peaks) < 2 or len(row_peaks) < 2:
        return None

    half_w = int(col_spacing * 0.52)
    half_h = int(row_spacing * 0.52)

    boxes = []
    for cy in row_peaks:
        for cx in col_peaks:
            x1 = max(0, cx - half_w)
            y1 = max(0, cy - half_h)
            x2 = min(w, cx + half_w)
            y2 = min(h, cy + half_h)
            boxes.append((x1, y1, x2, y2))

    return boxes


def detect_parts_by_color_sample(img_pil):
    """
    Detect heads by sampling the dominant color from the center of the image
    (where a head is most likely to be), then finding all regions that match.
    Works for any head color on any background — no manual tuning needed.
    """
    try:
        import cv2
    except ImportError:
        return None

    import numpy as np
    w, h = img_pil.size
    img_arr = np.array(img_pil.convert("RGB"))
    hsv = cv2.cvtColor(img_arr, cv2.COLOR_RGB2HSV)

    # Sample a small region near the image center — almost certainly a head
    # Use a 5% patch of the image center
    patch_size = max(20, int(min(w, h) * 0.05))
    cx, cy = w // 2, h // 2
    patch = hsv[cy-patch_size:cy+patch_size, cx-patch_size:cx+patch_size]
    
    # Get median HSV of the center patch — this is our target color
    target_h = int(np.median(patch[:,:,0]))
    target_s = int(np.median(patch[:,:,1]))
    target_v = int(np.median(patch[:,:,2]))
    
    # Build a tolerance mask around the sampled color
    h_tol = 15   # hue tolerance ±15
    s_tol = 60   # saturation tolerance
    v_tol = 60   # value/brightness tolerance
    
    lower = np.array([max(0, target_h - h_tol), max(0, target_s - s_tol), max(0, target_v - v_tol)])
    upper = np.array([min(179, target_h + h_tol), min(255, target_s + s_tol), min(255, target_v + v_tol)])
    
    mask = cv2.inRange(hsv, lower, upper)
    
    # Morphological ops to fill holes and separate touching heads
    kernel = np.ones((5, 5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=3)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=2)
    
    # Find contours
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Filter contours by size — heads are a specific size range
    min_side = min(w, h)
    min_area = (min_side * 0.025) ** 2
    max_area = (min_side * 0.20) ** 2
    
    boxes = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area or area > max_area:
            continue
        x, y, bw, bh = cv2.boundingRect(cnt)
        # Must be roughly square (head-shaped), not a long thin strip
        aspect = max(bw, bh) / max(min(bw, bh), 1)
        if aspect > 2.5:
            continue
        pad = int(max(bw, bh) * 0.30)
        boxes.append((
            max(0, x - pad), max(0, y - pad),
            min(w, x + bw + pad), min(h, y + bh + pad)
        ))
    
    print(f"   Color sample: HSV({target_h},{target_s},{target_v}) → {len(boxes)} regions found")
    return boxes


def detect_parts_circles(img_pil):
    """
    Detect LEGO heads using Hough Circle Transform.
    Auto-tunes threshold to find a reasonable number of circles.
    Works on carpet, cloth, any background.
    """
    try:
        import cv2
    except ImportError:
        return None

    import numpy as np
    w, h = img_pil.size
    img_arr = np.array(img_pil.convert("RGB"))
    gray = cv2.cvtColor(img_arr, cv2.COLOR_RGB2GRAY)

    # Stronger blur to suppress carpet/cloth texture
    blur = cv2.GaussianBlur(gray, (15, 15), 3)

    # Expected head radius: 2.5-13% of shorter image dimension
    min_side = min(w, h)
    min_r = int(min_side * 0.025)
    max_r = int(min_side * 0.13)

    # Auto-tune param2: start strict, relax until we find at least some circles
    # param2 = accumulator threshold — lower finds more circles
    best_circles = None
    for param2 in [70, 55, 45, 35, 28, 22]:
        circles = cv2.HoughCircles(
            blur,
            cv2.HOUGH_GRADIENT,
            dp=1.0,
            minDist=min_r * 1.8,
            param1=60,
            param2=param2,
            minRadius=min_r,
            maxRadius=max_r
        )
        if circles is not None:
            count = len(circles[0])
            # Stop when we find a plausible number (3-200 heads)
            if 3 <= count <= 200:
                best_circles = circles
                print(f"   Circle detection: param2={param2} → {count} circles")
                break
            elif count > 200:
                # Too many — too sensitive, keep trying stricter
                continue

    if best_circles is None:
        print(f"   Circle detection: no circles found")
        return []

    circles = np.round(best_circles[0]).astype(int)
    boxes = []
    for (cx, cy, r) in circles:
        pad = int(r * 0.40)
        x1 = max(0, cx - r - pad)
        y1 = max(0, cy - r - pad)
        x2 = min(w, cx + r + pad)
        y2 = min(h, cy + r + pad)
        boxes.append((x1, y1, x2, y2))

    return boxes


def detect_parts_yolo(img_pil, conf=0.01):
    """
    Use YOLOv8 to detect objects.
    yolo11n.pt is a general model — we use it at very low confidence
    to catch any object, then size-filter to keep only head-sized regions.
    """
    try:
        from ultralytics import YOLO
    except ImportError:
        return None

    import tempfile, os
    model = YOLO("yolo11n.pt")
    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
        tmp_path = tmp.name
    img_pil.save(tmp_path, "JPEG", quality=95)
    try:
        results = model(tmp_path, conf=conf, verbose=False)
    finally:
        os.unlink(tmp_path)

    w, h = img_pil.size
    # Expected head size: roughly 2-15% of image shorter dimension
    min_side = min(w, h)
    min_box = int(min_side * 0.02)
    max_box = int(min_side * 0.35)

    boxes = []
    for result in results:
        for box in result.boxes:
            x1, y1, x2, y2 = map(int, box.xyxy[0].tolist())
            bw, bh = x2 - x1, y2 - y1
            # Filter out boxes that are too small or too large
            if bw < min_box or bh < min_box:
                continue
            if bw > max_box or bh > max_box:
                continue
            pad = max(15, int(max(bw, bh) * 0.25))
            boxes.append((
                max(0, x1 - pad), max(0, y1 - pad),
                min(w, x2 + pad), min(h, y2 + pad)
            ))
    return boxes


def detect_parts_stud_based(img_pil, bg_color=None, padding_pct=10):
    """
    LEGO stud-based detection pipeline.

    Key insight: studs exist on every LEGO brick at fixed 8mm spacing.
    Even when two same-color bricks touch and edges disappear,
    their studs remain visible and spatially separated.

    Pipeline:
    1. CLAHE + bilateral filter — normalize lighting, preserve edges
    2. Hough circle detection — find studs
    3. DBSCAN clustering — group studs into bricks
    4. Bounding box from stud cluster extent + margin
    5. Returns None if not enough studs → caller falls back to blob detection
    """
    import numpy as np
    import cv2

    arr = np.array(img_pil.convert("RGB"))
    img_bgr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    h_img, w_img = arr.shape[:2]

    # Step 1: Normalize lighting — CLAHE on L channel + bilateral filter
    lab = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
    l2 = clahe.apply(l)
    img_norm = cv2.cvtColor(cv2.merge((l2, a, b)), cv2.COLOR_LAB2BGR)
    img_norm = cv2.bilateralFilter(img_norm, 9, 75, 75)
    gray = cv2.cvtColor(img_norm, cv2.COLOR_BGR2GRAY)

    # Step 2: Hough circle detection — try range of stud sizes
    all_circles = []
    for dp in [1.0, 1.2]:
        for (minR, maxR) in [(4, 8), (6, 12), (8, 16)]:
            circles = cv2.HoughCircles(
                gray, cv2.HOUGH_GRADIENT, dp=dp,
                minDist=int(minR * 1.5),
                param1=40, param2=15,
                minRadius=minR, maxRadius=maxR
            )
            if circles is not None:
                for c in circles[0]:
                    all_circles.append((float(c[0]), float(c[1]), float(c[2])))

    # Deduplicate circles within 5px
    deduped = []
    for c in all_circles:
        if not any((c[0]-d[0])**2+(c[1]-d[1])**2 < 25 for d in deduped):
            deduped.append(c)

    # Filter: keep only studs on foreground
    if bg_color is not None:
        br, bg_c, bb_c = bg_color
        studs = []
        for (cx, cy, r) in deduped:
            ix, iy = int(cx), int(cy)
            if 0 <= iy < h_img and 0 <= ix < w_img:
                px = arr[iy, ix].astype(int)
                d = np.sqrt(((px - np.array([br, bg_c, bb_c]))**2).sum())
                if d > 25:
                    studs.append((cx, cy, r))
    else:
        studs = deduped

    print(f"   🔵  Studs detected: {len(studs)}")
    if len(studs) < 2:
        return None

    # Step 3: Cluster studs — pure numpy, no external dependencies
    pts = np.array([(s[0], s[1]) for s in studs])
    # Nearest-neighbour distances using broadcasting
    diff = pts[:,None,:] - pts[None,:,:]
    dists = np.sqrt((diff**2).sum(axis=2))
    np.fill_diagonal(dists, np.inf)
    stud_spacing = float(np.median(dists.min(axis=1)))
    stud_spacing = max(12, min(stud_spacing, 60))
    print(f"   📏  Stud spacing: {stud_spacing:.1f}px")

    # Pure numpy union-find clustering — no sklearn/scipy needed
    def _cluster(pts, eps):
        n = len(pts)
        parent = list(range(n))
        def find(x):
            while parent[x] != x: parent[x] = parent[parent[x]]; x = parent[x]
            return x
        def union(a, b):
            a, b = find(a), find(b)
            if a != b: parent[b] = a
        for i in range(n):
            for j in range(i+1, n):
                dx = pts[i,0]-pts[j,0]; dy = pts[i,1]-pts[j,1]
                if dx*dx + dy*dy <= eps*eps: union(i, j)
        from collections import defaultdict
        groups = defaultdict(list)
        for i in range(n): groups[find(i)].append(i)
        return groups

    # Use wider epsilon so studs on same part cluster together
    eps = stud_spacing * 2.5
    groups = _cluster(pts, eps)

    # Step 4: Bounding box per cluster
    stud_boxes = []
    for idxs in groups.values():
        cluster_pts = pts[idxs]
        r_avg = float(np.mean([studs[i][2] for i in idxs]))
        n_studs = len(idxs)
        margin = stud_spacing * 1.2 + r_avg
        x1 = int(max(0, cluster_pts[:,0].min() - margin))
        y1 = int(max(0, cluster_pts[:,1].min() - margin))
        x2 = int(min(w_img, cluster_pts[:,0].max() + margin))
        y2 = int(min(h_img, cluster_pts[:,1].max() + margin))
        cx, cy = (x1+x2)//2, (y1+y2)//2
        min_half = int(stud_spacing * 1.0)
        x1 = max(0, min(x1, cx - min_half))
        y1 = max(0, min(y1, cy - min_half))
        x2 = min(w_img, max(x2, cx + min_half))
        y2 = min(h_img, max(y2, cy + min_half))
        if (x2-x1) > 8 and (y2-y1) > 8:
            stud_boxes.append((x1, y1, x2, y2))

    # Merge boxes that are very close — catches split-part problem
    # where two stud clusters on the same part are just barely outside eps
    def _merge_nearby(boxes, gap_frac=0.5):
        """Merge any two boxes whose gap is less than gap_frac * stud_spacing."""
        changed = True
        while changed:
            changed = False
            merged = []
            used = [False] * len(boxes)
            for i, a in enumerate(boxes):
                if used[i]: continue
                ax1,ay1,ax2,ay2 = a
                for j, b in enumerate(boxes):
                    if i == j or used[j]: continue
                    bx1,by1,bx2,by2 = b
                    gap_x = max(0, max(ax1,bx1) - min(ax2,bx2))
                    gap_y = max(0, max(ay1,by1) - min(ay2,by2))
                    if gap_x < stud_spacing * gap_frac and gap_y < stud_spacing * gap_frac:
                        ax1=min(ax1,bx1); ay1=min(ay1,by1)
                        ax2=max(ax2,bx2); ay2=max(ay2,by2)
                        used[j] = True; changed = True
                merged.append((ax1,ay1,ax2,ay2))
                used[i] = True
            boxes = merged
        return boxes

    stud_boxes = _merge_nearby(stud_boxes, gap_frac=0.6)

    # Remove tiny boxes — tip/shadow studs produce very small boxes
    min_box = stud_spacing * 1.5
    stud_boxes = [(x1,y1,x2,y2) for x1,y1,x2,y2 in stud_boxes
                  if (x2-x1) >= min_box and (y2-y1) >= min_box]

    print(f"   🧱  Stud clusters → {len(stud_boxes)} brick regions")
    return stud_boxes if stud_boxes else None


def detect_parts_geometric(img_pil, bg_color=None, padding_pct=10):
    """
    Geometric detection using OpenCV contours — finds solid foreground blobs
    with morphological cleanup. Better than Hough lines for mixed part types.
    Works well for flat plates, tiles, technic beams.
    """
    import numpy as np
    import cv2

    arr = np.array(img_pil.convert("RGB"))
    h_img, w_img = arr.shape[:2]

    # Step 1: Build foreground mask
    if bg_color is not None:
        br, bg_c, bb = bg_color
        d = np.sqrt(((arr.astype(int) - [br, bg_c, bb])**2).sum(axis=2))
        fg = (d > 35).astype(np.uint8) * 255
    else:
        gray0 = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
        blurred = cv2.GaussianBlur(gray0, (5, 5), 0)
        p95 = int(np.percentile(blurred, 95))
        p05 = int(np.percentile(blurred, 5))
        if p95 > 150:  # light background
            thresh = max(80, int(p95 * 0.82))
            fg = (blurred < thresh).astype(np.uint8) * 255
        else:  # dark background
            thresh = min(220, int(p05 + (p95-p05)*0.4) + 20)
            fg = (blurred > thresh).astype(np.uint8) * 255

    # Step 2: Morphological cleanup — close gaps, remove noise
    k3  = np.ones((3, 3), np.uint8)
    k7  = np.ones((7, 7), np.uint8)
    k15 = np.ones((15, 15), np.uint8)
    fg = cv2.morphologyEx(fg, cv2.MORPH_CLOSE, k7,  iterations=2)  # fill gaps
    fg = cv2.morphologyEx(fg, cv2.MORPH_OPEN,  k3,  iterations=1)  # remove noise
    fg = cv2.dilate(fg, k3, iterations=1)                            # slight expand

    # Step 3: Find contours of foreground blobs
    contours, _ = cv2.findContours(fg, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    min_area = max(200, (min(w_img, h_img) * 0.015) ** 2)
    boxes = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue
        x, y, w, h = cv2.boundingRect(cnt)
        # Skip if it's nearly the full image (false full-frame detection)
        if w > w_img * 0.85 and h > h_img * 0.85:
            continue
        boxes.append((x, y, x+w, y+h))

    if not boxes:
        return None

    # Step 4: NMS — remove boxes mostly inside a larger one
    boxes.sort(key=lambda b: (b[2]-b[0])*(b[3]-b[1]), reverse=True)
    kept = []
    for box in boxes:
        bx1,by1,bx2,by2 = box
        b_area = max(1, (bx2-bx1)*(by2-by1))
        dominated = False
        for kx1,ky1,kx2,ky2 in kept:
            ix1=max(bx1,kx1); iy1=max(by1,ky1)
            ix2=min(bx2,kx2); iy2=min(by2,ky2)
            inter=max(0,ix2-ix1)*max(0,iy2-iy1)
            if inter/b_area > 0.6:
                dominated = True; break
        if not dominated:
            kept.append(box)

    print(f"   📐  Geometric: {len(contours)} contours → {len(kept)} parts")
    return kept if kept else None


def detect_parts_flood(img_pil, bg_color, shadow_color=None, padding_pct=40):
    """
    Inverse flood-fill detection.

    Theory (Maxime): start from the border (guaranteed background).
    Walk outward through pixels close to bg color OR shadow color.
    Everything the flood CANNOT reach = a part.

    Shadow pixels act as separators: they are passable so the flood
    walks through them, cutting two adjacent parts apart.
    """
    import numpy as np
    import cv2

    arr   = np.array(img_pil.convert("RGB")).astype(np.uint8)
    h_img, w_img = arr.shape[:2]
    w, h  = img_pil.size

    # ── Convert to Lab for perceptual color distance ─────────────────────────
    bgr = arr[:, :, ::-1]
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB).astype(np.float32)

    def to_lab(rgb):
        px = np.array([[[rgb[0], rgb[1], rgb[2]]]], dtype=np.uint8)
        return cv2.cvtColor(px[:, :, ::-1], cv2.COLOR_BGR2LAB)[0, 0].astype(np.float32)

    lab_bg   = to_lab(bg_color)
    dist_bg  = np.sqrt(((lab - lab_bg) ** 2).sum(axis=2))

    # ── Adaptive tolerance from actual bg pixels on the image border ─────────
    edge_band = 40
    border = np.zeros((h_img, w_img), dtype=bool)
    border[:edge_band, :] = True; border[-edge_band:, :] = True
    border[:, :edge_band] = True; border[:, -edge_band:] = True

    border_dists = dist_bg[border]
    close = border_dists[border_dists < 25]
    if len(close) > 100:
        bg_tol = float(np.percentile(close, 85)) * 1.6
        bg_tol = float(np.clip(bg_tol, 12, 40))
    else:
        bg_tol = 28.0

    # ── Shadow tolerance (independent from bg) ───────────────────────────────
    sh_tol = 0.0
    dist_sh = None
    if shadow_color is not None:
        lab_sh  = to_lab(shadow_color)
        dist_sh = np.sqrt(((lab - lab_sh) ** 2).sum(axis=2))
        sh_close = dist_sh[dist_sh < 25]
        if len(sh_close) > 100:
            sh_tol = float(np.percentile(sh_close, 85)) * 1.6
            sh_tol = float(np.clip(sh_tol, 10, 35))
        else:
            sh_tol = 22.0

    print(f"   Flood: bg_tol={bg_tol:.1f}" + (f" sh_tol={sh_tol:.1f}" if shadow_color else ""))

    # ── Build passable mask ───────────────────────────────────────────────────
    # Passable = close to bg OR close to shadow
    # The flood walks through passable pixels from the border inward.
    # Shadow pixels in the middle of the image are passable → they cut
    # between two adjacent parts even when not reachable from the border,
    # because we ALSO erase them unconditionally after the fill.
    passable = (dist_bg <= bg_tol).astype(np.uint8)
    if dist_sh is not None:
        passable[dist_sh <= sh_tol] = 1

    # ── Connected components on passable — keep only border-touching ones ────
    # Label all passable regions
    n_labels, labels = cv2.connectedComponents(passable, connectivity=8)

    # Find labels that touch the image border
    border_labels = set()
    border_labels.update(map(int, labels[0, :]))
    border_labels.update(map(int, labels[-1, :]))
    border_labels.update(map(int, labels[:, 0]))
    border_labels.update(map(int, labels[:, -1]))
    border_labels.discard(0)  # 0 = non-passable

    # Background mask = passable AND reachable from border
    bg_mask = np.isin(labels, list(border_labels)).astype(np.uint8)

    # Parts mask = NOT background
    parts_mask = (1 - bg_mask).astype(np.uint8)

    # ── Shadow separator: erase shadow pixels regardless of connectivity ──────
    # This handles shadows surrounded by parts that the flood couldn't reach
    if dist_sh is not None:
        parts_mask[dist_sh <= sh_tol] = 0

    print(f"   Flood: {n_labels} regions, {len(border_labels)} bg → {int(parts_mask.sum()):,} fg px")

    # ── Morphological cleanup ─────────────────────────────────────────────────
    k3 = np.ones((3, 3), np.uint8)
    k5 = np.ones((5, 5), np.uint8)
    parts_mask = cv2.morphologyEx(parts_mask, cv2.MORPH_CLOSE, k5, iterations=2)
    parts_mask = cv2.morphologyEx(parts_mask, cv2.MORPH_OPEN,  k3, iterations=1)

    # ── Label blobs → bounding boxes ─────────────────────────────────────────
    try:
        from scipy import ndimage
        labeled, _ = ndimage.label(parts_mask)
        min_part_px = max(80, int((min(w, h) * 0.008) ** 2))
        min_dim_px  = max(6,  int(min(w, h) * 0.008))
        raw_boxes = []
        for sl in ndimage.find_objects(labeled):
            if sl is None: continue
            ys, xs = sl
            if int(parts_mask[ys, xs].sum()) < min_part_px: continue
            bw2 = xs.stop - xs.start; bh2 = ys.stop - ys.start
            if bw2 < min_dim_px or bh2 < min_dim_px: continue
            raw_boxes.append((xs.start, ys.start, xs.stop, ys.stop))
    except ImportError:
        raw_boxes = detect_parts_projection(parts_mask, w, h)

    if not raw_boxes:
        return []

    merge_gap, _ = auto_gap(raw_boxes)
    boxes = merge_boxes(raw_boxes, gap=merge_gap)

    _pad_frac = padding_pct / 100.0
    padded = []
    for (x1, y1, x2, y2) in boxes:
        bw2, bh2 = x2-x1, y2-y1
        pad_x = max(5, int(bw2 * _pad_frac))
        pad_y = max(5, int(bh2 * _pad_frac))
        padded.append((max(0,x1-pad_x), max(0,y1-pad_y),
                       min(w,x2+pad_x), min(h,y2+pad_y)))

    print(f"   Flood: {len(raw_boxes)} blobs → {len(padded)} parts")
    return sort_boxes_reading_order(padded)



def detect_parts(img_pil, debug=False, debug_dir=None, threshold=200, gap=None, hue_mode=False, use_fixed_grid=False, bg_color=None, shadow_color=None, padding_pct=40, use_watershed=False):
    """
    Detect part bounding boxes — restored to clean good version.
    shadow_color: optional (r,g,b) to suppress shadows/edges from binary mask.
    """
    import numpy as np
    w, h = img_pil.size

    if use_fixed_grid:
        grid_boxes = detect_parts_fixed_grid(img_pil, cols=_grid_cols, rows=_grid_rows)
        if grid_boxes:
            return grid_boxes

    gray = img_pil.convert("L")
    pixels = np.array(gray)
    h_img, w_img = pixels.shape

    from PIL import ImageFilter
    blur_radius = max(2, min(w, h) // 150)
    gray_blurred = gray.filter(ImageFilter.GaussianBlur(radius=blur_radius))
    pixels_blurred = np.array(gray_blurred)

    # ── Lab-space flood-fill from outside (best quality — both colors known) ──
    if bg_color is not None and shadow_color is not None:
        import cv2
        arr_rgb = np.array(img_pil.convert("RGB")).astype(np.uint8)

        def _color_to_lab(rgb):
            px = np.array([[[rgb[0], rgb[1], rgb[2]]]], dtype=np.uint8)
            return cv2.cvtColor(px[:,:,::-1], cv2.COLOR_BGR2LAB)[0,0].astype(np.float32)

        lab_arr = cv2.cvtColor(arr_rgb[:,:,::-1], cv2.COLOR_BGR2LAB).astype(np.float32)
        lab_bg  = _color_to_lab(bg_color)
        lab_sh  = _color_to_lab(shadow_color)

        dist_bg = np.sqrt(((lab_arr - lab_bg) ** 2).sum(axis=2))
        dist_sh = np.sqrt(((lab_arr - lab_sh) ** 2).sum(axis=2))

        # Auto-calibrate tolerance from actual image pixels near each known color
        def _tol(dist, seed_thresh=25, lo=15, hi=55):
            seed = dist < seed_thresh
            if seed.sum() > 200:
                return float(np.clip(np.percentile(dist[seed], 95) * 1.4, lo, hi))
            return 28.0

        bg_tol = _tol(dist_bg)
        sh_tol = _tol(dist_sh)

        # Initial mask: pixels matching bg or shadow color
        is_known = (dist_bg <= bg_tol) | (dist_sh <= sh_tol)

        # Flood fill from image border outward — ensures full connectivity.
        # Any pixel reachable from the border through known-color pixels = background.
        # Unreachable pixels = parts (completely surrounded by non-background).
        filled = is_known.copy().astype(np.uint8)
        # Use OpenCV floodFill on the INVERSE — fill from each border pixel
        # that is known background, spreading through connected known pixels
        # We do this by labeling connected components and keeping only those
        # touching the border
        from scipy import ndimage as _ndi
        labeled, _ = _ndi.label(filled)
        # Find labels touching any border
        border_labels = set()
        border_labels.update(labeled[0, :].tolist())
        border_labels.update(labeled[-1, :].tolist())
        border_labels.update(labeled[:, 0].tolist())
        border_labels.update(labeled[:, -1].tolist())
        border_labels.discard(0)  # 0 = already foreground

        # All border-connected background regions = definite background
        bg_mask = np.isin(labeled, list(border_labels))

        # Also include interior known-color pixels (shadows between parts)
        bg_mask |= is_known

        binary = (~bg_mask).astype(np.uint8)

        # Small morphological cleanup — remove 1-2px noise at part edges
        k = np.ones((3,3), np.uint8)
        binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, k, iterations=1)

        print(f"   Lab flood-fill: bg_tol={bg_tol:.1f} sh_tol={sh_tol:.1f} "
              f"→ {binary.sum():,} foreground px / {bg_mask.sum():,} background px")

    elif bg_color is not None:
        # Single bg color — adaptive grayscale (original good method)
        bg_gray = int(0.299*bg_color[0] + 0.587*bg_color[1] + 0.114*bg_color[2])
        seed_mask = (pixels_blurred >= max(0, bg_gray - 50)) &                     (pixels_blurred <= min(255, bg_gray + 50))
        if seed_mask.sum() > 200:
            bg_pixels = pixels_blurred[seed_mask].astype(float)
            bg_std  = float(np.std(bg_pixels))
            bg_mean = float(np.mean(bg_pixels))
            variance = int(np.clip(bg_std * 2.5, 25, 75))
            bg_lo = max(0,   int(bg_mean) - variance)
            bg_hi = min(255, int(bg_mean) + variance)
            print(f"   Background: USER rgb{bg_color} → gray≈{bg_gray}, "
                  f"measured mean={bg_mean:.0f} std={bg_std:.1f} → ±{variance} → [{bg_lo},{bg_hi}]")
        else:
            variance = 40
            bg_lo = max(0,   bg_gray - variance)
            bg_hi = min(255, bg_gray + variance)
            print(f"   Background: USER rgb{bg_color} → gray≈{bg_gray} ±{variance} (fallback) → [{bg_lo},{bg_hi}]")
        binary = ((pixels_blurred < bg_lo) | (pixels_blurred > bg_hi)).astype(np.uint8)

        # Shadow suppression on top
        if shadow_color is not None:
            sh_gray = int(0.299*shadow_color[0] + 0.587*shadow_color[1] + 0.114*shadow_color[2])
            seed = (pixels_blurred >= max(0, sh_gray-50)) & (pixels_blurred <= min(255, sh_gray+50))
            sh_mean = float(np.mean(pixels_blurred[seed].astype(float))) if seed.sum() > 100 else float(sh_gray)
            sh_var  = int(np.clip(float(np.std(pixels_blurred[seed].astype(float))) * 2.5, 20, 60)) if seed.sum() > 100 else 30
            binary[(pixels_blurred >= max(0, int(sh_mean)-sh_var)) &
                   (pixels_blurred <= min(255, int(sh_mean)+sh_var))] = 0
            print(f"   Shadow suppression: rgb{shadow_color} → gray≈{sh_gray}")

    else:
        # ── Corner-based background detection ────────────────────────────────
        # Sample four corners — guaranteed background, no parts placed there.
        import cv2 as _cv2
        band = max(20, min(w_img, h_img) // 20)  # ~5% of smaller dimension

        # Grayscale corners for threshold detection
        corners_gray = [
            pixels_blurred[:band,   :band  ],
            pixels_blurred[:band,   -band: ],
            pixels_blurred[-band:,  :band  ],
            pixels_blurred[-band:,  -band: ],
        ]
        corner_pixels = np.concatenate([c.ravel() for c in corners_gray])
        bg_mean = float(np.mean(corner_pixels))
        bg_std  = float(np.std(corner_pixels))
        variance = int(np.clip(bg_std * 2.5, 15, 80))
        bg_lo = max(0,   int(bg_mean) - variance)
        bg_hi = min(255, int(bg_mean) + variance)

        if bg_mean > 160:
            mat_type = "LIGHT"
        elif bg_mean < 80:
            mat_type = "DARK"
        else:
            mat_type = "MID"

        if mat_type == "DARK":
            # ── Dark mat: use LAB color distance from corner color ────────────
            # Grayscale threshold misses colorful parts on dark backgrounds
            # because dark blue/red/green are close in luminance to black.
            # Use full color (Lab) distance from the sampled corner color instead.
            arr_rgb = np.array(img_pil.convert("RGB")).astype(np.uint8)
            # Sample corner color in RGB
            corner_rgbs = np.concatenate([
                arr_rgb[:band,   :band  ].reshape(-1,3),
                arr_rgb[:band,   -band: ].reshape(-1,3),
                arr_rgb[-band:,  :band  ].reshape(-1,3),
                arr_rgb[-band:,  -band: ].reshape(-1,3),
            ])
            bg_rgb_mean = corner_rgbs.mean(axis=0)
            bg_px = np.array([[[int(bg_rgb_mean[0]), int(bg_rgb_mean[1]), int(bg_rgb_mean[2])]]], dtype=np.uint8)
            bg_lab = _cv2.cvtColor(bg_px[:,:,::-1], _cv2.COLOR_BGR2LAB)[0,0].astype(np.float32)

            lab_arr = _cv2.cvtColor(arr_rgb[:,:,::-1], _cv2.COLOR_BGR2LAB).astype(np.float32)
            dist    = np.sqrt(((lab_arr - bg_lab)**2).sum(axis=2))

            # Tolerance: measure actual spread of corner pixels in Lab
            corner_lab = np.concatenate([
                lab_arr[:band,   :band  ].reshape(-1,3),
                lab_arr[:band,   -band: ].reshape(-1,3),
                lab_arr[-band:,  :band  ].reshape(-1,3),
                lab_arr[-band:,  -band: ].reshape(-1,3),
            ])
            corner_dists = np.sqrt(((corner_lab - bg_lab)**2).sum(axis=1))
            lab_tol = float(np.clip(np.percentile(corner_dists, 90) * 2.5, 15, 45))

            binary = (dist > lab_tol).astype(np.uint8)
            print(f"   Background: AUTO-DARK Lab color distance tol={lab_tol:.1f} bg=rgb({int(bg_rgb_mean[0])},{int(bg_rgb_mean[1])},{int(bg_rgb_mean[2])})")
        else:
            binary = ((pixels_blurred < bg_lo) | (pixels_blurred > bg_hi)).astype(np.uint8)
            print(f"   Background: AUTO-{mat_type} corners mean={bg_mean:.0f} std={bg_std:.1f} → ±{variance} → [{bg_lo},{bg_hi}]")

    try:
        from scipy import ndimage
        labeled, num_features = ndimage.label(binary)
        min_part_px = max(80, int((min(w, h) * 0.008) ** 2))
        min_dim_px  = max(6,  int(min(w, h) * 0.008))
        slices = ndimage.find_objects(labeled)
        raw_boxes = []
        for sl in slices:
            if sl is None: continue
            ys, xs = sl
            bh = ys.stop - ys.start
            bw = xs.stop - xs.start
            px_count = int(binary[ys, xs].sum())
            if px_count < min_part_px or bw < min_dim_px or bh < min_dim_px:
                continue
            raw_boxes.append((xs.start, ys.start, xs.stop, ys.stop))
    except ImportError:
        raw_boxes = detect_parts_projection(binary, w, h)

    print(f"   Raw blobs before merge: {len(raw_boxes)}")
    if not raw_boxes:
        return []

    if gap is None:
        merge_gap, measured_spacing = auto_gap(raw_boxes)
        print(f"   Auto-gap: ~{measured_spacing:.0f}px → merge {merge_gap}px")
    else:
        merge_gap = gap

    boxes = merge_boxes(raw_boxes, gap=merge_gap)
    print(f"   After merge (gap={merge_gap}px): {len(boxes)} box(es)")

    def box_area(b): return max(1, (b[2]-b[0]) * (b[3]-b[1]))

    def split_oversize(boxes, gap_used):
        if len(boxes) < 2:
            return boxes
        areas = sorted(box_area(b) for b in boxes)
        median_area = areas[len(areas) // 2]
        max_ok = max(median_area * 2.5, w * h * 0.03)
        result = []
        for box in boxes:
            if box_area(box) <= max_ok:
                result.append(box); continue
            bx1, by1, bx2, by2 = box
            inside = [rb for rb in raw_boxes
                      if rb[0] < bx2 and rb[2] > bx1 and rb[1] < by2 and rb[3] > by1]
            if len(inside) <= 1:
                result.append(box); continue
            best = [box]
            for divisor in [4, 8, 16, 32]:
                tighter = max(2, gap_used // divisor)
                sub = merge_boxes(inside, gap=tighter)
                sub_oversized = [b for b in sub if box_area(b) > max_ok]
                if len(sub) > len(best) and not sub_oversized:
                    best = sub; break
                elif len(sub) > 1:
                    best = sub
            result.extend(best)
        return result

    boxes = split_oversize(boxes, merge_gap)

    if len(raw_boxes) >= 4 and len(boxes) == 1:
        print(f"   ⚠  Only 1 box from {len(raw_boxes)} raw blobs — parts may be touching.")

    padded = []
    _pad_frac = padding_pct / 100.0
    for (x1, y1, x2, y2) in boxes:
        bw, bh = x2 - x1, y2 - y1
        pad_x = max(5, int(bw * _pad_frac))
        pad_y = max(5, int(bh * _pad_frac))
        padded.append((
            max(0, x1 - pad_x), max(0, y1 - pad_y),
            min(w, x2 + pad_x), min(h, y2 + pad_y)
        ))

    return sort_boxes_reading_order(padded)


def detect_parts_projection(binary, w, h):
    """Coarse detection using row/col projection when scipy unavailable."""
    boxes = []
    # Divide into a grid and check each cell
    cols, rows = 8, 5  # default 8×5 = 40
    cw, ch = w // cols, h // rows
    for r in range(rows):
        for c in range(cols):
            x1, y1 = c * cw, r * ch
            x2, y2 = x1 + cw, y1 + ch
            region = binary[y1:y2, x1:x2]
            if region.sum() > 200:
                boxes.append((x1, y1, x2, y2))
    return boxes


def sort_boxes_reading_order(boxes):
    """
    Sort bounding boxes in reading order: top-to-bottom, left-to-right.
    Row grouping uses adaptive tolerance = half the median vertical gap
    between parts, so staggered/irregular layouts still sort correctly.
    Parts only need to be visually spaced — no grid required.
    """
    if not boxes:
        return boxes

    centers = [(x1, y1, x2, y2, (y1 + y2) / 2, (x1 + x2) / 2)
               for x1, y1, x2, y2 in boxes]
    centers.sort(key=lambda c: c[4])  # sort by vertical center

    # Compute adaptive row threshold from actual vertical gaps between parts
    cy_vals = [c[4] for c in centers]
    gaps = [cy_vals[i+1] - cy_vals[i] for i in range(len(cy_vals)-1)]
    if gaps:
        # Use median gap * 0.6 as row threshold
        # Parts in the same row have small gaps; parts in different rows have large gaps
        median_gap = sorted(gaps)[len(gaps) // 2]
        row_thresh = max(median_gap * 0.6, 10)
    else:
        row_thresh = 30

    # Group into rows
    rows = []
    current_row = [centers[0]]
    for box in centers[1:]:
        if abs(box[4] - current_row[0][4]) <= row_thresh:
            current_row.append(box)
        else:
            rows.append(current_row)
            current_row = [box]
    rows.append(current_row)

    # Within each row, sort left to right
    sorted_boxes = []
    for row in rows:
        row.sort(key=lambda c: c[5])
        for x1, y1, x2, y2, cy, cx in row:
            sorted_boxes.append((x1, y1, x2, y2))

    return sorted_boxes


def merge_boxes(boxes, gap=15):
    """Merge bounding boxes within `gap` pixels using Union-Find — O(n·α(n))."""
    if not boxes:
        return []
    n = len(boxes)
    parent = list(range(n))
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    def union(a, b):
        a, b = find(a), find(b)
        if a != b: parent[b] = a
    for i in range(n):
        ax1, ay1, ax2, ay2 = boxes[i]
        for j in range(i + 1, n):
            bx1, by1, bx2, by2 = boxes[j]
            if (ax1 - gap <= bx2 and ax2 + gap >= bx1 and
                    ay1 - gap <= by2 and ay2 + gap >= by1):
                union(i, j)
    from collections import defaultdict
    groups = defaultdict(lambda: [float("inf"), float("inf"), 0, 0])
    for i, (x1, y1, x2, y2) in enumerate(boxes):
        g = groups[find(i)]
        g[0] = min(g[0], x1); g[1] = min(g[1], y1)
        g[2] = max(g[2], x2); g[3] = max(g[3], y2)
    return [(int(v[0]), int(v[1]), int(v[2]), int(v[3])) for v in groups.values()]


# ─── BrickLink API color lookup ───────────────────────────────────────────────
def load_bl_credentials():
    """Load BrickLink OAuth credentials from .env file."""
    creds = {}
    env_path = Path(".env")
    if not env_path.exists():
        return None
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            creds[k.strip()] = v.strip().strip('"').strip("'")
    # Support both naming conventions used in the project
    key_map = {
        "CONSUMER_KEY":    ["CONSUMER_KEY", "BL_CONSUMER_KEY", "BRICKLINK_CONSUMER_KEY"],
        "CONSUMER_SECRET": ["CONSUMER_SECRET", "BL_CONSUMER_SECRET", "BRICKLINK_CONSUMER_SECRET"],
        "TOKEN":           ["TOKEN", "TOKEN_VALUE", "ACCESS_TOKEN", "BL_TOKEN", "BRICKLINK_TOKEN", "BRICKLINK_ACCESS_TOKEN"],
        "TOKEN_SECRET":    ["TOKEN_SECRET", "BL_TOKEN_SECRET", "BRICKLINK_TOKEN_SECRET"],
    }
    result = {}
    for canonical, aliases in key_map.items():
        for alias in aliases:
            if alias in creds:
                result[canonical] = creds[alias]
                break
    if len(result) == 4:
        return result
    return None


_bl_colors_cache = {}  # part_id → list of color_ids

def _make_oauth_header(creds: dict, method: str, url: str) -> dict:
    """Generate OAuth1 Authorization header as a plain string dict."""
    from requests_oauthlib import OAuth1Session
    oauth = OAuth1Session(
        creds["CONSUMER_KEY"],
        client_secret=creds["CONSUMER_SECRET"],
        resource_owner_key=creds["TOKEN"],
        resource_owner_secret=creds["TOKEN_SECRET"],
    )
    req = oauth.prepare_request(requests.Request(method, url))
    # req.headers is a CaseInsensitiveDict with str keys — safe to pass to aiohttp
    return {str(k): str(v) for k, v in req.headers.items()}


# Colors that are rare/limited-production — penalized in full-table fallback
# to avoid spurious matches when BL known-colors list is incomplete.
# A rare color must be ΔE 12 better than the best common color to win.
_RARE_COLOR_IDS = frozenset([
    168, 169,        # Umber, Sienna (2024 skin tones)
    240, 241,        # Medium Brown, Medium Tan (2022 skin tones)
    167,             # Reddish Orange (2024)
    173,             # Ochre Yellow (2025)
    175,             # Warm Pink (2026)
    174,             # Blue Violet (2026)
    171, 236,        # Lemon, Neon Yellow
    165, 166,        # Neon Orange, Neon Green
    160, 172, 96,    # Fabuland Orange, Warm Yellowish Orange, Very Light Orange
    161,             # Dark Yellow
    248,             # Fabuland Lime
    242,             # Dark Olive Green
    247, 245, 246,   # Little Robots Blue, Lilac, Light Lilac
    97,              # Royal Blue
    227,             # Clikits Lavender
    94,              # Medium Dark Pink
    56,              # Rose Pink
    231,             # Dark Salmon
    239, 238, 237,   # Bionicle Silver/Gold/Copper
    255, 252, 253, 254,  # Pearl Brown/Red/Green/Blue
    83, 61,          # Pearl White, Pearl Light Gold
    235,             # Reddish Gold
    244,             # Pearl Black
])

# Color IDs for minifig flesh/skin tones — only IDs present in BRICKLINK_COLORS.
# Used to constrain candidate list when a printed body part falls back to base.
_FLESH_COLOR_IDS = frozenset([
    3,    # Yellow       — classic pre-2004 heads/torsos/legs
    88,   # Light Flesh  — modern standard skin (post-2004)
    86,   # Dark Flesh   — medium-dark skin (older sets)
    91,   # Nougat       — warm medium skin
    68,   # Dark Orange  — used for some darker skin variants
])

# Body-part prefixes (heads, torsos, legs) — printed variants use flesh-only candidates
_BODY_PART_PREFIXES = ("3626", "24204", "973", "59349",  # heads + torsos
                       "970", "971", "972",               # hips + legs
                       "981", "982", "983",               # arms + hands
                       "74261", "3901", "90541")


def _is_body_part(part_id: str) -> bool:
    pid = part_id.lower()
    return any(pid.startswith(p) for p in _BODY_PART_PREFIXES)


async def get_bl_known_colors_async(session, part_id: str, creds: dict) -> list:
    """
    Fetch BrickLink known colors for a part.
    For printed body parts falling back to base part, restricts candidates
    to flesh/skin tones — prevents printed decoration confusing pixel matcher.
    """
    if part_id in _bl_colors_cache:
        return _bl_colors_cache[part_id]

    if not creds:
        _bl_colors_cache[part_id] = []
        return []

    import re
    ids_to_try = [part_id]
    base = re.sub(r'pb[a-zA-Z0-9]+.*$', '', part_id)
    if base == part_id:
        base = re.sub(r'p[cx][a-zA-Z0-9]+.*$', '', part_id)
    is_printed = base != part_id and bool(base)
    if is_printed:
        ids_to_try.append(base)

    def fetch_sync(try_id):
        try:
            url = f"https://api.bricklink.com/api/store/v1/items/part/{try_id}/colors"
            resp = requests.get(url, auth=_get_oauth(creds), timeout=8)
            if resp.status_code == 200:
                data = resp.json()
                return [c["color_id"] for c in data.get("data", [])]
            else:
                print(f"   BL colors for {try_id}: HTTP {resp.status_code}")
        except Exception as e:
            print(f"   BL colors for {try_id}: ERROR {e}")
        return None

    loop = asyncio.get_event_loop()
    for i, try_id in enumerate(ids_to_try):
        color_ids = await loop.run_in_executor(None, fetch_sync, try_id)
        if color_ids is not None and len(color_ids) > 0:
            if i > 0 and is_printed and _is_body_part(part_id):
                flesh_only = [c for c in color_ids if c in _FLESH_COLOR_IDS]
                if flesh_only:
                    print(f"   BL colors for {part_id}: printed body part, "
                          f"{len(color_ids)} base colors → {len(flesh_only)} flesh candidates")
                    color_ids = flesh_only
                else:
                    color_ids = color_ids[:4]
                    print(f"   BL colors for {part_id}: no flesh colors, using first 4")
            else:
                print(f"   BL colors for {part_id} (via {try_id}): {color_ids}")
            _bl_colors_cache[part_id] = color_ids
            return color_ids

    print(f"   BL colors for {part_id}: no results found")
    _bl_colors_cache[part_id] = []
    return []


def sample_part_color_rgb(crop_img):
    """
    Sample the dominant plastic body color from a LEGO part crop.
    Wrapped in try/except — numpy heap issues on Windows return a safe grey fallback.
    """
    try:
        return _sample_part_color_rgb_inner(crop_img)
    except Exception as e:
        print(f"   ⚠  Color sampling failed: {e} — using grey fallback")
        return (128, 128, 128)


def debug_color_decision(part_id: str, crop_img, brightness_bias: int,
                         scan_color: dict, brick_color: dict, merged: dict):
    """
    Optional verbose debug hook for GUI: print one-line summary of color decision.
    Safe to call from anywhere; no side effects beyond stdout.
    """
    try:
        sampled = sample_part_color_rgb(crop_img)
    except Exception:
        sampled = (128, 128, 128)
    print(f"   🧪 Color debug {part_id}: sampled={sampled}  "
          f"scan={scan_color.get('color_name')}[{scan_color.get('color_id')}] "
          f"({scan_color.get('color_method')},{scan_color.get('color_conf')})  "
          f"brick={brick_color.get('color_name')}[{brick_color.get('color_id')}] "
          f"({brick_color.get('color_method')},{brick_color.get('color_conf')})  "
          f"→ final={merged.get('color_name')}[{merged.get('color_id')}] "
          f"src={merged.get('color_source')}")


def _sample_part_color_rgb_inner(crop_img):
    """
    NEW (2026): Core-pixel sampling.
    We build a foreground mask for the plastic, then take only the *core* pixels
    (far from edges) to avoid:
      - cast shadows around the part
      - edge highlights / specular glints
      - background bleed from padding
    This makes color detection much more stable against lighting and cropping.

    After masking, we use a histogram bucket mode to pick the dominant body color.
    Returns plain Python (R, G, B) ints.
    """
    # Ensure OpenCV is available in this scope (core-pixel sampler uses cv2 directly).
    import cv2
    # Keep a moderate working size for speed; mask logic is resolution-invariant.
    arr = np.array(crop_img.convert("RGB").resize((120, 120), Image.LANCZOS))
    h, w = arr.shape[:2]

    # --- Core-pixel mask -------------------------------------------------------
    # Background ≈ border pixels (crop padding usually contains paper/background).
    border = 6
    bp = np.concatenate([
        arr[:border, :, :].reshape(-1, 3),
        arr[-border:, :, :].reshape(-1, 3),
        arr[:, :border, :].reshape(-1, 3),
        arr[:, -border:, :].reshape(-1, 3),
    ], axis=0).astype(np.uint8)
    bg_rgb = tuple(int(x) for x in np.median(bp, axis=0))

    # Convert to Lab so distance behaves closer to human perception.
    # (OpenCV uses BGR, so reverse channels.)
    lab = cv2.cvtColor(arr[:, :, ::-1], cv2.COLOR_BGR2LAB).astype(np.int16)
    bg_lab = cv2.cvtColor(np.array([[bg_rgb[::-1]]], dtype=np.uint8), cv2.COLOR_BGR2LAB).astype(np.int16)[0, 0]
    dL = lab[:, :, 0] - int(bg_lab[0])
    da = lab[:, :, 1] - int(bg_lab[1])
    db = lab[:, :, 2] - int(bg_lab[2])
    dist = np.sqrt(dL * dL + da * da + db * db)

    # Foreground = sufficiently different from background.
    fg = (dist > 12.0).astype(np.uint8) * 255
    # Clean speckles and fill small holes.
    k = max(3, int(min(h, w) * 0.03) | 1)  # odd kernel size
    ker = np.ones((k, k), np.uint8)
    fg = cv2.morphologyEx(fg, cv2.MORPH_OPEN, ker, iterations=1)
    fg = cv2.morphologyEx(fg, cv2.MORPH_CLOSE, ker, iterations=1)

    # Keep the largest connected component (the part).
    nlab, labels, stats, _ = cv2.connectedComponentsWithStats(fg, connectivity=8)
    if nlab > 1:
        areas = stats[1:, cv2.CC_STAT_AREA]
        keep = 1 + int(np.argmax(areas))
        fg = (labels == keep).astype(np.uint8) * 255

    # Core region = pixels far from edges of the foreground mask.
    # Distance transform gives distance to nearest background (edge).
    dt = cv2.distanceTransform(fg, distanceType=cv2.DIST_L2, maskSize=5)
    dt_max = float(dt.max()) if dt.size else 0.0
    if dt_max > 0:
        core = dt >= (dt_max * 0.35)  # inner “heart” of the piece
    else:
        # Fallback: inner rectangle (old behavior) if mask failed
        core = np.zeros((h, w), dtype=bool)
        y0 = int(h * 0.25); y1 = int(h * 0.75)
        x0 = int(w * 0.25); x1 = int(w * 0.75)
        core[y0:y1, x0:x1] = True

    core_px = arr[core].reshape(-1, 3).astype(float)

    def _clean_pixels(region):
        px = region.reshape(-1, 3).astype(float)
        mx = px.max(axis=1); mn = px.min(axis=1)
        brightness = mx
        saturation = np.where(mx > 0, (mx - mn) / mx, 0)
        # Exclude white background (very bright, low saturation)
        not_white = ~((px[:,0] > 210) & (px[:,1] > 210) & (px[:,2] > 210))
        # Exclude pure cast-shadow pixels — threshold kept low (18) so genuinely
        # dark/black LEGO parts (which read ~20-60) are NOT discarded.
        not_black = ~((brightness < 18))
        # Exclude specular highlights (bright + low saturation)
        not_specular = ~((saturation < 0.10) & (brightness > 190))
        # IMPORTANT: do NOT discard low-saturation pixels unconditionally.
        # Grey/black LEGO parts are legitimately low-saturation, and removing them can bias sampling
        # toward edge noise or tiny colored artifacts.
        # We only drop "neutral + bright" pixels (usually background/paper or glare).
        not_neutral_bright = ~((saturation < 0.07) & (brightness > 160))
        return px[not_white & not_black & not_specular & not_neutral_bright]

    # Primary: clean only core pixels (best for shadow/edge rejection)
    best_pixels = _clean_pixels(core_px) if len(core_px) else np.empty((0, 3))

    # Fallback: if core is too small (thin parts), progressively widen to include more.
    if len(best_pixels) < 20:
        for frac in [0.60, 0.75, 0.90, 1.00]:
            y0 = int(h * (1 - frac) / 2); y1 = int(h * (1 + frac) / 2)
            x0 = int(w * (1 - frac) / 2); x1 = int(w * (1 + frac) / 2)
            px = _clean_pixels(arr[y0:y1, x0:x1])
            if len(px) >= 20:
                best_pixels = px
                break
            if len(px) > len(best_pixels):
                best_pixels = px

    if len(best_pixels) < 4:
        # Last resort: median of full image excluding white
        px = arr.reshape(-1, 3).astype(float)
        mask = ~((px[:,0] > 215) & (px[:,1] > 215) & (px[:,2] > 215))
        best_pixels = px[mask] if mask.sum() >= 4 else px

    # Use finer histogram (32 bins) for better color discrimination
    BINS = 32
    scale = 256.0 / BINS
    bins = np.floor(best_pixels / scale).astype(int).clip(0, BINS - 1)
    bucket_idx = bins[:,0] * BINS * BINS + bins[:,1] * BINS + bins[:,2]
    counts = np.bincount(bucket_idx, minlength=BINS**3)
    best_bucket = int(counts.argmax())
    br = best_bucket // (BINS * BINS)
    bg_b = (best_bucket // BINS) % BINS
    bb = best_bucket % BINS
    r = int((br + 0.5) * scale)
    g = int((bg_b + 0.5) * scale)
    b = int((bb + 0.5) * scale)
    return (r, g, b)


# Prefer Bluish Gray over plain Gray — plain grays are rare in modern LEGO
_GRAY_PENALTY = {9: 5.0, 10: 5.0}  # Light Gray, Dark Gray

# ── Color merge logic (Scanner core-pixels + Brickognize) ─────────────────────
# NEW (2026): The GUI can now display a merged color decision.
# Historically the project stored only one "color_id" field, so there was no way
# for the GUI to “switch to Brickognize color” after the fact — it had been
# overwritten by the scanner’s own sampling (or vice‑versa). We now preserve both.
#
# Hooks for future tuning:
COLOR_MERGE_ENABLE = True
COLOR_MERGE_SCAN_MIN_CONF   = 0.55  # below this, scanner color is considered weak
COLOR_MERGE_BRICK_MIN_CONF  = 0.60  # below this, Brickognize color is considered weak
COLOR_MERGE_PREFER_MARGIN   = 0.08  # Brickognize must beat scanner by this to override when both are “OK”

# color_mode:
#   "merge"       → smart merge (default, existing behaviour)
#   "brickognize" → always use Brickognize color when it has a BL id
#   "scan"        → always use scanner/core-pixel color
COLOR_MODE = "merge"

def _color_conf_from_method(method: str) -> float:
    """Map our scanner color_method to a rough confidence score [0..1]."""
    m = (method or "").lower()
    if m == "exact":
        return 1.0
    if m == "forced":
        return 1.0
    if m == "brickognize":
        return 0.85
    if m == "matched":
        return 0.80
    if "unreliable" in m:
        return 0.30
    if m == "guessed":
        return 0.55
    if m == "unknown":
        return 0.10
    return 0.40

def merge_color_decision(scan_color: dict, brick_color: dict, known_ids: list) -> dict:
    """
    Decide the final displayed color using both sources.
    Returns dict with keys: color_id, color_name, color_method, color_conf, color_source.
    """
    if not COLOR_MERGE_ENABLE:
        return {
            "color_id": scan_color.get("color_id", 0),
            "color_name": scan_color.get("color_name", "Unknown"),
            "color_method": scan_color.get("color_method", "matched"),
            "color_conf": float(scan_color.get("color_conf", 0.5)),
            "color_source": "scan",
        }

    sc_id = scan_color.get("color_id", 0) or 0
    bc_id = brick_color.get("color_id", 0) or 0
    sc_cf = float(scan_color.get("color_conf", 0.0) or 0.0)
    bc_cf = float(brick_color.get("color_conf", 0.0) or 0.0)

    # If Brickognize didn’t produce a usable BL color id, stick with scan.
    if not bc_id:
        return {**scan_color, "color_source": "scan"}

    # Hard modes first.
    if COLOR_MODE == "brickognize":
        return {**brick_color, "color_source": "brickognize"}
    if COLOR_MODE == "scan":
        return {**scan_color, "color_source": "scan"}

    # If scanner is weak/unreliable and Brickognize is reasonably confident, use Brickognize.
    if sc_cf < COLOR_MERGE_SCAN_MIN_CONF and bc_cf >= COLOR_MERGE_BRICK_MIN_CONF:
        return {**brick_color, "color_source": "brickognize"}

    # If Brickognize is weak but scanner is strong, keep scanner.
    if bc_cf < COLOR_MERGE_BRICK_MIN_CONF and sc_cf >= COLOR_MERGE_SCAN_MIN_CONF:
        return {**scan_color, "color_source": "scan"}

    # If both are “OK”, prefer Brickognize only if it clearly wins.
    if bc_cf >= sc_cf + COLOR_MERGE_PREFER_MARGIN:
        return {**brick_color, "color_source": "brickognize"}

    # Otherwise keep the scanner (core-pixel) result.
    return {**scan_color, "color_source": "scan"}

def resolve_color_from_cache(part_id: str, crop_img, brightness_bias: int = 0):
    """
    Resolve color for a part:
    - 1 known color  → use it directly (exact)
    - 2+ known colors → compare crop RGB to candidates, pick closest (matched)
    - 0 known colors / cache miss → pixel-sample against full BL color table (guessed)
    brightness_bias: shift sampled RGB brighter (+) or darker (−) before Lab matching.
    Returns (color_id, color_name, method)
    """
    known = _bl_colors_cache.get(part_id, None)

    if known is None:
        print(f"   ⚠  Color cache MISS for {part_id} — falling back to full table")
    elif len(known) == 1:
        # Only one color this part exists in — no need to look at image
        cid = known[0]
        cname = BRICKLINK_COLORS.get(cid, (str(cid), (128, 128, 128)))[0]
        return cid, cname, "exact"
    elif len(known) >= 2:
        # Multiple known colors — pick closest by perceptual Lab ΔE
        sampled_rgb = sample_part_color_rgb(crop_img)
        # Apply brightness bias — shift L channel in Lab space
        _ensure_lab_cache()
        _sl = _rgb_to_lab(sampled_rgb)
        if brightness_bias != 0:
            sampled_lab = (max(0.0, min(100.0, _sl[0] + brightness_bias * 0.6)), _sl[1], _sl[2])
        else:
            sampled_lab = _sl
        candidates = [(cid, BRICKLINK_COLORS[cid]) for cid in known if cid in BRICKLINK_COLORS]
        if candidates:
            best_id, best_name, best_dist = candidates[0][0], candidates[0][1][0], float("inf")
            for cid, (cname, ref_rgb) in candidates:
                ref_lab = _BL_COLORS_LAB.get(cid)
                if ref_lab is None:
                    ref_lab = _rgb_to_lab(ref_rgb)
                d = ((sampled_lab[0]-ref_lab[0])**2 +
                     (sampled_lab[1]-ref_lab[1])**2 +
                     (sampled_lab[2]-ref_lab[2])**2) ** 0.5
                d += _GRAY_PENALTY.get(cid, 0.0)  # penalize rare plain grays
                if d < best_dist:
                    best_dist = d
                    best_id = cid
                    best_name = cname
            print(f"   🎨  sampled {sampled_rgb} → {best_name} (ΔE {best_dist:.0f}, {len(candidates)} candidates)")
            # Luminance sanity check: if overall crop is very dark but we matched a light
            # color, prefer the darkest known candidate instead (handles black parts under
            # bright light where specular highlights skew the sample).
            # Dark-part sanity check.
            # IMPORTANT: don't let bright background/paper dominate this measurement.
            _arr_chk = np.array(crop_img.convert("RGB").resize((60, 60), Image.LANCZOS))
            _br = _arr_chk.max(axis=2).astype(np.float32)
            # ignore very bright pixels (likely background/specular)
            _mask = _br < 210
            _med_br  = float(np.median(_br[_mask])) if _mask.sum() > 50 else float(np.median(_br))
            # Also treat the sampled RGB itself as evidence of a dark part
            if _med_br < 70 or max(sampled_rgb) < 60:  # image is overall dark — part is likely dark
                _best_lab = _BL_COLORS_LAB.get(best_id, (100, 0, 0))
                if _best_lab[0] > 50:  # but we matched a light color — suspicious
                    dark_cands = [(cid, cname, ref_rgb)
                                  for cid, (cname, ref_rgb) in candidates
                                  if (_BL_COLORS_LAB.get(cid, (100,))[0]) < 40]
                    if dark_cands:
                        dc_id, dc_name, _ = min(
                            dark_cands,
                            key=lambda x: (
                                (sampled_lab[0] - _BL_COLORS_LAB.get(x[0], (0,0,0))[0])**2 +
                                (sampled_lab[1] - _BL_COLORS_LAB.get(x[0], (0,0,0))[1])**2 +
                                (sampled_lab[2] - _BL_COLORS_LAB.get(x[0], (0,0,0))[2])**2
                            ))
                        print(f"   🎨  Dark image override: {best_name}→{dc_name} "
                              f"(img brightness {_med_br:.0f})")
                        best_id, best_name = dc_id, dc_name
            # Body part with high uncertainty → default to Yellow (most common classic skin)
            if best_dist > 30 and _is_body_part(part_id) and len(candidates) > 2:
                if any(cid == 3 for cid, _ in candidates):
                    print(f"   🎨  ΔE {best_dist:.0f} too uncertain for body part — defaulting to Yellow")
                    return 3, "Yellow", "matched"
            # Known colors are authoritative — always stay within them, never fall back to full table
            method = "matched" if best_dist <= 35 else "unreliable"
            return best_id, best_name, method
        # All candidate IDs outside our color map — fall through to full table
        print(f"   ⚠  Known color IDs for {part_id} not in local table — falling back to full table")
    # known == [] (BL returned no colors) or candidates empty or cache miss
    # — pixel-sample against full BrickLink color table as best-effort guess
    sampled_rgb = sample_part_color_rgb(crop_img)
    _ensure_lab_cache()
    _sl = _rgb_to_lab(sampled_rgb)
    if brightness_bias != 0:
        sampled_lab = (max(0.0, min(100.0, _sl[0] + brightness_bias * 0.6)), _sl[1], _sl[2])
    else:
        sampled_lab = _sl

    # Luminance pre-filter: if the overall image is very dark, only consider
    # dark BL colors (L < 40). This prevents black parts with specular highlights
    # from being matched to white/light colors in the full-table search.
    _arr_full = np.array(crop_img.convert("RGB").resize((60, 60), Image.LANCZOS))
    _br_ch    = _arr_full.max(axis=2).astype(np.float32)
    # Ignore very bright pixels (likely specular/background) when measuring median
    _non_spec = _br_ch[_br_ch < 210]
    _median_brightness = float(np.median(_non_spec)) if len(_non_spec) > 50 else float(np.median(_br_ch))
    # Raise threshold to 80 — glossy black parts can have specular peaks raising median
    _dark_part = _median_brightness < 80 or max(sampled_rgb) < 70
    # Also: if the sampled RGB itself is very dark, restrict immediately regardless of median
    _sampled_very_dark = max(sampled_rgb) < 55  # near-black sample → force dark candidates
    _lum_filter = set()
    if _dark_part or _sampled_very_dark:
        _lum_filter = {cid for cid, lab in _BL_COLORS_LAB.items() if lab[0] < 42}
        print(f"   🎨  Dark image (med brightness {_median_brightness:.0f}, sample {sampled_rgb}) — restricting to {len(_lum_filter)} dark colors")
    # For near-black sampled colors: weight L heavily and de-weight chroma (a,b)
    # so that warm-light noise (slight reddish tint on black plastic) doesn't
    # push the match toward 'Dark Bluish Gray' or similar.
    _wL = 3.0 if (sampled_lab[0] < 30) else 1.0   # triple L weight for very dark samples
    _wC = 0.4 if (sampled_lab[0] < 30) else 1.0   # reduce chroma weight
    best_id, best_name, best_dist = 0, "Unknown", float("inf")
    for cid, (cname, _) in BRICKLINK_COLORS.items():
        if _lum_filter and cid not in _lum_filter:
            continue  # skip light colors when part is clearly dark
        ref_lab = _BL_COLORS_LAB.get(cid)
        if ref_lab is None: continue
        dL = sampled_lab[0]-ref_lab[0]
        da = sampled_lab[1]-ref_lab[1]
        db = sampled_lab[2]-ref_lab[2]
        d = (_wL*dL*dL + _wC*(da*da + db*db)) ** 0.5
        d += _GRAY_PENALTY.get(cid, 0.0)  # penalize rare plain grays
        if d < best_dist:
            best_dist = d
            best_id = cid
            best_name = cname
    if best_id == 0:
        return 0, "Unknown", "unknown"
    method = "guessed" if best_dist > 18 else "matched"
    print(f"   🎨  full-table guess {sampled_rgb} → {best_name} (ΔE {best_dist:.0f})")
    return best_id, best_name, method


_brickognize_last = 0.0  # timestamp of last request dispatched
_BRICKOGNIZE_INTERVAL = 0.26  # 4 req/sec max = 250ms between requests + margin

async def query_brickognize(session, crop_bytes, filename, semaphore, retries=3):
    """Send one cropped image to Brickognize, with retries on failure."""
    if not crop_bytes:
        print(f"  ⚠ No image data for {filename} — skipping")
        return None
    global _brickognize_last
    url = "https://api.brickognize.com/predict/"
    for attempt in range(1, retries + 1):
        async with semaphore:
            # Pace requests to stay under 4/sec
            now = asyncio.get_event_loop().time()
            wait = _BRICKOGNIZE_INTERVAL - (now - _brickognize_last)
            if wait > 0:
                await asyncio.sleep(wait)
            _brickognize_last = asyncio.get_event_loop().time()
            data = aiohttp.FormData()
            data.add_field(
                "query_image",
                crop_bytes,
                filename=filename,
                content_type="image/jpeg"
            )
            try:
                async with session.post(url, data=data, headers={"accept": "application/json"}, timeout=aiohttp.ClientTimeout(total=30)) as resp:
                    if resp.status == 200:
                        result = await resp.json()
                        if result and result.get("items"):
                            return result
                        # Empty response — retry
                        if attempt < retries:
                            await asyncio.sleep(1.5 * attempt)
                            continue
                        return result
                    elif resp.status == 429:
                        # Rate limited — wait longer
                        await asyncio.sleep(3 * attempt)
                        continue
                    else:
                        text = await resp.text()
                        print(f"  ⚠ API error {resp.status} for {filename}: {text[:80]}")
                        return None
            except asyncio.TimeoutError:
                if attempt < retries:
                    await asyncio.sleep(2 * attempt)
                    continue
                print(f"  ⚠ Timeout for {filename} after {retries} attempts")
                return None
            except Exception as e:
                if attempt < retries:
                    await asyncio.sleep(1.5 * attempt)
                    continue
                print(f"  ⚠ Failed for {filename}: {e}")
                return None
    return None


# Module-level OAuth1 instance cache — avoids re-init on every API call
_oauth_cache: dict = {}

def _get_oauth(creds: dict):
    key = (creds.get("CONSUMER_KEY",""), creds.get("TOKEN",""))
    if key not in _oauth_cache:
        _oauth_cache[key] = OAuth1(creds["CONSUMER_KEY"], creds["CONSUMER_SECRET"],
                                   creds["TOKEN"], creds["TOKEN_SECRET"])
    return _oauth_cache[key]


# Cache for mold variants: part_id → list of {id, name}
_mold_variants_cache = {}

def fetch_mold_variants(part_id: str, creds: dict) -> list:
    """
    Fetch mold variants for a part via BrickLink API GET /items/part/{id}.
    Returns list of dicts: [{id, name, type}] or []
    BrickLink surfaces these as "This item is a mold variant of..." on catalog pages.
    The API returns alternate_no in the item data.
    """
    pid = part_id.lower().strip()
    if pid in _mold_variants_cache:
        return _mold_variants_cache[pid]
    if not creds:
        return []
    try:
        url = f"https://api.bricklink.com/api/store/v1/items/part/{part_id}"
        resp = requests.get(url, auth=_get_oauth(creds), timeout=8)
        if resp.status_code == 200:
            data = resp.json().get("data", {})
            variants = []
            # alternate_no contains mold variant part numbers (comma-separated)
            alt_no = data.get("alternate_no", "")
            if alt_no:
                for alt in alt_no.replace(" ", "").split(","):
                    if alt and alt.lower() != pid:
                        # Fetch name for this variant
                        try:
                            r2 = requests.get(
                                f"https://api.bricklink.com/api/store/v1/items/part/{alt}",
                                auth=_get_oauth(creds), timeout=5)
                            if r2.status_code == 200:
                                d2 = r2.json().get("data", {})
                                variants.append({
                                    "id": alt,
                                    "name": d2.get("name", alt),
                                    "type": "P",
                                    "score": 0,
                                    "img_url": d2.get("thumbnail_url", ""),
                                    "is_mold_variant": True
                                })
                        except Exception:
                            variants.append({"id": alt, "name": alt, "type": "P",
                                             "score": 0, "img_url": "", "is_mold_variant": True})
            _mold_variants_cache[pid] = variants
            if variants:
                print(f"   🔀  Mold variants for {part_id}: {[v['id'] for v in variants]}")
            return variants
    except Exception:
        pass
    _mold_variants_cache[pid] = []
    return []


def fetch_medium_price(part_id, color_id, creds, condition="U", currency="CAD", item_type="P"):
    """Fetch medium sale price from BrickLink for a part+color (or minifig). Returns float or None."""
    if not creds:
        return None
    try:
        auth = _get_oauth(creds)
        if item_type == "M":
            url = f"https://api.bricklink.com/api/store/v1/items/minifig/{part_id}/price?guide_type=sold&new_or_used={condition}&currency_code={currency}"
        else:
            url = f"https://api.bricklink.com/api/store/v1/items/part/{part_id}/price?color_id={color_id}&guide_type=sold&new_or_used={condition}&currency_code={currency}"
        resp = requests.get(url, auth=auth, timeout=8)
        if resp.status_code == 200:
            data = resp.json().get("data", {})
            avg = data.get("avg_price") or data.get("qty_avg_price")
            if avg:
                return round(float(avg), 3)
        elif resp.status_code == 404 and item_type == "M":
            # Minifig not found — retry as part (Brickognize sometimes misclassifies)
            url2 = f"https://api.bricklink.com/api/store/v1/items/part/{part_id}/price?color_id={color_id}&guide_type=sold&new_or_used={condition}&currency_code={currency}"
            resp2 = requests.get(url2, auth=auth, timeout=8)
            if resp2.status_code == 200:
                data = resp2.json().get("data", {})
                avg = data.get("avg_price") or data.get("qty_avg_price")
                if avg:
                    return round(float(avg), 3)
    except Exception:
        pass
    return None


def build_xml(results, qty=1, bl_creds=None, currency="CAD", prices=None):
    """
    Build BrickLink XML inventory string.
    - Merges duplicate part+color into one entry with summed quantity
    - Fetches medium sold price from BrickLink if credentials available
    """
    # Merge duplicates: key = (part_id, color_id)
    merged = {}
    for r in results:
        if not r.get("part_id"):
            continue
        key = (r["part_id"], r["color_id"])
        if key in merged:
            merged[key]["qty"] += qty
        else:
            merged[key] = {**r, "qty": qty}

    # Use pre-fetched prices if provided, else fetch now
    if prices is None:
        if bl_creds:
            import concurrent.futures
            def fetch(entry):
                return entry, fetch_medium_price(
                    entry["part_id"], entry["color_id"], bl_creds,
                    currency=currency, item_type=entry.get("item_type", "P"))
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
                price_results = list(ex.map(fetch, merged.values()))
            prices = {(e["part_id"], e["color_id"]): p for e, p in price_results}
        else:
            prices = {}

    inventory = ET.Element("INVENTORY")
    for (part_id, color_id), entry in merged.items():
        item_type = entry.get("item_type", "P")
        item = ET.SubElement(inventory, "ITEM")
        ET.SubElement(item, "ITEMTYPE").text = "M" if item_type == "M" else "P"
        ET.SubElement(item, "ITEMID").text = str(part_id)
        if item_type != "M":
            ET.SubElement(item, "COLOR").text = str(color_id)
            # Note: COLOR_NOTE is not a valid BrickLink XML field — omitted
        ET.SubElement(item, "QTY").text = str(entry["qty"])
        ET.SubElement(item, "CONDITION").text = "U"
        price = prices.get((part_id, color_id))
        if price is not None and price > 0:
            ET.SubElement(item, "PRICE").text = f"{price:.4f}"

    xml_str = minidom.parseString(ET.tostring(inventory)).toprettyxml(indent="  ")
    # Strip XML declaration — BrickLink upload expects raw <INVENTORY> with no header noise
    lines = xml_str.split("\n")
    if lines and lines[0].startswith("<?xml"):
        lines = lines[1:]
    return "\n".join(lines)


def build_tsv(results, qty=1):
    """Build BrickLink tab-separated upload format. Merges duplicates."""
    merged = {}
    for r in results:
        if not r.get("part_id"):
            continue
        key = (r["part_id"], r["color_id"])
        if key in merged:
            merged[key]["qty"] += qty
        else:
            merged[key] = {**r, "qty": qty}

    lines = ["Part ID\tColor ID\tQuantity\tCondition\tPart Name\tColor Name\tConfidence"]
    for entry in merged.values():
        lines.append(
            f"{entry['part_id']}\t{entry['color_id']}\t{entry['qty']}\tU\t"
            f"{entry['part_name']}\t{entry['color_name']}\t{entry['confidence']:.0%}"
        )
    return "\n".join(lines)


def build_html_report(results, image_path, output_dir):
    """Build a visual HTML review page with all detections."""
    rows = ""
    for i, r in enumerate(results):
        status_color = "#2ecc71" if r.get("part_id") else "#e74c3c"
        status = "✓" if r.get("part_id") else "✗ skipped"
        confidence_bar = f'<div style="width:{r.get("confidence",0)*100:.0f}%;height:6px;background:#3498db;border-radius:3px"></div>'
        thumb_src = r.get("thumb_url", "")
        crop_src = r.get("crop_path", "")

        rows += f"""
        <tr>
          <td style="text-align:center;color:#888">{i+1}</td>
          <td><img src="{crop_src}" style="max-width:80px;max-height:80px;border-radius:4px"></td>
          <td><img src="{thumb_src}" style="max-width:80px;max-height:80px"></td>
          <td style="font-weight:bold">{r.get("part_id","—")}</td>
          <td>{r.get("part_name","—")}</td>
          <td>
            <span style="display:inline-block;width:16px;height:16px;background:rgb{r.get('color_rgb',(200,200,200))};border:1px solid #ccc;border-radius:3px;vertical-align:middle"></span>
            {r.get("color_name","—")} <small style="color:#888">({r.get("color_id","—")})</small>
          </td>
          <td>{confidence_bar} <small>{r.get("confidence",0):.0%}</small></td>
          <td style="color:{status_color};font-weight:bold">{status}</td>
        </tr>"""

    total = len(results)
    identified = sum(1 for r in results if r.get("part_id"))
    skipped = total - identified

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Sheet Scan Results</title>
<style>
  body {{ font-family: system-ui, sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }}
  h1 {{ color: #2c3e50; }}
  .stats {{ display:flex; gap:20px; margin-bottom:20px; }}
  .stat {{ background:#fff; border-radius:8px; padding:16px 24px; box-shadow:0 1px 4px rgba(0,0,0,.1); }}
  .stat-value {{ font-size:2em; font-weight:bold; color:#3498db; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,.1); }}
  th {{ background:#2c3e50; color:#fff; padding:10px 12px; text-align:left; font-size:.85em; text-transform:uppercase; }}
  td {{ padding:8px 12px; border-bottom:1px solid #eee; vertical-align:middle; }}
  tr:hover td {{ background:#f9f9f9; }}
  .dl-links {{ margin: 20px 0; display: flex; gap: 12px; }}
  .btn {{ padding: 10px 20px; border-radius: 6px; text-decoration: none; color: #fff; font-weight: bold; }}
  .btn-xml {{ background: #e67e22; }}
  .btn-tsv {{ background: #27ae60; }}
</style>
</head>
<body>
<h1>🧱 Sheet Scan Results</h1>
<p style="color:#888">Source: {image_path} — {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
<div class="stats">
  <div class="stat"><div class="stat-value">{total}</div><div>Parts detected</div></div>
  <div class="stat"><div class="stat-value">{identified}</div><div>Identified</div></div>
  <div class="stat"><div class="stat-value">{skipped}</div><div>Skipped / low confidence</div></div>
</div>
<div class="dl-links">
  <a class="btn btn-xml" href="scan-results.xml" download>⬇ Download XML (BrickStore)</a>
  <a class="btn btn-tsv" href="scan-results.tsv" download>⬇ Download TSV (BrickLink)</a>
</div>
<table>
<thead>
  <tr><th>#</th><th>Your Crop</th><th>BL Thumb</th><th>Part ID</th><th>Name</th><th>Color</th><th>Confidence</th><th>Status</th></tr>
</thead>
<tbody>
{rows}
</tbody>
</table>
</body>
</html>"""
    return html


def preview_detections(img_pil, boxes, output_path):
    """Draw numbered bounding boxes on the photo. No API calls — just shows what was detected."""
    from PIL import ImageDraw
    img = img_pil.copy().convert("RGB")
    draw = ImageDraw.Draw(img)
    palette = [(255,50,50),(50,200,50),(50,100,255),(255,180,0),(200,0,200),(0,200,200)]
    for i, (x1, y1, x2, y2) in enumerate(boxes):
        color = palette[i % len(palette)]
        for t in range(3):
            draw.rectangle([x1-t, y1-t, x2+t, y2+t], outline=color)
        label = str(i + 1)
        lx, ly = x1 + 4, y1 + 2
        draw.rectangle([lx-2, ly-1, lx + len(label)*8 + 2, ly+14], fill=color)
        draw.text((lx, ly), label, fill=(255,255,255))
    max_side = 1400
    w, h = img.size
    if max(w, h) > max_side:
        scale = max_side / max(w, h)
        img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
    img.save(output_path)
    print(f"\n   Preview: {output_path}  ({len(boxes)} parts detected)")
    print(f"   Check the image. If a part is split into multiple boxes: increase --gap")
    print(f"   If two parts are merged into one box: decrease --gap\n")
    try:
        import subprocess, platform
        cmd = {"Windows": ["start", str(output_path)], "Darwin": ["open", str(output_path)]}.get(platform.system(), ["xdg-open", str(output_path)])
        subprocess.Popen(cmd, shell=(platform.system()=="Windows"))
    except Exception:
        pass


def autocrop_white_paper(img_pil):
    """
    Automatically detect and crop the white paper/cardboard the parts are laid on.
    Works when white paper sits on a darker table/floor surface.

    Steps:
    1. Convert to grayscale, threshold to isolate bright white region
    2. Find the largest contour (the paper)
    3. Approximate to a quadrilateral
    4. Perspective-correct and crop to that quad

    Returns the cropped (and deskewed) image, or the original if detection fails.
    """
    try:
        import cv2
        import numpy as np

        arr = np.array(img_pil.convert("RGB"))
        gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)

        # Blur to reduce noise, then threshold for bright white areas
        blurred = cv2.GaussianBlur(gray, (7, 7), 0)
        # Use Otsu's method to find the right threshold automatically
        # then only keep regions significantly brighter than that
        otsu_val, _ = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        # Require pixels to be at least 210 brightness AND above Otsu — avoids cardboard false triggers
        white_thresh = max(210, int(otsu_val * 1.2))
        _, thresh = cv2.threshold(blurred, white_thresh, 255, cv2.THRESH_BINARY)

        # Morphological close to fill small gaps/shadows between parts
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
        closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

        contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return img_pil, False

        # Find the largest contour by area — that's the paper
        img_area = arr.shape[0] * arr.shape[1]
        largest = max(contours, key=cv2.contourArea)
        contour_area = cv2.contourArea(largest)

        # Must be at least 20% of image and not the whole image
        if contour_area < img_area * 0.20 or contour_area > img_area * 0.98:
            return img_pil, False

        # Solidity check — real paper is a solid rectangle, not a jagged blob
        # solidity = contour area / convex hull area, should be > 0.85 for paper
        hull = cv2.convexHull(largest)
        hull_area = cv2.contourArea(hull)
        if hull_area == 0:
            return img_pil, False
        solidity = contour_area / hull_area
        if solidity < 0.85:
            return img_pil, False

        # Approximate the contour to a polygon
        peri = cv2.arcLength(largest, True)
        approx = cv2.approxPolyDP(largest, 0.02 * peri, True)

        if len(approx) == 4:
            # Perfect quad — do perspective correction
            pts = approx.reshape(4, 2).astype(np.float32)

            # Order points: top-left, top-right, bottom-right, bottom-left
            rect = np.zeros((4, 2), dtype=np.float32)
            s = pts.sum(axis=1)
            rect[0] = pts[np.argmin(s)]   # top-left
            rect[2] = pts[np.argmax(s)]   # bottom-right
            diff = np.diff(pts, axis=1)
            rect[1] = pts[np.argmin(diff)]  # top-right
            rect[3] = pts[np.argmax(diff)]  # bottom-left

            # Output size based on the quad dimensions
            w1 = np.linalg.norm(rect[1] - rect[0])
            w2 = np.linalg.norm(rect[2] - rect[3])
            h1 = np.linalg.norm(rect[3] - rect[0])
            h2 = np.linalg.norm(rect[2] - rect[1])
            out_w = int(max(w1, w2))
            out_h = int(max(h1, h2))

            dst = np.array([[0, 0], [out_w-1, 0], [out_w-1, out_h-1], [0, out_h-1]], dtype=np.float32)
            M = cv2.getPerspectiveTransform(rect, dst)
            warped = cv2.warpPerspective(arr, M, (out_w, out_h))
            return Image.fromarray(warped), True

        else:
            # Not a clean quad — fall back to bounding rect crop with small padding
            x, y, w, h = cv2.boundingRect(largest)
            pad = 10
            x1 = max(0, x - pad)
            y1 = max(0, y - pad)
            x2 = min(arr.shape[1], x + w + pad)
            y2 = min(arr.shape[0], y + h + pad)
            cropped = arr[y1:y2, x1:x2]
            return Image.fromarray(cropped), True

    except Exception as e:
        return img_pil, False


_grid_cols = 8
_grid_rows = 6


async def main():
    parser = argparse.ArgumentParser(description="Scan a LEGO parts sheet photo → BrickLink upload files")
    parser.add_argument("images", nargs="*", help="One or two photos (JPG or PNG). Omit if using --folder.")
    parser.add_argument("--folder", "-f", default=None, help="Folder of images to batch process. All JPG/PNG files processed in alphabetical order.")
    parser.add_argument("--output", "-o", default="reports", help="Output directory (default: reports/)")
    parser.add_argument("--mode", default="all", choices=["all", "parts", "minifig"], help="Scan mode: 'all' (default, recommended) detects parts + minifig parts + whole minifigs. 'parts' or 'minifig' for legacy use.")
    parser.add_argument("--confidence", "-c", type=float, default=None, help="Minimum confidence threshold 0-1 (default: 0.5 for parts, 0.3 for minifig)")
    parser.add_argument("--qty", "-q", type=int, default=1, help="Quantity per part (default: 1)")
    parser.add_argument("--concurrency", type=int, default=4, help="API concurrency (default: 4, max 4 for Brickognize rate limit)")
    parser.add_argument("--debug", action="store_true", help="Save cropped images for inspection")
    parser.add_argument("--annotate", action="store_true", help="Save annotated detection image showing blob boxes")
    parser.add_argument("--preview", action="store_true", help="Draw detection boxes on photo then stop. No API calls.")
    parser.add_argument("--cols", type=int, default=8, help="Number of columns in fixed grid (default: 8)")
    parser.add_argument("--rows", type=int, default=6, help="Number of rows in fixed grid (default: 6)")
    parser.add_argument("--color", type=int, default=None, help="Force ALL images to this BrickLink color ID. Overrides --colors.")
    parser.add_argument("--colors", default=None, help="Comma-separated color IDs, one per image in order. Example: --colors 11,72,3,5 (image1=Black, image2=Dark Bluish Gray, ...)")
    parser.add_argument("--currency", default="CAD", help="Currency for prices (default: CAD). Options: CAD, USD, EUR, GBP.")
    parser.add_argument("--autocrop", action="store_true", help="Auto-detect and crop white paper from background. Use when photo has visible table/floor around the paper.")
    parser.add_argument("--no-autocrop", action="store_true", help="(Deprecated — autocrop is now opt-in via --autocrop)")
    parser.add_argument("--gap", type=int, default=15, help="Pixel gap for merging nearby blobs (default 15).")
    parser.add_argument("--padding", type=int, default=40, help="Crop padding percent around each detected part (default 40).")
    parser.add_argument("--watershed", action="store_true", default=False, help="Use watershed segmentation to split touching parts.")
    parser.add_argument("--studs", action="store_true", default=False, help="Use stud-based detection.")
    parser.add_argument("--geometric", action="store_true", default=False, help="Use geometric (Canny+Hough) detection. Better for regular rectangular bricks.")
    parser.add_argument("--brightness_bias", type=int, default=0, dest="brightness_bias",
                        help="Shift sampled color brightness before matching (-60 to +60). Positive = assume lighting makes parts look darker than they are.")
    parser.add_argument("--threshold", type=int, default=200, help="Brightness cutoff 0-255 (default 200). Auto-adjusted from background.")
    parser.add_argument("--shadow-color", type=str, default=None,
                        help="Shadow color to suppress e.g. 150,153,162")
    parser.add_argument("--bg-color", type=str, default=None,
        help="Background color as R,G,B (e.g. 0,0,0 for black, 255,255,255 for white). "
             "Overrides auto-detection. A ±30 variance window is applied automatically.")
    parser.add_argument("--color-mode", type=str, default="merge",
                        choices=["merge", "brickognize", "scan"],
                        help="Color source: 'merge' (smart merge), 'brickognize' (force Brickognize when available), or 'scan' (scanner/core-pixel only).")
    # NEW (2026): control detection downscale limit. Higher = more detail, slower.
    parser.add_argument("--max-side", type=int, default=2000,
                        help="Max image side used for detection/Brickognize crops (default 2000). "
                             "Higher values improve accuracy but are slower (e.g. 3500-4500 for HD scan).")
    args = parser.parse_args()
    global _grid_cols, _grid_rows, COLOR_MODE
    _grid_cols = args.cols
    _grid_rows = args.rows
    COLOR_MODE = getattr(args, "color_mode", "merge") or "merge"

    # Fixed grid only when user explicitly passes --cols or --rows
    use_fixed_grid = ("--cols" in sys.argv or "--rows" in sys.argv)

    # ── Mode defaults ────────────────────────────────────────────────────────
    is_minifig_mode = (args.mode == "minifig")
    if args.confidence is not None:
        conf_parts   = args.confidence
        conf_minifig = args.confidence
        conf_whole   = args.confidence
    elif is_minifig_mode:
        conf_parts   = 0.3
        conf_minifig = 0.3
        conf_whole   = 0.3
    else:
        conf_parts   = 0.5   # regular parts
        conf_minifig = 0.3   # minifig body parts — looser
        conf_whole   = 0.5   # whole minifigures

    # ── Resolve image list ───────────────────────────────────────────────────
    if args.folder:
        folder = Path(args.folder)
        if not folder.is_dir():
            print(f"Error: Folder not found: {folder}")
            sys.exit(1)
        image_paths = sorted([
            p for p in folder.iterdir()
            if p.suffix.lower() in (".jpg", ".jpeg", ".png")
        ])
        if not image_paths:
            print(f"Error: No JPG/PNG files found in {folder}")
            sys.exit(1)
        print(f"   Found {len(image_paths)} images in {folder}")
        batch_mode = True
    elif args.images:
        image_paths = [Path(p) for p in args.images]
        for p in image_paths:
            if not p.exists():
                print(f"Error: File not found: {p}")
                sys.exit(1)
        batch_mode = len(image_paths) > 2
    else:
        print("Error: Provide image file(s) or use --folder.")
        sys.exit(1)

    # ── Resolve per-image color assignments ──────────────────────────────────
    # --color overrides everything → same color for all images
    # --colors → comma list, one per image
    # neither → auto-detect per image
    if args.color is not None:
        color_per_image = [args.color] * len(image_paths)
    elif args.colors:
        try:
            color_per_image = [int(x.strip()) for x in args.colors.split(",")]
        except ValueError:
            print("Error: --colors must be comma-separated integers, e.g. --colors 11,72,3")
            sys.exit(1)
        if len(color_per_image) < len(image_paths):
            # Pad with None (auto-detect) for images without an assigned color
            color_per_image += [None] * (len(image_paths) - len(color_per_image))
    else:
        color_per_image = [None] * len(image_paths)

    # In non-batch mode keep the old 2-photo merge behaviour
    is_two_photo_merge = (not batch_mode and len(image_paths) == 2)

    # Output folder
    base_output = Path(args.output)
    if batch_mode or args.folder:
        folder_name = Path(args.folder).name if args.folder else "batch"
        output_dir = base_output / folder_name
    else:
        output_dir = base_output / image_paths[0].stem
    output_dir.mkdir(parents=True, exist_ok=True)

    debug_dir = output_dir / "crops" if args.debug else None
    if debug_dir:
        debug_dir.mkdir(exist_ok=True)

    _load_rb_parts_csv()
    print(f"\n🧱 Brickognize Sheet Scanner {'(BATCH)' if batch_mode or args.folder else ''}")
    if batch_mode or args.folder:
        print(f"   Images: {len(image_paths)} files")
        for i, (p, c) in enumerate(zip(image_paths, color_per_image)):
            cname = BRICKLINK_COLORS.get(c, (str(c),))[0] if c else "auto-detect"
            print(f"     [{i+1:2d}] {p.name}  →  color: {cname}{f' ({c})' if c else ''}")
    else:
        print(f"   Photos: {[p.name for p in image_paths]}")
    print(f"   Output: {output_dir}/")
    print(f"   Confidence: {conf_parts:.0%} (parts) / {conf_minifig:.0%} (minifig parts)  |  Currency: {args.currency}")
    if use_fixed_grid:
        print(f"   Detection: fixed grid {args.cols}×{args.rows} = {args.cols*args.rows} cells per image")
    else:
        print(f"   Detection: auto-blob (white background) — parts found individually")
    if args.mode == "all":
        print(f"   Mode: ✨ ALL — parts + minifig parts + whole minifigs")
    elif is_minifig_mode:
        print(f"   Mode: 🧍 MINIFIG — loose threshold applied to all parts")
    print()

    import io
    MIN_SIDE = 200
    semaphore = asyncio.Semaphore(args.concurrency)
    bl_creds = load_bl_credentials()
    if bl_creds:
        print("   BrickLink credentials: ✓ loaded")
    else:
        print("   BrickLink credentials: ✗ NOT found — color will use image detection only")

    all_photo_results = []  # one list of results per photo
    all_batch_results = []  # flat list accumulating ALL results across all images (batch mode)

    async with aiohttp.ClientSession() as session:
        for photo_idx, image_path in enumerate(image_paths):
            img_color = color_per_image[photo_idx]  # color forced for this specific image
            color_label = BRICKLINK_COLORS.get(img_color, (str(img_color),))[0] if img_color else "auto"
            print(f"\n📷 [{photo_idx+1}/{len(image_paths)}] {image_path.name}  (color: {color_label})")
            img_orig = Image.open(image_path).convert("RGB")
            orig_w, orig_h = img_orig.width, img_orig.height
            # Downscale large phone photos — keeps detection fast.
            # NEW (2026): configurable via --max-side for “HD scan” in the GUI.
            MAX_SIDE = max(800, int(getattr(args, "max_side", 2000) or 2000))
            if max(orig_w, orig_h) > MAX_SIDE:
                _ds = MAX_SIDE / max(orig_w, orig_h)
                img = img_orig.resize((int(orig_w*_ds), int(orig_h*_ds)), Image.LANCZOS)
                _bbox_scale = 1.0 / _ds   # factor to convert detection coords → original coords
                print(f"   Size: {img.width}×{img.height}px (downscaled from {orig_w}×{orig_h})")
            else:
                img = img_orig
                _bbox_scale = 1.0
                print(f"   Size: {img.width}×{img.height}px")

            # Auto-crop white paper from background — only when explicitly requested
            if getattr(args, 'autocrop', False):
                cropped_img, did_crop = autocrop_white_paper(img)
                if did_crop:
                    print(f"   ✂  Auto-cropped: {img.width}×{img.height} → {cropped_img.width}×{cropped_img.height}px")
                    img = cropped_img
                else:
                    print(f"   ✂  Auto-crop: no white paper detected, using full image")

            # Detect parts
            if use_fixed_grid:
                print(f"🔍 Fixed grid: {args.cols}×{args.rows} = {args.cols*args.rows} cells")
            else:
                print(f"🔍 Auto-detecting parts on white background...")
            _bg_color_arg = None
            if getattr(args, "bg_color", None):
                try:
                    _bg_color_arg = tuple(int(x) for x in args.bg_color.split(","))
                except Exception:
                    _bg_color_arg = None
            _shadow_color_arg = None
            if getattr(args, "shadow_color", None):
                try:
                    _shadow_color_arg = tuple(int(x) for x in args.shadow_color.split(","))
                except Exception:
                    _shadow_color_arg = None
            if getattr(args, "studs", False):
                if _bg_color_arg is not None:
                    flood_result = detect_parts_flood(img, _bg_color_arg, _shadow_color_arg,
                                                      padding_pct=args.padding)
                    if flood_result:
                        boxes = flood_result
                        print(f"   ✅  Flood-fill detection: {len(boxes)} parts")
                    else:
                        print("   ⚠  Flood-fill found nothing — falling back to standard")
                        boxes = detect_parts(img, threshold=args.threshold, gap=args.gap,
                                             padding_pct=args.padding,
                                             hue_mode=getattr(args, "hue", False),
                                             use_fixed_grid=use_fixed_grid,
                                             bg_color=_bg_color_arg,
                                             shadow_color=_shadow_color_arg,
                                             use_watershed=False)
                else:
                    print("   ⚠  Flood-fill requires background color — falling back to standard")
                    boxes = detect_parts(img, threshold=args.threshold, gap=args.gap,
                                         padding_pct=args.padding,
                                         hue_mode=getattr(args, "hue", False),
                                         use_fixed_grid=use_fixed_grid,
                                         bg_color=_bg_color_arg,
                                         shadow_color=_shadow_color_arg,
                                         use_watershed=False)
            elif getattr(args, "geometric", False):
                geo_result = detect_parts_geometric(img, bg_color=_bg_color_arg, padding_pct=args.padding)
                if geo_result:
                    boxes = geo_result
                    print(f"   ✅  Geometric detection: {len(boxes)} parts")
                else:
                    print("   ⚠  Geometric detection found nothing — falling back to blob")
                    boxes = detect_parts(img, threshold=args.threshold, gap=args.gap, padding_pct=args.padding,
                                         hue_mode=getattr(args, "hue", False),
                                         use_fixed_grid=use_fixed_grid,
                                         bg_color=_bg_color_arg,
                                         shadow_color=_shadow_color_arg,
                                         use_watershed=False)
            else:
                boxes = detect_parts(img, threshold=args.threshold, gap=args.gap,
                                     padding_pct=args.padding,
                                     hue_mode=getattr(args, "hue", False),
                                     use_fixed_grid=use_fixed_grid,
                                     bg_color=_bg_color_arg,
                                     shadow_color=_shadow_color_arg,
                                     use_watershed=args.watershed)
            print(f"   ✅ {len(boxes)} parts detected")

            if len(boxes) == 0:
                print("   ⚠ No parts detected.")
                all_photo_results.append([])
                continue

            # Preview mode — only on first photo
            if args.preview and photo_idx == 0:
                preview_path = output_dir / "preview-detection.jpg"
                preview_detections(img, boxes, preview_path)
                print(f"   Preview only — no API calls made.")
                sys.exit(0)

            # Always save annotated image if --annotate or --debug
            if getattr(args, "annotate", False) or args.debug:
                from PIL import ImageDraw
                ann = img.copy().convert("RGB")
                draw = ImageDraw.Draw(ann)
                palette = [(255,50,50),(50,200,50),(50,100,255),(255,180,0),(200,0,200),(0,200,200)]
                for i, (x1,y1,x2,y2) in enumerate(boxes):
                    color = palette[i % len(palette)]
                    for t in range(3):
                        draw.rectangle([x1-t,y1-t,x2+t,y2+t], outline=color)
                    lbl = str(i+1)
                    lx,ly = x1+4,y1+2
                    draw.rectangle([lx-2,ly-1,lx+len(lbl)*8+2,ly+14], fill=color)
                    draw.text((lx,ly), lbl, fill=(255,255,255))
                # Downscale for display
                mw,mh = ann.size
                if max(mw,mh) > 1400:
                    sc = 1400/max(mw,mh)
                    ann = ann.resize((int(mw*sc),int(mh*sc)), Image.LANCZOS)
                ann_path = output_dir / f"annotated-p{photo_idx+1}.jpg"
                ann.save(ann_path)
                print(f"ANNOTATED_PATH={ann_path}")

            # Crop
            crops = []
            for i, (x1, y1, x2, y2) in enumerate(boxes):
                crop = img.crop((x1, y1, x2, y2))
                crops.append(crop)
                if debug_dir:
                    # NEW (2026): save *full-resolution* crops for the GUI preview.
                    # Detection + Brickognize are done on the downscaled image for speed,
                    # but the GUI should display true detail.
                    fx1 = int(x1 * _bbox_scale); fy1 = int(y1 * _bbox_scale)
                    fx2 = int(x2 * _bbox_scale); fy2 = int(y2 * _bbox_scale)
                    fx1 = max(0, min(fx1, orig_w - 1)); fx2 = max(1, min(fx2, orig_w))
                    fy1 = max(0, min(fy1, orig_h - 1)); fy2 = max(1, min(fy2, orig_h))
                    full_crop = img_orig.crop((fx1, fy1, fx2, fy2))
                    full_crop.save(debug_dir / f"p{photo_idx+1}_crop_{i+1:03d}.jpg", "JPEG", quality=92)

            # Encode all crops in parallel, then fire all API calls at once
            print(f"🌐 Querying Brickognize for {len(crops)} parts...")
            import concurrent.futures as _cf
            def _encode(args_t):
                idx, crop = args_t
                cw, ch = crop.size
                if min(cw, ch) < MIN_SIDE:
                    sc = MIN_SIDE / min(cw, ch)
                    crop = crop.resize((int(cw*sc), int(ch*sc)), Image.LANCZOS)
                buf = io.BytesIO()
                crop.save(buf, "JPEG", quality=82)
                return idx, buf.getvalue()
            crop_bytes_list = [None] * len(crops)
            with _cf.ThreadPoolExecutor(max_workers=min(8, len(crops))) as _pool:
                futures = {_pool.submit(_encode, (i, c)): i for i, c in enumerate(crops)}
                for fut in _cf.as_completed(futures):
                    try:
                        idx, data = fut.result()
                        crop_bytes_list[idx] = data
                    except Exception as enc_err:
                        print(f"  ⚠ Crop encode error (idx {futures[fut]}): {enc_err}")
            brickognize_tasks = [
                query_brickognize(session, crop_bytes_list[i],
                                  f"p{photo_idx+1}_{i+1:03d}.jpg", semaphore)
                for i in range(len(crops))
            ]
            # Pipeline: as each Brickognize response arrives, immediately kick off BL color fetch
            # This overlaps network waits instead of doing them sequentially
            color_fetch_tasks = {}  # part_id → task, deduplicated

            async def query_and_prefetch(task, idx):
                """Run one Brickognize query then immediately start BL color fetch
                for top-8 candidates — re-ranking may pick any of them."""
                resp = await task
                if resp and resp.get("items") and bl_creds:
                    # Prefetch colors for ALL top-8 candidates regardless of score —
                    # size-penalty re-ranking may promote a low-score candidate,
                    # and the on-demand fallback stalls the pipeline if it misses.
                    for cand in resp["items"][:8]:
                        pid = cand.get("id", "")
                        if pid and pid not in color_fetch_tasks:
                            color_fetch_tasks[pid] = asyncio.ensure_future(
                                get_bl_known_colors_async(session, pid, bl_creds)
                            )
                return resp

            # Fire all Brickognize queries; color prefetches start as results come in
            pipelined = [query_and_prefetch(t, i) for i, t in enumerate(brickognize_tasks)]
            responses = await asyncio.gather(*pipelined)

            # Wait for any still-running color fetches
            if color_fetch_tasks:
                await asyncio.gather(*color_fetch_tasks.values())
                hits   = [pid for pid in color_fetch_tasks if _bl_colors_cache.get(pid)]
                misses = [pid for pid in color_fetch_tasks if not _bl_colors_cache.get(pid)]
                print(f"   Color cache: {len(hits)} hits, {len(misses)} misses")
                if misses:
                    print(f"   Cache misses: {misses}")

            # Compute per-box size ratios for size-aware re-ranking
            box_wh_list = [(x2 - x1, y2 - y1) for x1, y1, x2, y2 in boxes]

            # Head mode: no color lookup needed
            photo_results = []
            for i, (resp, crop) in enumerate(zip(responses, crops)):
                bx1, by1, bx2, by2 = boxes[i]
                entry = {
                    "index": i + 1,
                    "part_id": None, "part_name": "Unknown",
                    "confidence": 0.0, "color_id": 0,
                    "color_name": "Unknown", "color_rgb": (200,200,200),
                    "color_method": None,
                    "thumb_url": "", "crop_path": f"crops/p{photo_idx+1}_crop_{i+1:03d}.jpg" if args.debug else "",
                    "source_image": str(image_path),
                    "source_image_name": image_path.name,
                    "bbox": [int(bx1*_bbox_scale), int(by1*_bbox_scale),
                             int(bx2*_bbox_scale), int(by2*_bbox_scale)],
                    "photo_idx": photo_idx,
                }
                if resp and resp.get("items"):
                    # Re-rank candidates using size-adjusted score
                    sr = box_size_ratio(box_wh_list[i], box_wh_list)
                    candidates_ranked = []
                    for cand in resp["items"][:8]:  # consider top-8
                        pen = size_score_penalty(cand.get("name", ""), sr)
                        adjusted = cand.get("score", 0) - pen
                        candidates_ranked.append((adjusted, pen, cand))
                    candidates_ranked.sort(key=lambda x: x[0], reverse=True)
                    adj_score, penalty, best = candidates_ranked[0]
                    if penalty > 0.05:
                        orig = best.get("name","?"); orig_score = best.get("score",0)
                        print(f"   📐  [{i+1}] size_ratio={sr:.1f}x → "
                              f"penalized '{orig}' ({orig_score:.0%} → {adj_score:.0%}, -{penalty:.2f})")
                    conf = best.get("score", 0)
                    part_id_candidate = best.get("id", "")

                    # Extract Brickognize color from img_url: .../part/{id}/{color_id}.webp
                    _img_url = best.get("img_url", "")
                    _cm = re.search(r"/(\d+)\.webp$", _img_url)
                    brickognize_color_id = int(_cm.group(1)) if _cm and _cm.group(1) != "0" else None

                    # Save top-5 Brickognize alternatives for GUI picker
                    brickognize_alts = [
                        {"id": c.get("id",""), "name": c.get("name",""),
                         "score": round(adj, 3), "type": c.get("type","P"),
                         "img_url": c.get("img_url","")}
                        for adj, pen, c in candidates_ranked[1:6]
                        if c.get("id")
                    ]
                    entry["alternatives"] = brickognize_alts
                    item_type = best.get("type", "P")  # "P"=part, "M"=minifig, "S"=set
                    # Brickognize returns type="P" for all alphabetic IDs regardless of catalog.
                    # Use BL API to definitively determine type for any alphabetic ID.
                    if part_id_candidate and not part_id_candidate[0].isdigit():
                        item_type = verify_item_type_with_bl(part_id_candidate, bl_creds)

                    # Fetch mold variants from BrickLink and append to alternatives
                    if bl_creds and item_type == "P":
                        mold_vars = fetch_mold_variants(part_id_candidate, bl_creds)
                        if mold_vars:
                            entry["alternatives"] = entry.get("alternatives", []) + mold_vars
                            entry["has_mold_variants"] = True

                    # Print variants (pb, pr, pat, ps patterns): auto-add base mold as alternative
                    # e.g. 60581pb038R → base mold is 60581
                    if item_type == "P" and part_id_candidate:
                        import re as _re
                        base_match = _re.match(r'^([0-9]+[a-z]?)(pb|pr|pat|ps|p)[0-9]', part_id_candidate.lower())
                        if base_match:
                            base_pid = base_match.group(1)
                            if base_pid != part_id_candidate.lower():
                                already = any(
                                    a.get("part_id","").lower() == base_pid or
                                    a.get("id","").lower() == base_pid
                                    for a in entry.get("alternatives", []))
                                if not already:
                                    entry["alternatives"] = entry.get("alternatives", []) + [{
                                        "part_id":   base_pid,
                                        "name":      f"Base mold: {base_pid}",
                                        "score":     0.0,
                                        "is_mold_variant": True,
                                    }]
                                    entry["has_mold_variants"] = True
                                    print(f"   🔀  Print variant {part_id_candidate} → base mold {base_pid} added as alternative")
                    # Choose threshold based on what was identified
                    is_whole_minifig = (item_type == "M")
                    if is_whole_minifig:
                        effective_conf = conf_whole
                    elif is_minifig_part(part_id_candidate):
                        effective_conf = conf_minifig
                    else:
                        effective_conf = conf_parts
                    if conf >= effective_conf:
                        entry["part_id"] = part_id_candidate
                        entry["part_name"] = best.get("name", "Unknown")
                        entry["confidence"] = conf
                        entry["size_ratio"] = round(sr, 2)
                        entry["thumb_url"] = best.get("img_url", "")
                        entry["item_type"] = item_type  # pass through to GUI
                        # Whole minifigs: no color lookup needed
                        if is_whole_minifig:
                            entry["color_id"] = 0
                            entry["color_name"] = "—"
                            entry["color_method"] = "minifig"
                            entry["color_rgb"] = (180, 180, 180)
                            entry["known_color_ids"] = []
                        elif img_color is not None and img_color in BRICKLINK_COLORS:
                            # Override forced color if BL known colors proves exactly 1 color exists
                            known_ids = _bl_colors_cache.get(part_id_candidate, [])
                            if len(known_ids) == 1 and known_ids[0] != img_color and known_ids[0] in BRICKLINK_COLORS:
                                real_id = known_ids[0]
                                print(f"   🎨  Forced color overridden: part only exists in {BRICKLINK_COLORS[real_id][0]} (id {real_id})")
                                entry["color_id"]     = real_id
                                entry["color_name"]   = BRICKLINK_COLORS[real_id][0]
                                entry["color_method"] = "exact"
                            else:
                                entry["color_id"]     = img_color
                                entry["color_name"]   = BRICKLINK_COLORS[img_color][0]
                                entry["color_method"] = "forced"
                            entry["color_rgb"]       = BRICKLINK_COLORS.get(entry["color_id"], ("?", (200,200,200)))[1]
                            entry["known_color_ids"] = known_ids
                        else:
                            # If winner's colors weren't prefetched, fetch sync now
                            # (safe — we're inside an async context but requests is blocking)
                            if part_id_candidate not in _bl_colors_cache and bl_creds:
                                try:
                                    _url = f"https://api.bricklink.com/api/store/v1/items/part/{part_id_candidate}/colors"
                                    _r = requests.get(_url, auth=_get_oauth(bl_creds), timeout=8)
                                    if _r.status_code == 200:
                                        _ids = [c["color_id"] for c in _r.json().get("data", [])]
                                        _bl_colors_cache[part_id_candidate] = _ids
                                        print(f"   🎨  on-demand fetch {part_id_candidate}: {_ids}")
                                    else:
                                        _bl_colors_cache[part_id_candidate] = []
                                except Exception as _e:
                                    print(f"   ⚠  on-demand fetch failed {part_id_candidate}: {_e}")
                                    _bl_colors_cache[part_id_candidate] = []
                            known_ids = _bl_colors_cache.get(part_id_candidate, [])
                            # NEW (2026): compute BOTH sources, then merge.
                            # - scan_color = our core-pixel sampler (robust against shadows/edges)
                            # - brickognize_color = inferred from Brickognize img_url (when available)

                            # Source A: scanner/core-pixel color
                            sc_id, sc_name, sc_method = resolve_color_from_cache(
                                part_id_candidate, crop, brightness_bias=args.brightness_bias
                            )
                            sc_conf = _color_conf_from_method(sc_method)
                            scan_color = {
                                "color_id": sc_id,
                                "color_name": sc_name,
                                "color_method": sc_method,
                                "color_conf": sc_conf,
                            }

                            # Source B: Brickognize color (if present)
                            brick_color = {"color_id": 0, "color_name": "Unknown", "color_method": "brickognize", "color_conf": 0.0}
                            if brickognize_color_id and brickognize_color_id in BRICKLINK_COLORS:
                                bc_id = brickognize_color_id
                                bc_name = BRICKLINK_COLORS[bc_id][0]
                                # Confidence heuristic: if BrickLink known-colors includes it, treat as stronger.
                                bc_conf = 0.80 if (known_ids and bc_id in known_ids) else 0.62
                                brick_color = {
                                    "color_id": bc_id,
                                    "color_name": bc_name,
                                    "color_method": "brickognize",
                                    "color_conf": bc_conf,
                                }
                                if bc_id in known_ids:
                                    print(f"   🎨  Brickognize color: {bc_name} (id {bc_id})")
                                else:
                                    print(f"   🎨  Brickognize color (not in known list): {bc_name} (id {bc_id})")

                            # Final merge decision
                            merged = merge_color_decision(scan_color, brick_color, known_ids)

                            entry["scan_color_id"]     = scan_color["color_id"]
                            entry["scan_color_name"]   = scan_color["color_name"]
                            entry["scan_color_method"] = scan_color["color_method"]
                            entry["scan_color_conf"]   = round(float(scan_color["color_conf"]), 3)

                            entry["brickognize_color_id"]   = brick_color["color_id"]
                            entry["brickognize_color_name"] = brick_color["color_name"]
                            entry["brickognize_color_conf"] = round(float(brick_color["color_conf"]), 3)

                            entry["color_id"]      = merged["color_id"]
                            entry["color_name"]    = merged["color_name"]
                            entry["color_method"]  = merged.get("color_method", "matched")
                            entry["color_conf"]    = round(float(merged.get("color_conf", 0.0)), 3)
                            entry["color_source"]  = merged.get("color_source", "scan")

                            entry["color_rgb"] = BRICKLINK_COLORS.get(entry["color_id"], (entry["color_name"], (200,200,200)))[1]
                            entry["known_color_ids"] = known_ids
                        if is_whole_minifig:
                            type_tag = " 👤 minifig"
                        elif is_minifig_part(part_id_candidate):
                            type_tag = " 🧍 minifig-part"
                        else:
                            type_tag = " 🧱 part"
                        print(f"   [{i+1:3d}] ✓ {part_id_candidate:8s}  {best.get('name','')[:24]:24s}  {conf:.0%}{type_tag}")
                    else:
                        hint = ""
                        if is_whole_minifig: hint = " (whole minifig)"
                        elif is_minifig_part(part_id_candidate): hint = " (minifig part)"
                        if sr < 0.4: hint += " ⚠ small crop"
                        elif sr > 4.0: hint += " ⚠ large crop"
                        print(f"   [{i+1:3d}] ✗ Low conf ({conf:.0%}){hint}")
                else:
                    print(f"   [{i+1:3d}] ✗ No response")
                photo_results.append(entry)
            all_photo_results.append(photo_results)
            # In batch mode, accumulate all results into one flat list
            if batch_mode or args.folder:
                all_batch_results.extend(photo_results)
                identified_this = sum(1 for r in photo_results if r.get("part_id"))
                print(f"   → {identified_this}/{len(photo_results)} identified")

    # ── Resolve final results list ───────────────────────────────────────────
    if batch_mode or args.folder:
        # Batch: all images contribute to one combined XML
        results = all_batch_results
        print(f"\n📦 Batch complete: {len(image_paths)} images processed")
    elif is_two_photo_merge and len(all_photo_results) == 2 and all_photo_results[0] and all_photo_results[1]:
        # Two-photo merge of the same tray
        r1, r2 = all_photo_results[0], all_photo_results[1]
        n = max(len(r1), len(r2))
        print(f"\n🔀 Merging results from 2 photos ({len(r1)} + {len(r2)} parts)...")
        results = []
        for i in range(n):
            a = r1[i] if i < len(r1) else None
            b = r2[i] if i < len(r2) else None
            if a is None:
                results.append(b)
            elif b is None:
                results.append(a)
            elif a.get("part_id") and not b.get("part_id"):
                results.append(a)
            elif b.get("part_id") and not a.get("part_id"):
                results.append(b)
            elif a.get("confidence", 0) >= b.get("confidence", 0):
                results.append(a)
            else:
                results.append(b)
        identified_count = sum(1 for r in results if r.get("part_id"))
        print(f"   Merged: {identified_count}/{n} identified")
    else:
        results = all_photo_results[0] if all_photo_results else []

    # ── Step 6: Export files ────────────────────────────────────────────────
    identified = [r for r in results if r["part_id"]]
    unique_parts = len({(r["part_id"], r["color_id"]) for r in identified})
    print(f"\n📊 {len(identified)}/{len(results)} parts identified → {unique_parts} unique part+color combinations")
    if bl_creds:
        print(f"   💲 Fetching medium prices from BrickLink...")

    # Fetch prices concurrently — one thread per unique part+color, max 8 workers
    top_prices = {}
    if bl_creds and identified:
        import concurrent.futures
        merged_keys = {}
        for r in identified:
            key = (r["part_id"], r["color_id"])
            if key not in merged_keys:
                merged_keys[key] = r
        print(f"   💲 Fetching {len(merged_keys)} prices concurrently...")
        def _fetch_price(entry):
            return (entry["part_id"], entry["color_id"]), fetch_medium_price(
                entry["part_id"], entry["color_id"], bl_creds, currency=args.currency,
                item_type=entry.get("item_type", "P"))
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(merged_keys))) as ex:
            futures = {ex.submit(_fetch_price, e): e for e in merged_keys.values()}
            for fut in concurrent.futures.as_completed(futures):
                key, p = fut.result()
                if p:
                    top_prices[key] = float(p)

    # Inject prices into results for GUI display
    for r in results:
        pid, cid = r.get("part_id"), r.get("color_id")
        if pid and cid is not None:
            p = top_prices.get((pid, cid))
            if p: r["price"] = p

    # XML
    xml_path = output_dir / "scan-results.xml"
    xml_content = build_xml(identified, qty=args.qty, bl_creds=None, currency=args.currency, prices=top_prices)
    xml_path.write_text(xml_content, encoding="utf-8")
    print(f"   ✅ XML saved:  {xml_path}")

    # TSV
    tsv_path = output_dir / "scan-results.tsv"
    tsv_content = build_tsv(identified, qty=args.qty)
    tsv_path.write_text(tsv_content, encoding="utf-8")
    print(f"   ✅ TSV saved:  {tsv_path}")

    # HTML report
    html_path = output_dir / "scan-results.html"
    html_content = build_html_report(results, " + ".join(p.name for p in image_paths), output_dir)
    html_path.write_text(html_content, encoding="utf-8")
    print(f"   ✅ HTML report: {html_path}")

    # JSON — sanitize numpy types that break json.dumps
    class _NpEnc(json.JSONEncoder):
        def default(self, obj):
            import numpy as np
            if isinstance(obj, np.integer): return int(obj)
            if isinstance(obj, np.floating): return float(obj)
            if isinstance(obj, np.ndarray): return obj.tolist()
            return super().default(obj)
    for _r in results:
        if "color_rgb" in _r and _r["color_rgb"] is not None:
            _r["color_rgb"] = [int(x) for x in _r["color_rgb"]]
    json_path = output_dir / "scan-results.json"
    json_path.write_text(json.dumps(results, indent=2, cls=_NpEnc), encoding="utf-8")

    print(f"\n✨ Done! Open {html_path} to review results before uploading.")
    print(f"   → Import scan-results.xml in BrickStore to push to BrickLink")
    print(f"   → Or paste scan-results.tsv on BrickLink's upload page\n")


if __name__ == "__main__":
    asyncio.run(main())
