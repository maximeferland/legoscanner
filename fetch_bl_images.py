#!/usr/bin/env python3
"""
fetch_bl_images.py
Downloads BrickLink catalog images using the BrickLink API's known image URL format.

The correct BrickLink image URL format (from the API's img_url field) is:
  https://img.bricklink.com/ItemImage/PN/{color_id}/{part_id}.png

The issue: Rebrickable part numbers differ from BrickLink part numbers for many parts.
Solution: Use the BrickLink API GET /items/part/{id} to validate + get the right img_url.
But since that's slow for 17k parts, we use the Brickognize thumbnail URL format instead:
  https://storage.googleapis.com/brickognize-static/thumbnails/v2.21.2/part/{id}/0.webp

This is what scan-heads.py already uses successfully — it's the most reliable source.
"""

import csv, gzip, io, json, time, re, sys
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import requests
except ImportError:
    print("pip install requests")
    sys.exit(1)

# ── Config ────────────────────────────────────────────────────────────────────
OUT_DIR       = Path("images")
MANIFEST_PATH = OUT_DIR / "manifest.json"
PARTS_CACHE   = Path("parts_list.json")
MAX_WORKERS   = 8

REBRICKABLE_PARTS_URL = "https://cdn.rebrickable.com/media/downloads/parts.csv.gz"

# Exclude printed/decorated/stickered variants and minifig body assemblies
EXCLUDE_RE = re.compile(
    r'pb\d|pat\d|stk\d|\dstk|pr\d{4}|'
    r'^973c|^970c|^3626c|c\d{2}$',
    re.IGNORECASE
)

def img_url_rebrickable(part_id: str) -> str:
    # Rebrickable part images — CC licensed, officially allowed
    return f"https://cdn.rebrickable.com/media/parts/photos/0/{part_id}.jpg"

def img_url_rebrickable_elements(part_id: str) -> str:
    # Alternative Rebrickable URL format
    return f"https://cdn.rebrickable.com/media/parts/elements/{part_id}.jpg"

def fetch_part_ids() -> list:
    import gzip
    print("📋  Downloading Rebrickable parts list...")
    r = requests.get(REBRICKABLE_PARTS_URL, timeout=60, stream=True)
    r.raise_for_status()
    raw = gzip.decompress(r.content)
    reader = csv.DictReader(io.StringIO(raw.decode("utf-8")))
    all_ids = []
    skipped = 0
    for row in reader:
        pid = row.get("part_num", "").strip()
        if not pid:
            continue
        if EXCLUDE_RE.search(pid):
            skipped += 1
            continue
        all_ids.append(pid)
    ids = sorted(set(all_ids))
    print(f"✅  {len(ids)} parts  ({skipped} excluded)")
    return ids

def download_image(part_id: str, out_path: Path) -> tuple:
    if out_path.exists() and out_path.stat().st_size > 500:
        return part_id, True, "cached"
    for url in [img_url_rebrickable(part_id), img_url_rebrickable_elements(part_id)]:
        try:
            r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code == 200 and len(r.content) > 500:
                out_path.write_bytes(r.content)
                return part_id, True, "ok"
        except Exception:
            continue
    return part_id, False, "not found"

def main():
    OUT_DIR.mkdir(exist_ok=True)

    manifest = {}
    if MANIFEST_PATH.exists():
        try:
            manifest = json.loads(MANIFEST_PATH.read_text())
        except Exception:
            pass

    if PARTS_CACHE.exists():
        part_ids = json.loads(PARTS_CACHE.read_text())
        if not part_ids:
            PARTS_CACHE.unlink()
            part_ids = fetch_part_ids()
            PARTS_CACHE.write_text(json.dumps(part_ids, indent=2))
        else:
            print(f"📋  Loaded {len(part_ids)} parts from cache")
    else:
        part_ids = fetch_part_ids()
        PARTS_CACHE.write_text(json.dumps(part_ids, indent=2))
        print(f"💾  Saved → {PARTS_CACHE}")

    todo = [p for p in part_ids if manifest.get(p) != "ok"]
    already = len(part_ids) - len(todo)
    print(f"\n📥  {len(todo)} to download  ({already} already done)\n")
    if not todo:
        print("✅  All done!")
        return

    done = ok = fail = 0
    t0 = time.time()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {
            ex.submit(download_image, pid, OUT_DIR / f"{pid}.png"): pid
            for pid in todo
        }
        for fut in as_completed(futures):
            pid, success, reason = fut.result()
            done += 1
            manifest[pid] = "ok" if success else f"fail:{reason}"
            if success: ok += 1
            else: fail += 1

            if done % 200 == 0 or done == len(todo):
                elapsed = time.time() - t0
                rate = done / max(elapsed, 0.1)
                eta = (len(todo) - done) / rate
                print(f"   {done}/{len(todo)}  ✓{ok} ✗{fail}  {rate:.1f}/s  ETA {eta/60:.0f}min")
                MANIFEST_PATH.write_text(json.dumps(manifest, indent=2))

    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2))
    print(f"\n✅  Done — {ok} downloaded, {fail} failed")
    print(f"📁  {OUT_DIR.resolve()}")

if __name__ == "__main__":
    main()
