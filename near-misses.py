"""
near-misses.py — Find almost-complete minifigs from your Fig Finder run
========================================================================
TRUE ROI = fig sells for − parts you pull (your listed prices) − parts to buy (market price)
A negative ROI means assembling this fig destroys more value than it creates.

Usage:
    py near-misses.py torsos.xml -i inventory.xml --async -c 10

Output:
    near_misses_report.html
    near-misses-wanted.xml  (BrickLink Wanted List — only profitable figs)
"""

import argparse
import asyncio
import sys
import time
import webbrowser
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path
from typing import List, Optional

from fig_finder.api.client import BrickLinkClient
from fig_finder.api.endpoints import BrickLinkAPI
from fig_finder.api.models import InventoryItem
from fig_finder.cache.disk import DiskCache
from fig_finder.config import load_config
from fig_finder.core.finder import MinifigFinder
from fig_finder.parsers.inventory import parse_inventory_xml
from fig_finder.parsers.parts import parse_parts_xml


# ============================================================
# Inventory parser that also reads <PRICE> from the XML
# ============================================================
def parse_inventory_xml_with_price(path) -> list:
    """Like fig_finder's parse_inventory_xml but also reads <PRICE> and <QTY>."""
    import xml.etree.ElementTree as ET
    path = Path(path)
    root = ET.parse(path).getroot()
    items = []
    for item in root.findall("ITEM"):
        # Skip non-parts (minifigs in inventory, sets, etc.)
        item_type = item.findtext("ITEMTYPE", default="P")
        if item_type not in ("P", "M"):
            continue
        color_text = item.findtext("COLOR", default="0")
        color_id = int(color_text) if color_text and color_text.strip() not in ("0", "") else None
        qty_text = item.findtext("QTY", item.findtext("MINQTY", "1"))
        try:
            qty = int(qty_text)
        except (ValueError, TypeError):
            qty = 1
        price_text = item.findtext("PRICE", "0")
        try:
            my_price = float(price_text)
        except (ValueError, TypeError):
            my_price = 0.0
        inv = InventoryItem(
            part_no=item.findtext("ITEMID", ""),
            color_id=color_id,
            quantity=qty,
            remarks=item.findtext("REMARKS", "") or "",
        )
        inv.my_price = my_price  # attach price dynamically
        items.append(inv)
    return items


# ============================================================
# Near-miss checker
# ============================================================
def check_near_miss(required_parts, inventory_lookup, color_map, max_missing=2):
    have = []
    missing = []

    for subset in required_parts:
        for entry in subset.get("entries", []):
            if entry.get("is_alternate"):
                continue
            part_no = entry["item"]["no"]
            if entry["item"].get("type", "PART") not in ("PART", "MINIFIG"):
                continue

            color_id = entry.get("color_id")
            qty_needed = entry.get("quantity", 1)
            color_name = color_map.get(color_id, "Unknown") if color_id else "N/A"
            part_name = entry["item"].get("name", part_no)
            key = (part_no, color_id)
            inv_item = inventory_lookup.get(key)

            if inv_item and inv_item.quantity >= qty_needed:
                have.append({
                    "no": part_no, "name": part_name,
                    "color_id": color_id, "color_name": color_name,
                    "qty": qty_needed, "remarks": inv_item.remarks,
                    "my_price": getattr(inv_item, "my_price", 0.0) or 0.0,
                })
            else:
                missing.append({
                    "no": part_no, "name": part_name,
                    "color_id": color_id, "color_name": color_name,
                    "qty": qty_needed,
                    "have_qty": inv_item.quantity if inv_item else 0,
                })

    if not missing:
        return None
    if len(missing) > max_missing:
        return None
    return {"have": have, "missing": missing}


def build_inventory_lookup(inventory):
    lookup = {}
    for item in inventory:
        key = (item.part_no, item.color_id)
        if key not in lookup:
            lookup[key] = item
        else:
            existing = lookup[key]
            merged = InventoryItem(
                part_no=item.part_no, color_id=item.color_id,
                quantity=existing.quantity + item.quantity,
                remarks=existing.remarks,
            )
            # Keep the price from the existing entry (first one wins)
            merged.my_price = getattr(existing, "my_price", 0.0)
            lookup[key] = merged
    return lookup


# ============================================================
# Enrich missing parts with market prices
# ============================================================
def enrich_missing_parts(missing_parts, api):
    enriched = []
    for p in missing_parts:
        try:
            price = api.get_part_price(p["no"], p.get("color_id"))
        except Exception:
            price = 0.0
        enriched.append({**p, "price": price})
    return enriched


# ============================================================
# True ROI
# ============================================================
def compute_roi(r):
    pulled_value = sum(p.get("my_price", 0.0) * p.get("qty", 1) for p in r["have"])
    missing_cost = sum(p.get("price", 0.0) * p.get("qty", 1) for p in r["missing"])
    true_cost = pulled_value + missing_cost
    roi = r["fig_price"] - true_cost
    r["pulled_value"] = pulled_value
    r["missing_cost"] = missing_cost
    r["true_cost"] = true_cost
    r["roi"] = roi
    r["good_deal"] = roi > 0
    return r


# ============================================================
# Wanted List XML — only missing parts from profitable figs
# ============================================================
def build_wanted_list_xml(results):
    wanted = {}
    for r in results:
        if not r.get("good_deal"):
            continue
        for p in r["missing"]:
            key = (p["no"], p.get("color_id", 0))
            if key not in wanted:
                wanted[key] = {"no": p["no"], "color_id": p.get("color_id", 0),
                               "name": p["name"], "qty": p.get("qty", 1)}
            else:
                wanted[key]["qty"] += p.get("qty", 1)

    root = ET.Element("WANTEDLIST")
    for d in wanted.values():
        item = ET.SubElement(root, "ITEM")
        ET.SubElement(item, "ITEMTYPE").text = "P"
        ET.SubElement(item, "ITEMID").text = str(d["no"])
        ET.SubElement(item, "COLOR").text = str(d["color_id"])
        ET.SubElement(item, "MINQTY").text = str(d["qty"])
        ET.SubElement(item, "REMARKS").text = d["name"][:50]

    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")
    return "\n".join(xml_str.split("\n")[1:])


# ============================================================
# HTML Report
# ============================================================
def generate_report(results, output_file):
    from datetime import datetime
    date_str = datetime.now().strftime("%B %d, %Y")

    # Compute ROI for all
    for r in results:
        compute_roi(r)

    # Sort: good deals first, then by ROI desc
    results_sorted = sorted(results, key=lambda r: (0 if r["good_deal"] else 1, -r["roi"]))

    total = len(results_sorted)
    good_count = sum(1 for r in results_sorted if r["good_deal"])
    bad_count = total - good_count
    missing_1 = sum(1 for r in results_sorted if len(r["missing"]) == 1)
    missing_2 = sum(1 for r in results_sorted if len(r["missing"]) == 2)

    rows = ""
    for r in results_sorted:
        # Missing parts HTML
        parts_html = ""
        for p in r["missing"]:
            ellipsis = "\u2026" if len(p["name"]) > 50 else ""
            price_str = ("$" + "{:.2f}".format(p["price"])) if p.get("price") else "?"
            parts_html += (
                '<div class="missing-part">'
                '<span class="part-no">' + str(p["no"]) + '</span>'
                '<span class="part-name">' + p["name"][:50] + ellipsis + '</span>'
                '<span class="part-color">' + p["color_name"] + '</span>'
                '<span class="part-qty">\u00d7' + str(p["qty"]) + '</span>'
                '<span class="part-price">' + price_str + '</span>'
                '<a href="https://www.bricklink.com/v2/catalog/catalogitem.page?P=' + str(p["no"]) + '" target="_blank" class="bl-link">BL</a>'
                '</div>'
            )

        deal_class = "good" if r["good_deal"] else "bad"
        deal_label = "\u2705 GOOD DEAL" if r["good_deal"] else "\U0001f6ab BAD DEAL \u2014 destroys more value than it creates"
        roi_class = "good-roi" if r["good_deal"] else "bad-roi"
        roi_sign = "+" if r["roi"] >= 0 else ""
        badge_type = "one" if len(r["missing"]) == 1 else "two"

        fp  = "{:.2f}".format(r["fig_price"])
        pv  = "{:.2f}".format(r["pulled_value"])
        mc  = "{:.2f}".format(r["missing_cost"])
        roi = "{:.2f}".format(abs(r["roi"]))

        rows += (
            '\n<div class="fig-card ' + deal_class + '" data-deal="' + deal_class + '" data-missing="' + str(len(r["missing"])) + '" data-roi="' + "{:.2f}".format(r["roi"]) + '">'
            '\n  <div class="deal-banner ' + deal_class + '">' + deal_label + '</div>'
            '\n  <div class="fig-hero">'
            '\n    <img src="https://img.bricklink.com/ItemImage/MN/0/' + r["fig_id"] + '.png" class="fig-img" onerror="this.src=\'https://img.bricklink.com/ItemImage/MN/0/default.png\'">'
            '\n    <div class="fig-info">'
            '\n      <div class="fig-id"><a href="https://www.bricklink.com/v2/catalog/catalogitem.page?M=' + r["fig_id"] + '" target="_blank">' + r["fig_id"] + '</a>'
            ' <span class="badge badge-' + badge_type + '">' + str(len(r["missing"])) + ' missing</span></div>'
            '\n      <div class="fig-math">'
            '\n        <span class="fig-price-val">$' + fp + '</span><span class="lbl"> sells for</span>'
            '\n        <span class="math-op">\u2212</span><span class="pulled-val">$' + pv + '</span><span class="lbl"> parts you pull</span>'
            '\n        <span class="math-op">\u2212</span><span class="missing-val">$' + mc + '</span><span class="lbl"> parts to buy</span>'
            '\n        <span class="math-op">=</span><span class="' + roi_class + '">' + roi_sign + '$' + roi + ' profit</span>'
            '\n      </div>'
            '\n    </div>'
            '\n  </div>'
            '\n  <div class="missing-section">'
            '\n    <div class="missing-label">Parts to buy:</div>'
            '\n    ' + parts_html +
            '\n  </div>'
            '\n</div>'
        )

    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Near Misses \u2014 Almost Buildable Minifigs</title>
<style>
  :root { --red:#c91a09; --yellow:#f5cd2f; --blue:#0055bf; --green:#237841; --bg:#f4f5f7; --card:#fff; --border:#dee2e6; --text:#212529; --muted:#6c757d; }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; background:var(--bg); color:var(--text); }
  .header { background:linear-gradient(135deg,#c91a09,#7f0d03); color:white; padding:28px 32px; text-align:center; }
  .header h1 { font-size:1.8rem; margin-bottom:6px; }
  .header p { opacity:0.8; font-size:0.9rem; }
  .summary-bar { display:flex; background:white; border-bottom:3px solid var(--yellow); flex-wrap:wrap; }
  .stat { text-align:center; padding:16px 28px; flex:1; min-width:100px; border-right:1px solid var(--border); cursor:pointer; }
  .stat:last-child { border-right:none; }
  .stat:hover { background:#f8f9fa; }
  .stat .num { font-size:1.6rem; font-weight:700; }
  .stat .lbl { font-size:0.7rem; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-top:2px; }
  .controls { max-width:920px; margin:20px auto 0; padding:0 16px; display:flex; gap:10px; flex-wrap:wrap; }
  .controls select, .controls input { padding:8px 12px; border:1px solid var(--border); border-radius:8px; font-size:0.9rem; background:white; }
  .controls input { flex:1; min-width:180px; }
  .container { max-width:920px; margin:16px auto 40px; padding:0 16px; }
  .fig-card { background:var(--card); border-radius:12px; margin-bottom:14px; border:1px solid var(--border); box-shadow:0 2px 6px rgba(0,0,0,.05); overflow:hidden; }
  .fig-card.bad { opacity:0.7; }
  .deal-banner { padding:6px 18px; font-size:0.8rem; font-weight:700; }
  .deal-banner.good { background:#e8f5e9; color:#1b5e20; border-bottom:1px solid #a5d6a7; }
  .deal-banner.bad { background:#ffebee; color:#b71c1c; border-bottom:1px solid #ef9a9a; }
  .fig-hero { display:flex; align-items:center; gap:14px; padding:14px 18px; }
  .fig-img { width:64px; height:64px; object-fit:contain; border-radius:8px; border:1px solid var(--border); background:#fafafa; flex-shrink:0; }
  .fig-info { flex:1; }
  .fig-id { font-weight:700; font-size:0.95rem; margin-bottom:6px; }
  .fig-id a { color:var(--blue); text-decoration:none; }
  .fig-id a:hover { text-decoration:underline; }
  .fig-math { display:flex; align-items:center; gap:6px; flex-wrap:wrap; font-size:0.85rem; }
  .fig-price-val { font-size:1.1rem; font-weight:700; color:var(--blue); }
  .pulled-val { font-weight:600; color:var(--text); }
  .missing-val { font-weight:600; color:var(--blue); }
  .math-op { color:#aaa; font-size:1.1rem; font-weight:700; margin:0 2px; }
  .lbl { color:var(--muted); }
  .good-roi { font-weight:700; font-size:1rem; color:var(--green); }
  .bad-roi { font-weight:700; font-size:1rem; color:var(--red); }
  .badge { display:inline-block; padding:2px 8px; border-radius:12px; font-size:0.72rem; font-weight:700; margin-left:6px; }
  .badge-one { background:#fff3e0; color:#e65100; border:1px solid #e65100; }
  .badge-two { background:#fce4ec; color:#880e4f; border:1px solid #880e4f; }
  .missing-section { border-top:1px solid var(--border); padding:12px 18px; background:#fffbf5; }
  .missing-label { font-size:0.72rem; text-transform:uppercase; letter-spacing:0.5px; color:var(--muted); margin-bottom:8px; font-weight:600; }
  .missing-part { display:flex; align-items:center; gap:8px; margin-bottom:6px; font-size:0.85rem; flex-wrap:wrap; }
  .part-no { font-weight:700; min-width:70px; }
  .part-name { color:var(--muted); flex:1; }
  .part-color { background:#eee; padding:2px 6px; border-radius:4px; font-size:0.75rem; }
  .part-qty { color:var(--muted); font-size:0.8rem; }
  .part-price { font-weight:600; color:var(--red); min-width:50px; }
  .bl-link { font-size:0.72rem; background:var(--blue); color:white; padding:2px 7px; border-radius:4px; text-decoration:none; }
  .bl-link:hover { background:#003d8f; }
  .no-results { text-align:center; padding:48px; color:var(--muted); }
  .footer { text-align:center; padding:24px; color:var(--muted); font-size:0.8rem; }
</style>
</head>
<body>
<div class="header">
  <h1>\U0001f3af Near Misses</h1>
  <p>Almost-buildable minifigs \u2014 missing only 1 or 2 parts \u2014 """ + date_str + """</p>
</div>
<div class="summary-bar">
  <div class="stat" onclick="filterDeal('')"><div class="num">""" + str(total) + """</div><div class="lbl">All Near Misses</div></div>
  <div class="stat" onclick="filterDeal('good')"><div class="num" style="color:var(--green)">""" + str(good_count) + """</div><div class="lbl">\u2705 Good Deals</div></div>
  <div class="stat" onclick="filterDeal('bad')"><div class="num" style="color:var(--red)">""" + str(bad_count) + """</div><div class="lbl">\U0001f6ab Bad Deals</div></div>
  <div class="stat" onclick="filterMissing('1')"><div class="num">""" + str(missing_1) + """</div><div class="lbl">Missing 1 Part</div></div>
  <div class="stat" onclick="filterMissing('2')"><div class="num">""" + str(missing_2) + """</div><div class="lbl">Missing 2 Parts</div></div>
</div>
<div class="controls">
  <input type="text" id="search" placeholder="Search by fig ID..." oninput="applyFilters()">
  <select id="dealFilter" onchange="applyFilters()">
    <option value="">All deals</option>
    <option value="good">Good deals only</option>
    <option value="bad">Bad deals only</option>
  </select>
  <select id="missingFilter" onchange="applyFilters()">
    <option value="">Any missing count</option>
    <option value="1">Missing 1 part</option>
    <option value="2">Missing 2 parts</option>
  </select>
  <select id="sortBy" onchange="applyFilters()">
    <option value="roi">Sort: Best ROI first</option>
    <option value="price">Sort: Fig value (highest)</option>
    <option value="missing">Sort: Fewest missing parts</option>
  </select>
</div>
<div class="container" id="container">
""" + rows + """
  <div class="no-results" id="no-results" style="display:none">No results match your filter.</div>
</div>
<div class="footer">
  """ + str(total) + """ near misses \u2014 """ + str(good_count) + """ good deals \u2014 """ + date_str + """<br>
  <small>ROI = fig sells for \u2212 parts you pull (your listed prices) \u2212 parts to buy (market price)</small>
</div>
<script>
const allCards = Array.from(document.querySelectorAll('.fig-card'));

function filterDeal(v) {
  document.getElementById('dealFilter').value = v;
  applyFilters();
}
function filterMissing(v) {
  document.getElementById('missingFilter').value = v;
  applyFilters();
}
function applyFilters() {
  const search = document.getElementById('search').value.toLowerCase();
  const deal = document.getElementById('dealFilter').value;
  const missing = document.getElementById('missingFilter').value;
  const sortBy = document.getElementById('sortBy').value;

  let visible = allCards.filter(c => {
    const matchD = !deal || c.dataset.deal === deal;
    const matchM = !missing || c.dataset.missing === missing;
    const matchS = !search || c.innerHTML.toLowerCase().includes(search);
    return matchD && matchM && matchS;
  });

  visible.sort((a, b) => {
    if (sortBy === 'roi') return parseFloat(b.dataset.roi) - parseFloat(a.dataset.roi);
    if (sortBy === 'price') return parseFloat(b.dataset.figPrice||0) - parseFloat(a.dataset.figPrice||0);
    if (sortBy === 'missing') return parseInt(a.dataset.missing) - parseInt(b.dataset.missing);
    return 0;
  });

  const container = document.getElementById('container');
  allCards.forEach(c => { c.style.display = 'none'; });
  visible.forEach(c => { c.style.display = 'block'; container.appendChild(c); });
  document.getElementById('no-results').style.display = visible.length === 0 ? 'block' : 'none';
}
</script>
</body>
</html>"""

    output_file.write_text(html, encoding="utf-8")


# ============================================================
# Main
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Find almost-complete minifigs (near misses)")
    parser.add_argument("parts_file", nargs="?", help="Path to torsos.xml")
    parser.add_argument("-i", "--inventory", help="Path to inventory.xml")
    parser.add_argument("-o", "--output", default="near_misses_report.html", help="Output HTML file")
    parser.add_argument("--max-missing", type=int, default=2, help="Max missing parts (default: 2)")
    parser.add_argument("--async", dest="use_async", action="store_true", help="Use async mode")
    parser.add_argument("-c", "--concurrency", type=int, default=5)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    if not args.parts_file:
        print("Usage: py near-misses.py torsos.xml -i inventory.xml --async -c 10")
        sys.exit(1)

    config = load_config()
    client = BrickLinkClient(config.bricklink)
    api = BrickLinkAPI(client)
    cache = DiskCache(config.cache)

    parts = parse_parts_xml(Path(args.parts_file))
    print(f"Loaded {len(parts)} parts from {args.parts_file}")

    print("\nSearching for candidate minifigs...")
    if args.use_async:
        from fig_finder.core.async_finder import run_async_search
        all_minifig_ids, minifig_parts_map = asyncio.run(
            run_async_search(parts, config.bricklink, cache, args.concurrency)
        )
    else:
        finder = MinifigFinder(api, cache)
        all_minifig_ids = finder.find_minifigs_for_parts(parts)
        minifig_parts_map = {mid: finder.get_minifig_parts(mid) for mid in all_minifig_ids}
    print(f"Found {len(all_minifig_ids)} candidate minifigs")

    print("\nLoading inventory...")
    if args.inventory:
        inventory = parse_inventory_xml_with_price(args.inventory)
    else:
        print("Fetching from BrickLink API...")
        raw_inv = api.get_inventory()
        inventory = [
            InventoryItem(part_no=item["item"]["no"], color_id=item.get("color_id"),
                          quantity=item.get("quantity", 1), remarks=item.get("remarks", "") or "")
            for item in raw_inv
        ]
        for item in inventory:
            item.my_price = 0.0
    print(f"Loaded {len(inventory)} inventory items")

    inventory_lookup = build_inventory_lookup(inventory)
    color_map = api.get_colors()

    print(f"\nChecking for near misses (missing <= {args.max_missing} parts)...")
    near_misses = []
    for i, mid in enumerate(all_minifig_ids):
        if i % 100 == 0:
            print(f"  {i}/{len(all_minifig_ids)}...")
        req_parts = minifig_parts_map.get(mid, [])
        if not req_parts:
            continue
        result = check_near_miss(req_parts, inventory_lookup, color_map, args.max_missing)
        if result:
            near_misses.append({"fig_id": mid, "missing": result["missing"],
                                 "have": result["have"], "fig_price": 0.0})

    print(f"\nFound {len(near_misses)} near misses")
    if not near_misses:
        print("No near misses found.")
        sys.exit(0)

    print("Fetching prices...")
    for i, r in enumerate(near_misses):
        if i % 20 == 0:
            print(f"  {i}/{len(near_misses)}...")
        r["fig_price"] = api.get_minifig_price(r["fig_id"])
        r["missing"] = enrich_missing_parts(r["missing"], api)
        time.sleep(0.1)

    # Compute ROI then sort best deals first
    for r in near_misses:
        compute_roi(r)
    near_misses.sort(key=lambda r: (0 if r["good_deal"] else 1, -r["roi"]))

    # HTML report
    output_file = Path(args.output)
    generate_report(near_misses, output_file)

    # Wanted list XML — only profitable figs
    good_deals = [r for r in near_misses if r["good_deal"]]
    wanted_file = output_file.parent / "near-misses-wanted.xml"
    if good_deals:
        wanted_file.write_text(build_wanted_list_xml(good_deals), encoding="utf-8")
        print(f"\nWanted list → {wanted_file}  ({len(good_deals)} profitable figs)")
        print(f"  Import on BrickLink: Want > Upload Wanted List")
    else:
        print("\nNo profitable near misses — no wanted list generated.")

    print(f"Done! {len(near_misses)} near misses, {len(good_deals)} good deals → {output_file}")

    if not args.no_browser:
        webbrowser.open(output_file.resolve().as_uri())


if __name__ == "__main__":
    main()
