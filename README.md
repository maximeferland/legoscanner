<div align="center">

<!-- LEGO-style color bar -->
<img src="https://capsule-render.vercel.app/api?type=rect&color=gradient&customColorList=2,3,12,27&height=8&section=header" width="100%"/>

<br>

<!-- Header brick -->
<picture>
  <img alt="Fig Finder" src="https://img.shields.io/badge/%F0%9F%A7%B1_FIG_FINDER-Find_Minifigs_in_Your_Parts-c91a09?style=for-the-badge&labelColor=222"/>
</picture>

<br><br>

<a href="#-quick-start"><img src="https://img.shields.io/badge/Quick_Start-0055bf?style=flat-square&logo=rocket&logoColor=white" alt="Quick Start"/></a>&nbsp;
<a href="#-how-it-works"><img src="https://img.shields.io/badge/How_It_Works-237841?style=flat-square&logo=gear&logoColor=white" alt="How It Works"/></a>&nbsp;
<a href="#-setup"><img src="https://img.shields.io/badge/Setup-f5cd2f?style=flat-square&logo=wrench&logoColor=333" alt="Setup"/></a>&nbsp;
<a href="#-output"><img src="https://img.shields.io/badge/Output-c91a09?style=flat-square&logo=file&logoColor=white" alt="Output"/></a>&nbsp;
<a href="#-store-reports"><img src="https://img.shields.io/badge/Store_Reports-0055bf?style=flat-square&logo=bar-chart&logoColor=white" alt="Store Reports"/></a>

<br><br>

<img src="https://img.shields.io/badge/python-3.10+-0055bf?style=flat-square&logo=python&logoColor=white" alt="Python 3.10+"/>
<img src="https://img.shields.io/badge/node.js-18+-237841?style=flat-square&logo=node.js&logoColor=white" alt="Node.js 18+"/>
<img src="https://img.shields.io/badge/API-BrickLink-c91a09?style=flat-square" alt="BrickLink API"/>
<img src="https://img.shields.io/badge/license-MIT-237841?style=flat-square" alt="License"/>

---

**Fig Finder takes your BrickLink parts inventory and finds every complete<br>minifigure you can build — prioritized by market value so you build the best ones first.**

</div>

<br>

<!-- Stat cards like the HTML report -->
<table>
<tr>
<td align="center" width="33%">
<br>
<img src="https://img.shields.io/badge/SEARCHES-BrickLink_Catalog-0055bf?style=for-the-badge&labelColor=f0f0f0&logoColor=0055bf" alt=""/>
<br><br>
<sub>Finds every minifig that uses your parts</sub>
<br><br>
</td>
<td align="center" width="33%">
<br>
<img src="https://img.shields.io/badge/ALLOCATES-Smart_Part_Priority-237841?style=for-the-badge&labelColor=f0f0f0&logoColor=237841" alt=""/>
<br><br>
<sub>Most valuable minifigs get first pick of shared parts</sub>
<br><br>
</td>
<td align="center" width="33%">
<br>
<img src="https://img.shields.io/badge/REPORTS-Interactive_HTML-c91a09?style=for-the-badge&labelColor=f0f0f0&logoColor=c91a09" alt=""/>
<br><br>
<sub>Searchable cards with build checklists & progress tracking</sub>
<br><br>
</td>
</tr>
</table>

<br>

## 🧱 Quick Start

The fastest way to run it — async mode with 10 concurrent API connections:

```bash
python -m fig_finder torsos.xml -i inventory.xml --async -c 10 -o my_report.html
```

This will:
1. Read your parts from `torsos.xml`
2. Query BrickLink to find every minifig that uses those parts
3. Cross-reference your full `inventory.xml` for all required parts
4. Allocate parts intelligently (most valuable minifigs get priority)
5. Generate an interactive HTML report and open it in your browser

<br>

## 🔧 Setup

<details>
<summary><b>1 — Install dependencies</b></summary>
<br>

```bash
# Python (for Fig Finder)
pip install -r requirements.txt

# Node.js (for Store Reports)
npm install
```
</details>

<details>
<summary><b>2 — Get BrickLink API credentials</b></summary>
<br>

You need API access to your own BrickLink store:

1. Go to [BrickLink API Registration](https://www.bricklink.com/v2/api/register_consumer.page)
2. Register a new consumer (give it any name, like "Fig Finder")
3. Generate an access token — you'll get four values:
   - Consumer Key
   - Consumer Secret
   - Access Token
   - Token Secret

Copy `.env.example` to `.env` and fill them in:

```bash
cp .env.example .env
```

The `.env` file uses **two naming conventions** — one for the Python tools, one for the Node.js reports. Fill in the same credentials for both:

```ini
# Node.js report scripts (values MUST be wrapped in single quotes)
CONSUMER_KEY = 'your_consumer_key_here'
CONSUMER_SECRET = 'your_consumer_secret_here'
ACCESS_TOKEN = 'your_access_token_here'
TOKEN_SECRET = 'your_token_secret_here'
STORE_NAME = 'your_bricklink_username_here'

# Python fig_finder (same values, NO quotes)
BRICKLINK_CONSUMER_KEY=your_consumer_key_here
BRICKLINK_CONSUMER_SECRET=your_consumer_secret_here
BRICKLINK_ACCESS_TOKEN=your_access_token_here
BRICKLINK_TOKEN_SECRET=your_token_secret_here
```

> **`STORE_NAME`** is your BrickLink seller username. The report scripts use it to filter your sales from your purchases — without it, reports won't generate.
</details>

<details>
<summary><b>3 — Export your XML files from BrickLink</b></summary>
<br>

You need two XML files:

### `torsos.xml` — Parts to search

The list of parts Fig Finder will look up minifigures for. Typically **minifig torsos** (they're the most unique part), but can be anything.

**Export from BrickLink:** Inventory → filter to your target parts → Download as XML

```xml
<INVENTORY>
  <ITEM>
    <ITEMID>973c27</ITEMID>
    <ITEMTYPE>P</ITEMTYPE>
    <COLOR>11</COLOR>
    <QTY>5</QTY>
  </ITEM>
  <!-- ... -->
</INVENTORY>
```

**Why torsos?** A printed torso usually appears in only a few minifigs. Heads, legs, and accessories are shared across dozens. Starting from torsos gives the most targeted results.

### `inventory.xml` — Your full inventory

Your **complete** store inventory. Fig Finder checks this to verify you have ALL parts for each minifig (not just torsos — heads, legs, hair, accessories, everything).

**Export from BrickLink:** My Store → Inventory → Download → BrickLink XML format
</details>

<br>

## ⚙ How It Works

<table>
<tr>
<td width="40" align="center"><img src="https://img.shields.io/badge/1-c91a09?style=for-the-badge" alt="1"/></td>
<td><b>Parse</b> your parts XML to get part numbers and colors</td>
</tr>
<tr>
<td align="center"><img src="https://img.shields.io/badge/2-f5cd2f?style=for-the-badge" alt="2"/></td>
<td><b>Search</b> BrickLink API for every minifig containing each part</td>
</tr>
<tr>
<td align="center"><img src="https://img.shields.io/badge/3-0055bf?style=for-the-badge" alt="3"/></td>
<td><b>Fetch</b> complete parts lists for each candidate minifig</td>
</tr>
<tr>
<td align="center"><img src="https://img.shields.io/badge/4-237841?style=for-the-badge" alt="4"/></td>
<td><b>Check</b> your inventory for every required part in the right color & quantity</td>
</tr>
<tr>
<td align="center"><img src="https://img.shields.io/badge/5-c91a09?style=for-the-badge" alt="5"/></td>
<td><b>Allocate</b> shared parts using a greedy algorithm — highest value wins</td>
</tr>
<tr>
<td align="center"><img src="https://img.shields.io/badge/6-f5cd2f?style=for-the-badge" alt="6"/></td>
<td><b>Report</b> results as an interactive HTML page with build checklists</td>
</tr>
</table>

<br>

## 📋 All Options

```
python -m fig_finder [parts_file] [options]

Arguments:
  parts_file              Path to your BrickLink parts XML file
                          (runs in interactive mode if not provided)

Options:
  -i, --inventory FILE    Path to your inventory XML file
                          (fetches live from BrickLink API if not provided)
  -o, --output FILE       Output HTML report path (default: minifig_parts_list.html)
  --async                 Use async/concurrent processing (much faster)
  -c, --concurrency N     Max concurrent API requests with --async (default: 5)
  --no-browser            Don't auto-open the report in your browser
  --no-cache              Disable disk caching of API responses
  --clear-cache           Clear the cache and exit
  --version               Show version
```

### Examples

```bash
# Async mode, 10 concurrent connections (fastest)
python -m fig_finder torsos.xml -i inventory.xml --async -c 10

# Sync mode (slower but simpler)
python -m fig_finder torsos.xml -i inventory.xml

# Skip the local inventory file — fetch live from BrickLink
python -m fig_finder torsos.xml --async -c 10

# Interactive mode (prompts you for everything)
python -m fig_finder

# Custom output file, don't open browser
python -m fig_finder torsos.xml -i inventory.xml --async -c 10 -o results.html --no-browser

# Clear cached API responses
python -m fig_finder --clear-cache
```

<br>

## 📊 Output

The HTML report mirrors LEGO's signature color palette and includes:

<table>
<tr>
<td align="center">🃏</td>
<td><b>Minifig cards</b> with BrickLink images and market prices</td>
</tr>
<tr>
<td align="center">📦</td>
<td><b>Expandable parts lists</b> for each minifig with color and bin info</td>
</tr>
<tr>
<td align="center">✅</td>
<td><b>Build checkboxes</b> saved to your browser's localStorage</td>
</tr>
<tr>
<td align="center">🔍</td>
<td><b>Search & filter</b> by ID, part number, description, price range, status</td>
</tr>
<tr>
<td align="center">📤</td>
<td><b>Export built parts</b> as JSON for inventory removal</td>
</tr>
<tr>
<td align="center">📈</td>
<td><b>Progress bar</b> tracking your build completion and total value</td>
</tr>
</table>

<br>

## 🧩 BrickLink Upload Defaults (Optional)

BrickLink's XML upload page is [`invXML.asp`](https://www.bricklink.com/invXML.asp). Two options are easy to forget:

- **Concatenate inventories**: merges identical lots on upload
- **Remarks mode → New Remarks**: uses the newer remarks behavior on the upload page

### No extension installed (default)

Nothing in this repo changes BrickLink in your browser. If you don't install anything, just **set the two options manually** on the BrickLink page before uploading.

### Optional automation (Tampermonkey userscript)

If you want those two options set automatically every time you open the upload page:

1. Install a userscript manager (recommended: Tampermonkey)
2. Create a new script and paste the contents of `scripts/bricklink/bricklink-invxml-defaults.user.js`
3. Save, then open `invXML.asp` — the options should be prefilled

Safety notes:
- The script only runs on `invXML.asp`
- It does **not** submit the form — it only sets two UI controls

<br>

## 💡 Tips

- **Use `--async -c 10`** — concurrent requests make a huge difference against the BrickLink API bottleneck
- **Use a local inventory file (`-i`)** — downloading your full inventory from the API every run is slow
- **Caching is on by default** — responses are cached 24 hours in `~/.fig_finder/cache`. Use `--clear-cache` for fresh data
- **Start with torsos** — you'll get the most targeted minifig matches

<br>

## 📊 Store Reports

A full suite of standalone Node.js scripts that connect to the BrickLink API and generate interactive HTML reports. Every report uses the LEGO color palette and is a single self-contained HTML file — no build step, no server, just open it in your browser.

All scripts run from the repo root:

```bash
node scripts/reports/<script-name>.js
```

---

### Sales & Revenue

<details>
<summary><b>bricklink-sales-html-report.js</b> — Full Sales Dashboard</summary>
<br>

The big one. Generates a comprehensive sales dashboard with:
- Monthly revenue charts and trend lines
- Per-order breakdown with item details
- Category and theme analysis
- Buyer geographic distribution
- Filterable, searchable, sortable tables

```bash
node scripts/reports/bricklink-sales-html-report.js
```

Output: `reports/bricklink-sales-report.html`

> If you've issued refunds, add them to the `MANUAL_REFUNDS` object in the script — the BrickLink API doesn't expose refund data, so they need to be tracked manually.
</details>

<details>
<summary><b>monthly-revenue-report.js</b> — Monthly Revenue Bar Chart</summary>
<br>

Clean monthly revenue bar chart showing your store's revenue over time. Uses `subtotal` (items only, no shipping/tax) for accurate revenue tracking.

```bash
node scripts/reports/monthly-revenue-report.js
```

Output: `reports/monthly-revenue-report.html`
</details>

<details>
<summary><b>items-sold-report.js</b> — Per-Item Sales Detail</summary>
<br>

Shows every item sold in a given month with individual prices, quantities, and buyer info.

```bash
# Current month
node scripts/reports/items-sold-report.js

# Specific month
node scripts/reports/items-sold-report.js 3 2026

# All time
node scripts/reports/items-sold-report.js all
```

Output: `reports/items-sold-report.html`
</details>

<details>
<summary><b>best-sellers-report.js</b> — Top Sellers by Volume & Velocity</summary>
<br>

Identifies your best-selling items ranked by quantity sold and sales velocity (units per month).

```bash
# Current month
node scripts/reports/best-sellers-report.js

# Specific month
node scripts/reports/best-sellers-report.js 3 2026

# All time
node scripts/reports/best-sellers-report.js all
```

Output: `reports/best-sellers-report.html`
</details>

---

### Inventory Analysis

<details>
<summary><b>inventory-age-report.js</b> — How Long Items Have Been Listed</summary>
<br>

Shows how long each item has been sitting in your store. Highlights stale inventory that might need a price drop or removal.

```bash
node scripts/reports/inventory-age-report.js
```

Output: `reports/inventory-age-report.html`
</details>

<details>
<summary><b>time-to-sell-report.js</b> — Days from Listing to Sale</summary>
<br>

Analyzes how quickly items sell after being listed. Useful for understanding which categories or price points move fastest.

```bash
# Current month
node scripts/reports/time-to-sell-report.js

# Specific month
node scripts/reports/time-to-sell-report.js 3 2026
```

Output: `reports/time-to-sell-report.html`
</details>

<details>
<summary><b>build-or-not-report.js</b> — Build Minifigs or Sell Parts?</summary>
<br>

For each potential minifig you can build from parts, compares the assembled minifig value against the total value of selling the individual parts. Helps you decide whether it's worth building.

```bash
node scripts/reports/build-or-not-report.js
```

Output: `reports/build-or-not-to-build.html`

> Requires a `quick-sell-data.json` file (generated by the quick-sell analysis below).
</details>

---

### Pricing Tools

<details>
<summary><b>quick-sell-analysis.js</b> — Market-Based Pricing Analysis</summary>
<br>

Pulls current market data for all your minifig inventory and calculates competitive quick-sell prices using this algorithm:
- Start at 95% of average recent sold price
- If the cheapest current listing is lower, match it
- Uses median sold price as a floor to prevent bulk sales from dragging prices too low
- Items with no sales data: match the cheapest current listing

```bash
node scripts/reports/quick-sell-analysis.js
```

Output: `quick-sell-data.json` (used by other reports)
</details>

<details>
<summary><b>generate-quicksell-report.js</b> — Visual Quick-Sell Report</summary>
<br>

Transforms the `quick-sell-data.json` into a tiered HTML report showing recommended price changes grouped by urgency.

```bash
node scripts/reports/generate-quicksell-report.js
```

Output: `reports/quicksell-report.html`
</details>

<details>
<summary><b>pricing-update.js</b> — Bulk Price Updates (with safety net)</summary>
<br>

Applies quick-sell pricing to your store in a careful pipeline. **Never skip steps** — always backup first.

```bash
# Step 1: Create a backup of current prices
node scripts/reports/pricing-update.js backup

# Step 2: Preview changes without applying them
node scripts/reports/pricing-update.js dry-run

# Step 3: Apply the changes
node scripts/reports/pricing-update.js execute

# Step 4 (if needed): Roll back to the backup
node scripts/reports/pricing-update.js rollback price_backup_2026-03-05.xml
```

Output: Backup XML in `reports/`, update results JSON, and an HTML summary report.

> **Always run `backup` before `execute`.** The rollback command can restore your prices from the backup XML if anything goes wrong.
</details>

---

### Utility Scripts

<details>
<summary><b>capture-inventory-snapshot.js</b> — Periodic Inventory Snapshots</summary>
<br>

Takes a point-in-time snapshot of your store inventory and appends it to a history file. Useful for tracking inventory changes over time.

```bash
node scripts/tools/capture-inventory-snapshot.js
```

Output: Appends to `inventory-history.json`
</details>

<details>
<summary><b>fetch-categories.js</b> — Download BrickLink Categories</summary>
<br>

Downloads the full BrickLink category list for reference. Useful if you need category IDs for filtering.

```bash
node scripts/tools/fetch-categories.js
```
</details>

<br>

## 🗂 Project Structure

```
fig_finder/              Python minifig finder package
  api/                   BrickLink API client (sync + async)
  cache/                 Disk-based response caching
  core/                  Minifig finding, inventory checking, part allocation
  inventory/             Tools for removing built parts from inventory
  parsers/               XML file parsers
  report/                HTML report generation (Jinja2 templates)
scripts/
  reports/               Node.js report generators (10 scripts)
  tools/                 Utility scripts (snapshots, categories)
tests/                   Python test suite
.env.example             Template for API credentials (Python + Node.js)
package.json             Node.js dependencies
requirements.txt         Python dependencies
```

<br>

## 🔥 Troubleshooting

| Problem | Fix |
|---|---|
| `Missing required environment variable` | Your `.env` file is missing or incomplete — set all four API keys |
| `STORE_NAME not set` | Add `STORE_NAME = 'YourBrickLinkUsername'` to your `.env` file |
| `No parts found in file` | Check XML format: needs `<INVENTORY>` root with `<ITEM>` children |
| `No minifigures found` | Your parts might not appear in any known minifigs (e.g. regular bricks) |
| Reports show no sales data | Make sure `STORE_NAME` exactly matches your BrickLink seller username (case-sensitive) |
| `Cannot find module 'dotenv'` | Run `npm install` from the repo root |
| Slow performance | Use `--async -c 10` and make sure caching is enabled |
| 429 / rate limiting errors | Reduce concurrency: `--async -c 3` — built-in retry with backoff handles the rest |

<br>

---

<div align="center">
<sub>Built with the same colors as the bricks 🧱</sub>
<br>
<img src="https://capsule-render.vercel.app/api?type=rect&color=gradient&customColorList=2,3,12,27&height=4&section=footer" width="100%"/>
</div>
