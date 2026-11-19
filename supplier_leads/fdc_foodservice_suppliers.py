#!/usr/bin/env python3
"""
Build a list of food manufacturers that sell into foodservice distribution
(US Foods, Sysco, Gordon, PFG, ...) from the USDA FoodData Central
"Branded Foods" bulk dataset.

Why this data source works
--------------------------
Branded Foods is fed by GS1 GDSN data pools (1WorldSync etc. -- the same
item-data pipeline broadline distributors require from suppliers) and by
Label Insight. Each row carries the brand owner, GTIN, pack weight and,
for GDSN items, the trade channel the supplier published to (FOOD_SERVICE,
CHILD_NUTRITION_FOOD_PROGRAMS, GROCERY, ...).

Signals scored per brand owner
------------------------------
* fs_channel_items  - items published to the FOOD_SERVICE trade channel
* k12_items         - items published to CHILD_NUTRITION_FOOD_PROGRAMS
                      (K-12 foodservice, also sold through broadliners)
* bulk_pack_items   - foodservice-size packs: >= 6 lb net, or >= 1 gal
                      (#10 cans, 5-lb bags, gallon jugs, ...)
* gdsn_items        - items synced through GDSN (company already pushes
                      item data electronically to trading partners)

Tiers
-----
A  explicit foodservice / K-12 trade channel in GDSN
B  no channel tag, but >= MIN_BULK foodservice-size packs
C  everything else (dropped unless --all)

Distributor private labels (US Foods, Sysco, GFS, PFG, ...) are put on a
separate sheet: they are the customers, not the suppliers.

Usage
-----
    python3 fdc_foodservice_suppliers.py --download            # latest release
    python3 fdc_foodservice_suppliers.py --zip path/to/branded.zip
    python3 fdc_foodservice_suppliers.py --zip ... --out leads.xlsx --all
"""
from __future__ import annotations

import argparse
import collections
import csv
import io
import re
import sys
import unicodedata
import urllib.request
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

FDC_DOWNLOAD_PAGE = "https://fdc.nal.usda.gov/download-datasets/"
FDC_BASE = "https://fdc.nal.usda.gov"
MIN_BULK = 3

# Distributor / private-label brand owners (normalized-name substrings).
DISTRIBUTORS = [
    "US FOODS", "USFOODS", "U S FOODS", "SYSCO", "GORDON FOOD SERVICE",
    "PERFORMANCE FOOD GROUP", "PFG", "REINHART", "VISTAR", "CORE MARK",
    "CHEFS WAREHOUSE", "BEN E KEITH", "SHAMROCK FOODS", "CHENEY BROTHERS",
    "SAMS WEST", "COSTCO", "KIRKLAND", "RESTAURANT DEPOT", "JETRO",
    "UNITED NATURAL FOODS", "UNFI", "DOT FOODS", "MCLANE", "AMCON",
    "WALMART", "KROGER", "ALBERTSONS", "SAFEWAY", "TARGET", "ALDI",
    "TRADER JOE", "PUBLIX", "MEIJER", "H E B", "HEB", "WEGMANS",
    "FOOD LION", "HY VEE", "GIANT EAGLE", "DOLLAR GENERAL", "AHOLD",
    "TOPCO", "C S WHOLESALE", "ASSOCIATED WHOLESALE GROCERS",
    "WAL MART", "SMART AND FINAL", "SUPERVALU", "WAKEFERN", "WHOLE FOODS",
    "BJS WHOLESALE", "BJ S WHOLESALE", "HARRIS TEETER", "NASH FINCH",
    "H E BUTT", "TOPS MARKETS", "SAMS CLUB", "SCHNUCK", "GLENCOURT",
    "AMAZON", "SPROUTS", "WINCO", "SOUTHEASTERN GROCERS", "BI LO",
    "WINN DIXIE", "LIDL", "STOP AND SHOP", "GIANT FOOD", "HANNAFORD",
    "DELHAIZE", "FRED MEYER", "RALPHS", "VONS", "JEWEL", "SHOPRITE",
    "PRICE CHOPPER", "WEIS MARKETS", "INGLES", "RALEY", "STATER BROS",
    "SAVE MART", "BROOKSHIRE", "PIGGLY WIGGLY", "FAMILY DOLLAR",
    "DOLLAR TREE", "CVS", "WALGREEN", "RITE AID", "7 ELEVEN", "CASEY S",
    "SPARTANNASH", "SPARTAN STORES", "GROCERS SUPPLY", "BOZZUTO",
    "KEHE", "ASSOCIATED FOOD STORES", "AFFILIATED FOODS",
    "DEFENSE COMMISSARY", "FAREWAY", "SAMS", "BJS",
]

SUFFIXES = re.compile(
    r"\b(INCORPORATED|INC|LLC|L L C|LLP|LP|LTD|LIMITED|CO|COMPANY|CORP|"
    r"CORPORATION|THE|USA|US|NA|NORTH AMERICA|HOLDINGS|GROUP)\b"
)

# package_weight examples: "111 ONZ", "4 LBR", "46 OZA", "1.5 KGM",
# "16 oz/454 g", "1 GAL"
UNIT_TO_OZ = {
    "ONZ": 1, "OZ": 1, "LBR": 16, "LB": 16, "LBS": 16,
    "GRM": 1 / 28.3495, "G": 1 / 28.3495, "KGM": 35.274, "KG": 35.274,
}
UNIT_TO_FLOZ = {
    "OZA": 1, "FLOZ": 1, "FL OZ": 1, "GLL": 128, "GAL": 128,
    "LTR": 33.814, "L": 33.814, "MLT": 0.033814, "ML": 0.033814,
}
PW_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(FL OZ|[A-Z]+)", re.I)


def normalize(name: str) -> str:
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    name = re.sub(r"['\u2019`]", "", name)          # SAM'S -> SAMS
    name = re.sub(r"\d{8,}", " ", name)               # embedded GLN/GTIN codes
    n = re.sub(r"[^A-Z0-9 ]+", " ", name.upper().replace("&", " AND "))
    n = SUFFIXES.sub(" ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return ALIASES.get(n, n)


# Known mangled spellings in the source data -> canonical normalized name.
ALIASES = {
    "SOCIT DES PRODUITS NESTL S A": "SOCIETE DES PRODUITS NESTLE S A",
}


def is_distributor(norm: str) -> bool:
    padded = f" {norm} "
    return any(f" {d} " in padded for d in DISTRIBUTORS)


# Categories where a 1-gallon pack is a normal retail size (milk, juice...).
RETAIL_GALLON_CATS = re.compile(
    r"milk|water|soda|juice|drink|beverage|tea|coffee|ice cream", re.I)


# Pack weight in these categories is usually a retail multipack total
# (24 x 16.9 oz water) or a supplement tub, not a foodservice pack.
NO_BULK_CATS = re.compile(
    r"water|soda|drink|beverage|juice|energy|protein|sport|alcohol|"
    r"vitamin|supplement|kombucha", re.I)


def is_bulk(package_weight: str, category: str = "") -> bool:
    """Foodservice-size pack: >= 6 lb net (a #10 can is ~6.9 lb; 5-lb retail
    flour/sugar bags are excluded), > 1 gal, or exactly 1 gal outside
    retail-gallon categories (gallon sauces, dressings, oils)."""
    if NO_BULK_CATS.search(category):
        return False
    for qty, unit in PW_RE.findall(package_weight or ""):
        q, u = float(qty), unit.upper()
        if u in UNIT_TO_OZ and q * UNIT_TO_OZ[u] >= 96 - 0.5:
            return True
        if u in UNIT_TO_FLOZ:
            floz = q * UNIT_TO_FLOZ[u]
            if floz > 130 or (floz >= 127 and not RETAIL_GALLON_CATS.search(category)):
                return True
    return False


@dataclass
class Owner:
    names: collections.Counter = field(default_factory=collections.Counter)
    brands: collections.Counter = field(default_factory=collections.Counter)
    categories: collections.Counter = field(default_factory=collections.Counter)
    items: int = 0
    gdsn: int = 0
    fs: int = 0
    k12: int = 0
    bulk: int = 0
    latest: str = ""
    sample_gtin: str = ""
    fs_gtin: str = ""


def latest_zip_url() -> str:
    html = urllib.request.urlopen(FDC_DOWNLOAD_PAGE, timeout=60).read().decode()
    links = sorted(set(re.findall(
        r"/fdc-datasets/FoodData_Central_branded_food_csv_\d{4}-\d{2}-\d{2}\.zip",
        html)))
    if not links:
        sys.exit("No branded_food_csv release found on " + FDC_DOWNLOAD_PAGE)
    return FDC_BASE + links[-1]


def download(dest: Path) -> Path:
    url = latest_zip_url()
    print(f"Downloading {url}", file=sys.stderr)
    urllib.request.urlretrieve(url, dest)
    return dest


def aggregate(zip_path: Path) -> tuple[dict[str, Owner], str]:
    csv.field_size_limit(sys.maxsize)
    z = zipfile.ZipFile(zip_path)
    member = next(n for n in z.namelist() if n.endswith("/branded_food.csv"))
    release = member.split("/")[0]
    reader = csv.DictReader(io.TextIOWrapper(z.open(member), "utf-8", newline=""))
    owners: dict[str, Owner] = collections.defaultdict(Owner)
    for i, row in enumerate(reader, 1):
        if i % 500_000 == 0:
            print(f"  {i:,} rows", file=sys.stderr)
        if row["discontinued_date"]:
            continue
        raw = (row["brand_owner"] or "").strip()
        if not raw:
            continue
        o = owners[normalize(raw)]
        o.names[raw] += 1
        if row["brand_name"]:
            o.brands[row["brand_name"].strip()] += 1
        if row["branded_food_category"]:
            o.categories[row["branded_food_category"].strip()] += 1
        o.items += 1
        channel = row["trade_channel"] or ""
        if row["data_source"] == "NZGDSN":   # New Zealand/Australia market
            continue
        if row["data_source"] == "GDSN":
            o.gdsn += 1
        if "FOOD_SERVICE" in channel:
            o.fs += 1
            o.fs_gtin = o.fs_gtin or row["gtin_upc"]
        if "CHILD_NUTRITION" in channel:
            o.k12 += 1
        if is_bulk(row["package_weight"], row["branded_food_category"]):
            o.bulk += 1
        o.latest = max(o.latest, row["modified_date"] or "")
        o.sample_gtin = o.sample_gtin or row["gtin_upc"]
    return owners, release


def tier(o: Owner) -> str:
    if o.fs or o.k12:
        return "A"
    if o.bulk >= MIN_BULK:
        return "B"
    return "C"


def score(o: Owner) -> int:
    return 5 * o.fs + 3 * o.k12 + 2 * o.bulk + (10 if o.gdsn else 0)


def top(c: collections.Counter, n: int) -> str:
    return "; ".join(k for k, _ in c.most_common(n))


def to_rows(owners: dict[str, Owner], include_all: bool, gdsn_only: bool = False):
    prospects, distributors = [], []
    for norm, o in owners.items():
        t = tier(o)
        if gdsn_only:
            if not o.gdsn:
                continue
        elif t == "C" and not include_all:
            continue
        row = {
            "company": o.names.most_common(1)[0][0],
            "normalized_name": norm,
            "tier": t,
            "score": score(o),
            "active_items": o.items,
            "foodservice_channel_items": o.fs,
            "k12_channel_items": o.k12,
            "bulk_pack_items": o.bulk,
            "gdsn_items": o.gdsn,
            "gdsn_status": ("Syncs via GDSN" if o.gdsn else
                            "Not in GDSN (likely manual item setup)"),
            "top_brands": top(o.brands, 5),
            "top_categories": top(o.categories, 3),
            "latest_update": o.latest,
            "sample_gtin": o.fs_gtin or o.sample_gtin,
            "name_variants": top(o.names, 4),
        }
        (distributors if is_distributor(norm) else prospects).append(row)
    key = lambda r: (r["tier"], -r["score"], -r["active_items"])
    return sorted(prospects, key=key), sorted(distributors, key=key)


METHOD = [
    ("Source", "USDA FoodData Central - Branded Foods (bulk CSV)"),
    ("Release", None),
    ("Tier A", "Supplier published items to the FOOD_SERVICE or "
               "CHILD_NUTRITION_FOOD_PROGRAMS trade channel in GDSN."),
    ("Tier B", f"No channel tag, but >= {MIN_BULK} foodservice-size packs "
               "(>= 6 lb net, or >= 1 gallon outside milk/juice/beverage categories)."),
    ("Score", "5 x foodservice items + 3 x K-12 items + 2 x bulk-pack items "
              "+ 10 if any GDSN items."),
    ("gdsn_status", "'Not in GDSN' companies sell foodservice-size packs but do "
                    "not sync item data electronically - strongest prospects "
                    "for an EDI / item-setup solution."),
    ("Excluded", "Discontinued items. Distributor & retailer private labels are "
                 "on their own sheet."),
    ("Caveat", "Brand owner is the label owner, not always the manufacturer "
               "(co-packers are hidden). Being here means the company sells "
               "into foodservice/K-12, not specifically to US Foods."),
]


def write_xlsx(path: Path, prospects, distributors, release: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    sheets = [("Prospects", prospects), ("Distributor private label", distributors)]
    for idx, (title, rows) in enumerate(sheets):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = title
        if not rows:
            continue
        cols = list(rows[0])
        ws.append(cols)
        for r in rows:
            ws.append([r[c] for c in cols])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        widths = {"company": 40, "normalized_name": 30, "top_brands": 45,
                  "top_categories": 45, "name_variants": 45, "gdsn_status": 34}
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 14)
    ws = wb.create_sheet("Method")
    for k, v in METHOD:
        ws.append([k, release if v is None else v])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 110
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--zip", type=Path, help="FDC branded_food_csv zip")
    src.add_argument("--download", action="store_true",
                     help="download latest release into the current directory")
    ap.add_argument("--out", type=Path, default=Path("foodservice_suppliers.xlsx"))
    ap.add_argument("--all", action="store_true", help="include tier C")
    ap.add_argument("--gdsn", action="store_true",
                    help="every company with at least one GDSN item, any tier or category")
    args = ap.parse_args()

    zip_path = args.zip or download(Path("fdc_branded_food.zip"))
    owners, release = aggregate(zip_path)
    prospects, distributors = to_rows(owners, args.all, args.gdsn)
    write_xlsx(args.out, prospects, distributors, release)
    with open(args.out.with_suffix(".csv"), "w", newline="", encoding="utf-8") as f:
        if prospects:
            w = csv.DictWriter(f, fieldnames=list(prospects[0]))
            w.writeheader()
            w.writerows(prospects)
    counts = collections.Counter(r["tier"] for r in prospects)
    print(f"{release}: {len(owners):,} brand owners -> "
          f"{len(prospects):,} prospects {dict(counts)}, "
          f"{len(distributors):,} distributor/retailer labels. Wrote {args.out}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
