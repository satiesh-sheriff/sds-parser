#!/usr/bin/env python3
"""
Add "which retailers (US and worldwide) carry this brand" to the brand-owner list built by
fdc_foodservice_suppliers.py, using Open Food Facts store data.

How it works
------------
1. Stream the Open Food Facts product dump and keep barcodes that have a
   `stores` value.
2. Stream the USDA branded_food.csv and match each product's gtin_upc to
   those barcodes (leading zeros ignored, so UPC-12 / EAN-13 / GTIN-14 line up).
3. Map store names to a canonical US retailer (Walmart, Kroger, H-E-B...)
   and count, per brand owner, how many of its products were seen at each.
4. Merge into the input CSV and write an enriched .xlsx/.csv.

Open Food Facts is crowd-sourced: a missing retailer means "nobody logged
it", not "not carried". Coverage is best for big national brands.

Usage
-----
    python3 off_retailers.py \
        --usda-zip FoodData_Central_branded_food_csv_YYYY-MM-DD.zip \
        --off en.openfoodfacts.org.products.csv.gz \
        --brands output/all_brand_owners.csv \
        --out output/brand_owners_with_retailers.xlsx

Download the Open Food Facts dump (~1.5 GB) from
https://static.openfoodfacts.org/data/en.openfoodfacts.org.products.csv.gz
"""
from __future__ import annotations

import argparse
import collections
import csv
import gzip
import io
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fdc_foodservice_suppliers import normalize  # noqa: E402

# Canonical US retailer -> regex matched against a lower-cased store name.
# Banners are kept separate from their parent (Ralphs != Kroger) because
# that is how buyers and EDI connections are usually organized.
US_RETAILERS = {
    "Walmart": r"^(wal-?l?mart|walmart supercenter|walmart neighborhood market|walmart\.com)$",
    "Sam's Club": r"^sams? ?club$|^sam's club$",
    "Kroger": r"^kroger",
    "Ralphs": r"^ralphs",
    "Fred Meyer": r"^fred meyer",
    "QFC": r"^qfc$",
    "Market Street": r"^market street$",
    "Star Market": r"^star market$",
    "Lucky": r"^lucky( supermarkets)?$",
    "King Soopers": r"^king soopers",
    "Fry's": r"^fry'?s( food)?$",
    "Smith's": r"^smith'?s( food)?",
    "Harris Teeter": r"^harris teeter",
    "Meijer": r"^meijer",
    "H-E-B": r"^(h-?e-?b|heb)( plus!?)?$|^central market$",
    "Costco": r"^costco",
    "Wegmans": r"^wegmans",
    "Giant Eagle": r"^giant eagle",
    "Aldi": r"^aldi( us| usa)?$",
    "Whole Foods": r"^whole foods?( market)?$",
    "Trader Joe's": r"^trader joe",
    "ShopRite": r"^shop ?rite",
    "Safeway": r"^safeway",
    "Albertsons": r"^albertson'?s",
    "Vons": r"^vons",
    "Jewel-Osco": r"^jewel",
    "Acme": r"^acme( markets)?$",
    "Market Basket": r"^market basket",
    "WinCo": r"^win ?co",
    "Hy-Vee": r"^hy-?vee",
    "Lidl": r"^lidl( us| usa)?$",
    "Amazon / Amazon Fresh": r"^amazon( fresh)?(\.com)?$",
    "7-Eleven": r"^7-?eleven",
    "Big Y": r"^big y",
    "Schnucks": r"^schnuck",
    "Stop & Shop": r"^stop (&|and) shop",
    "Giant Food": r"^giant( food)?$",
    "Food Lion": r"^food lion",
    "Hannaford": r"^hannaford",
    "Price Chopper": r"^price ?chopper",
    "Target": r"^target",
    "Publix": r"^publix",
    "Dollar Tree": r"^dollar tree",
    "Dollar General": r"^dollar general",
    "Family Dollar": r"^family dollar",
    "Piggly Wiggly": r"^piggly wiggly",
    "Walgreens": r"^walgreens",
    "CVS": r"^cvs",
    "Sprouts": r"^sprouts",
    "Winn-Dixie": r"^winn-?dixie",
    "BJ's Wholesale": r"^bj'?s",
    "Food 4 Less": r"^food 4 less",
    "Stater Bros": r"^stater",
    "Save Mart": r"^save ?mart",
    "Raley's": r"^raley",
    "Weis": r"^weis",
    "Ingles": r"^ingles",
    "Tops": r"^tops( markets)?$",
    "Fairway": r"^fairway",
    "Smart & Final": r"^smart (&|and) final",
    "Grocery Outlet": r"^grocery outlet",
    "Natural Grocers": r"^natural grocers",
    "Fresh Thyme": r"^fresh thyme",
    "The Fresh Market": r"^(the )?fresh market",
    "Food City": r"^food city",
    "Brookshire's": r"^brookshire",
    "Lowes Foods": r"^lowes foods",
    "Harmons": r"^harmons",
    "Stew Leonard's": r"^stew leonard",
    "Kwik Trip": r"^kwik trip",
    "Circle K": r"^circle k",
    "Wawa": r"^wawa",
    "Sheetz": r"^sheetz",
    "Casey's": r"^casey'?s",
}
_COMPILED = [(name, re.compile(rx)) for name, rx in US_RETAILERS.items()]


def canonical(store: str) -> str | None:
    s = re.sub(r"\s+", " ", store.strip().lower())
    for name, rx in _COMPILED:
        if rx.search(s):
            return name
    return None


def key(code: str) -> str:
    return code.strip().lstrip("0")


def _store_key(store: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", store.lower())


JUNK_STORES = {"everywhere", "online", "internet", "various", "unknown",
               "supermarket", "supermarkets", "grocerystore", "local", "na"}


def _country(tag: str) -> str:
    tag = tag.strip()
    if not tag.startswith("en:"):
        return ""
    return tag[3:].replace("-", " ").title()


def load_off(path: Path):
    """barcode -> (set of retailer keys, set of countries), plus key -> display name.

    US chains are mapped to a canonical name (Walmart, H-E-B...). Every other
    store is kept too, merged on a punctuation/case-insensitive key and shown
    under its most common spelling (Tesco, Carrefour, Woolworths...)."""
    out: dict[str, tuple[set[str], set[str]]] = {}
    spellings: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    with gzip.open(path, "rt", encoding="utf-8", errors="replace", newline="") as f:
        header = f.readline().rstrip("\n").split("\t")
        i_code, i_stores = header.index("code"), header.index("stores")
        i_ctry = header.index("countries_tags")
        for n, line in enumerate(f, 1):
            if n % 1_000_000 == 0:
                print(f"  OFF {n:,} rows", file=sys.stderr)
            parts = line.rstrip("\n").split("\t")
            if len(parts) <= max(i_stores, i_ctry) or not parts[i_stores]:
                continue
            code = parts[i_code].strip()
            if not code:
                continue
            stores = set()
            for raw in parts[i_stores].split(","):
                raw = re.sub(r"\s+", " ", raw).strip()
                if len(raw) < 2:
                    continue
                if _store_key(raw) in JUNK_STORES:
                    continue
                us = canonical(raw)
                k = "US:" + us if us else _store_key(raw)
                if not k or len(k) < 2:
                    continue
                spellings[k][us or raw] += 1
                stores.add(k)
            if not stores:
                continue
            ctry = {c for t in parts[i_ctry].split(",") if (c := _country(t))}
            st, cs = out.setdefault(key(code), (set(), set()))
            st.update(stores)
            cs.update(ctry)
    display = {k: c.most_common(1)[0][0] for k, c in spellings.items()}
    return out, display


def match_usda(usda_zip: Path, off):
    """normalized brand owner -> matched item count, Counter of retailers,
    Counter of countries."""
    csv.field_size_limit(sys.maxsize)
    z = zipfile.ZipFile(usda_zip)
    member = next(n for n in z.namelist() if n.endswith("/branded_food.csv"))
    reader = csv.DictReader(io.TextIOWrapper(z.open(member), "utf-8", newline=""))
    matched = collections.Counter()
    retailers: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    countries: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    seen: set[tuple[str, str]] = set()   # same GTIN appears in several FDC releases
    for row in reader:
        if row["discontinued_date"] or row["data_source"] == "NZGDSN":
            continue
        k = key(row["gtin_upc"] or "")
        hit = off.get(k)
        if not hit:
            continue
        norm = normalize(row["brand_owner"] or "")
        if (norm, k) in seen:
            continue
        seen.add((norm, k))
        matched[norm] += 1
        retailers[norm].update(hit[0])
        countries[norm].update(hit[1])
    return matched, retailers, countries


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--usda-zip", type=Path, required=True)
    ap.add_argument("--off", type=Path, required=True,
                    help="Open Food Facts en.openfoodfacts.org.products.csv.gz")
    ap.add_argument("--brands", type=Path, required=True,
                    help="CSV written by fdc_foodservice_suppliers.py (needs normalized_name)")
    ap.add_argument("--out", type=Path,
                    default=Path("output/brand_owners_with_retailers.xlsx"))
    args = ap.parse_args()

    print("Loading Open Food Facts store data...", file=sys.stderr)
    off, display = load_off(args.off)
    print(f"  {len(off):,} barcodes with at least one store", file=sys.stderr)
    print("Matching USDA products by barcode...", file=sys.stderr)
    matched, retailers, countries = match_usda(args.usda_zip, off)

    with open(args.brands, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        norm = r["normalized_name"]
        c = retailers.get(norm, collections.Counter())
        cc = countries.get(norm, collections.Counter())
        us = [(k, v) for k, v in c.most_common() if k.startswith("US:")]
        intl = [(k, v) for k, v in c.most_common() if not k.startswith("US:")]
        own = _store_key(norm)
        intl = [(k, v) for k, v in intl
                if not (own and (own in k or k in own) and len(k) > 3)]
        fmt = lambda pairs: "; ".join(f"{display[k]} ({v})" for k, v in pairs)
        r["off_matched_items"] = matched.get(norm, 0)
        r["retailer_count"] = len(c)
        r["us_retailer_count"] = len(us)
        r["intl_retailer_count"] = len(intl)
        r["us_retailers"] = fmt(us)
        r["other_retailers"] = fmt(intl[:40])
        r["countries"] = "; ".join(f"{k} ({v})" for k, v in cc.most_common())
        r["country_count"] = len(cc)
        for flag in ("Walmart", "Kroger", "Costco", "Target", "Whole Foods"):
            r[f"at_{flag.lower().replace(' ', '_')}"] = "Y" if c.get("US:" + flag) else ""
    rows.sort(key=lambda r: (-int(r["retailer_count"]), -int(r["off_matched_items"])))

    out_csv = args.out.with_suffix(".csv")
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    write_xlsx(args.out, rows, retailers, display, countries)
    hit = sum(1 for r in rows if int(r["retailer_count"]))
    us_hit = sum(1 for r in rows if int(r["us_retailer_count"]))
    print(f"{hit:,} of {len(rows):,} brand owners have at least one known retailer "
          f"({us_hit:,} with a US retailer). "
          f"Wrote {args.out} and {out_csv}", file=sys.stderr)


def write_xlsx(path: Path, rows, retailers, display, countries):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    numeric = {"score", "active_items", "foodservice_channel_items", "k12_channel_items",
               "bulk_pack_items", "gdsn_items", "off_matched_items", "retailer_count",
               "us_retailer_count", "intl_retailer_count", "country_count"}
    wb = Workbook()
    ws = wb.active
    ws.title = "Brand owners + retailers"
    cols = list(rows[0])
    ws.append(cols)
    for r in rows:
        ws.append([int(r[c]) if c in numeric and str(r[c]).lstrip("-").isdigit() else r[c]
                   for c in cols])
    widths = {"company": 40, "normalized_name": 28, "us_retailers": 60,
              "other_retailers": 60, "countries": 45,
              "top_brands": 40, "top_categories": 40, "name_variants": 40, "gdsn_status": 30}
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 13)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    # Retailer summary: how many brand owners were seen at each retailer.
    ws2 = wb.create_sheet("Retailer summary")
    ws2.append(["retailer", "us_chain", "brand_owners_seen"])
    totals = collections.Counter()
    for c in retailers.values():
        totals.update(c.keys())
    for k, n in totals.most_common():
        if n >= 2:
            ws2.append([display[k], "Y" if k.startswith("US:") else "", n])
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["C"].width = 18

    ws4 = wb.create_sheet("Country summary")
    ws4.append(["country", "brand_owners_seen"])
    ct = collections.Counter()
    for c in countries.values():
        ct.update(c.keys())
    for k, n in ct.most_common():
        ws4.append([k, n])
    ws4.column_dimensions["A"].width = 30

    ws3 = wb.create_sheet("Method")
    for k, v in [
        ("Source", "Open Food Facts product dump (crowd-sourced 'stores' field), "
                   "matched to USDA Branded Foods by barcode."),
        ("off_matched_items", "Products of this brand owner found in Open Food Facts "
                              "with at least one US retailer."),
        ("us_retailers", "US chains, canonical names (number of this owner's products seen there)."),
        ("other_retailers", "All other stores worldwide, merged case/punctuation-insensitively "
                            "(top 40)."),
        ("countries", "Countries where Open Food Facts users logged the product "
                      "(number of products)."),
        ("Caveat", "Crowd-sourced and incomplete: a blank means nobody logged it, "
                   "not that the retailer doesn't carry the brand. Use Kroger/Walmart "
                   "APIs to confirm for priority accounts."),
    ]:
        ws3.append([k, v])
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 110
    wb.save(path)


if __name__ == "__main__":
    main()
