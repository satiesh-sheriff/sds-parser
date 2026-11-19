#!/usr/bin/env python3
"""
Check whether brands are sold on Instacart, and at which chains, using the
Apify actor piotrv1001/instacart-scraper (cross-retailer search).

For each company it searches its top brand at one ZIP code and keeps only the
results whose Instacart brand matches one of the company's brands. Results are
appended to a JSONL file as they come in, so the script can be stopped and
resumed (for example next month, when free Apify credit resets).

Cost control
------------
* --max-items caps results per search (default 7, ~$0.035 at $0.005/result).
  Cross-retailer search returns about 7 results per chain, so 7 shows the top
  chain and 14 shows the top two.
* --budget stops the run once the account's monthly Apify usage reaches it.
  Usage is checked before every batch.

Needs an Apify token. In Claude Code cloud sessions it is injected by the
network proxy; elsewhere set APIFY_TOKEN.

Usage
-----
    python3 instacart_check.py --input candidates.csv --limit 110 \
        --zip 60614 --max-items 7 --budget 4.8 \
        --results output/instacart_results.jsonl --out output/instacart_check.xlsx
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fdc_foodservice_suppliers import normalize  # noqa: E402

API = "https://api.apify.com/v2"
ACTOR = "piotrv1001~instacart-scraper"


def api(method: str, path: str, body: dict | None = None, timeout: int = 360):
    req = urllib.request.Request(API + path, method=method,
                                 data=json.dumps(body).encode() if body is not None else None)
    req.add_header("Content-Type", "application/json")
    token = os.environ.get("APIFY_TOKEN")
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())


def monthly_usage() -> float:
    return float(api("GET", "/users/me/limits")["data"]["current"]["monthlyUsageUsd"])


def search(brand: str, zip_code: str, max_items: int) -> tuple[list[dict], float]:
    run = api("POST", f"/acts/{ACTOR}/runs?waitForFinish=300",
              {"searchQueries": [brand], "postalCode": zip_code, "maxItems": max_items})["data"]
    for _ in range(20):                       # waitForFinish caps at 300 s
        if run["status"] in ("SUCCEEDED", "FAILED", "ABORTED", "TIMED-OUT"):
            break
        time.sleep(15)
        run = api("GET", f"/actor-runs/{run['id']}")["data"]
    if run["status"] != "SUCCEEDED":
        raise RuntimeError(f"run {run['id']} {run['status']}")
    items = api("GET", f"/datasets/{run['defaultDatasetId']}/items?clean=true&format=json")
    return items, float(run.get("usageTotalUsd") or 0)


def brand_keys(row: dict) -> set[str]:
    names = [row["brand"]] + (row.get("top_brands") or "").split(";")
    return {k for n in names if len(k := normalize(n)) >= 3}


def is_match(item_brand: str, item_name: str, keys: set[str]) -> bool:
    b = normalize(item_brand or "")
    name = f" {normalize(item_name or '')} "
    for k in keys:
        if b and (b == k or b.startswith(k + " ") or k.startswith(b + " ")):
            return True
        if b and b.replace(" ", "") == k.replace(" ", ""):   # GOOD SENSE ~ GOODSENSE
            return True
        if len(k) >= 5 and f" {k} " in name and b and (b in k or k in b):
            return True
    return False


def load_done(path: Path) -> dict[str, dict]:
    done = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip():
                rec = json.loads(line)
                done[rec["normalized_name"]] = rec
    return done


def check(row: dict, zip_code: str, max_items: int) -> dict:
    items, _ = search(row["brand"].title(), zip_code, max_items)
    keys = brand_keys(row)
    hits = [i for i in items if is_match(i.get("brand"), i.get("name"), keys)]
    return {
        "normalized_name": row["normalized_name"],
        "company": row["company"],
        "searched_brand": row["brand"],
        "zip": zip_code,
        "results_returned": len(items),
        "matching_results": len(hits),
        "on_instacart": "Y" if hits else "N",
        "retailers": sorted({h["retailerSlug"] for h in hits}),
        "other_retailers_in_results": sorted({i["retailerSlug"] for i in items} -
                                             {h["retailerSlug"] for h in hits}),
        "sample_products": [h.get("name") for h in hits[:3]],
        "instacart_brands_seen": sorted({i.get("brand") or "" for i in items}),
    }


def write_outputs(done: dict[str, dict], inputs: list[dict], xlsx: Path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    by = {r["normalized_name"]: r for r in inputs}
    cols = ["company", "searched_brand", "on_instacart", "retailers", "matching_results",
            "results_returned", "sample_products", "instacart_brands_seen",
            "active_items", "top_categories", "zip"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Instacart check"
    ws.append(cols)
    recs = sorted(done.values(), key=lambda r: (r["on_instacart"] != "Y", -len(r["retailers"])))
    for r in recs:
        src = by.get(r["normalized_name"], {})
        ws.append([
            r["company"], r["searched_brand"], r["on_instacart"], "; ".join(r["retailers"]),
            r["matching_results"], r["results_returned"], "; ".join(r["sample_products"]),
            "; ".join(r["instacart_brands_seen"]), src.get("active_items"),
            src.get("top_categories"), r["zip"],
        ])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for col, w in zip("ABCDEFGHIJK", [38, 24, 12, 40, 10, 10, 60, 40, 10, 40, 8]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    ws2 = wb.create_sheet("Method")
    for k, v in [
        ("Source", "Apify actor piotrv1001/instacart-scraper, cross-retailer search."),
        ("Query", "The company's top brand, searched at one ZIP code."),
        ("on_instacart", "Y when at least one result's Instacart brand matches one of "
                         "the company's brands."),
        ("retailers", "Instacart retailer slugs where matching products appeared. "
                      "Search shows only the top chains for each query, so this list is "
                      "not complete."),
        ("Caveat", "N means not found in the top results at this ZIP, not that the "
                   "brand is absent from Instacart or stores."),
    ]:
        ws2.append([k, v])
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 100
    wb.save(xlsx)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", type=Path, required=True,
                    help="CSV with company, normalized_name, brand, top_brands")
    ap.add_argument("--limit", type=int, default=110)
    ap.add_argument("--zip", default="60614")
    ap.add_argument("--max-items", type=int, default=7)
    ap.add_argument("--budget", type=float, default=4.8,
                    help="stop when monthly Apify usage (USD) reaches this")
    ap.add_argument("--concurrency", type=int, default=5)
    ap.add_argument("--results", type=Path, default=Path("output/instacart_results.jsonl"))
    ap.add_argument("--out", type=Path, default=Path("output/instacart_check.xlsx"))
    args = ap.parse_args()

    with open(args.input, newline="", encoding="utf-8") as f:
        inputs = list(csv.DictReader(f))[: args.limit]
    done = load_done(args.results)
    todo = [r for r in inputs if r["normalized_name"] not in done]
    est = len(todo) * args.max_items * 0.005
    print(f"{len(done)} done, {len(todo)} to check, worst case ~${est:.2f}", file=sys.stderr)

    start_usage = monthly_usage()
    spent = 0.0          # Apify's usage figure lags, so also count locally
    with open(args.results, "a", encoding="utf-8") as out:
        for i in range(0, len(todo), args.concurrency):
            used = max(monthly_usage(), start_usage + spent)
            batch_cost = args.concurrency * args.max_items * 0.005
            if used + batch_cost > args.budget:
                print(f"Budget stop: ${used:.2f} used of ${args.budget:.2f}", file=sys.stderr)
                break
            batch = todo[i: i + args.concurrency]
            with cf.ThreadPoolExecutor(len(batch)) as pool:
                futs = {pool.submit(check, r, args.zip, args.max_items): r for r in batch}
                for fut in cf.as_completed(futs):
                    r = futs[fut]
                    try:
                        rec = fut.result()
                    except (RuntimeError, urllib.error.URLError, KeyError) as e:
                        print(f"  ! {r['brand']}: {e}", file=sys.stderr)
                        continue
                    spent += rec["results_returned"] * 0.005 + 0.00005
                    done[rec["normalized_name"]] = rec
                    out.write(json.dumps(rec) + "\n")
                    out.flush()
                    print(f"  {rec['on_instacart']} {r['brand'][:30]:30} "
                          f"{', '.join(rec['retailers'])}", file=sys.stderr)
    write_outputs(done, inputs, args.out)
    hits = sum(1 for r in done.values() if r["on_instacart"] == "Y")
    print(f"{hits} of {len(done)} checked brands found on Instacart. "
          f"This run spent ~${spent:.2f}; account usage ~${max(monthly_usage(), start_usage + spent):.2f}. "
          f"Wrote {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
