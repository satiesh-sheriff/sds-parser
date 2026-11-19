# Foodservice supplier leads

`fdc_foodservice_suppliers.py` builds a list of food manufacturers that sell
into foodservice distribution (US Foods, Sysco, Gordon, PFG...) from the free
USDA FoodData Central **Branded Foods** bulk dataset. That dataset is fed by
GS1 GDSN (the item-data pipeline broadliners require from suppliers) and by
Label Insight.

```bash
pip install openpyxl
python3 fdc_foodservice_suppliers.py --download          # fetch latest release (~450 MB)
python3 fdc_foodservice_suppliers.py --zip FoodData_Central_branded_food_csv_YYYY-MM-DD.zip \
    --out output/foodservice_suppliers.xlsx [--all]

# Every US company that syncs item data through GDSN, any category:
python3 fdc_foodservice_suppliers.py --zip ... --gdsn --out output/gdsn_suppliers.xlsx
```

Runs in about 20 seconds once downloaded. Writes an `.xlsx` (Prospects, Distributor
private label, Method sheets) and a `.csv` of prospects.

| Tier | Meaning |
|------|---------|
| A | Supplier published items to the `FOOD_SERVICE` or `CHILD_NUTRITION_FOOD_PROGRAMS` trade channel in GDSN |
| B | No channel tag, but at least 3 foodservice-size packs (6 lb or more net, or a gallon or more outside drink categories) |
| C | Everything else (only with `--all`) |

`gdsn_status = Not in GDSN` marks companies with foodservice-size packs that
don't sync item data electronically. They're the most likely to need an
EDI / item-setup solution.

`--gdsn` keeps every brand owner with at least one US GDSN item, regardless of
tier or category (New Zealand/Australia `NZGDSN` rows are dropped). Tier is
still shown so you can sort.

Caveats: brand owner is the label owner, not always the manufacturer
(co-packers are hidden), and appearing here means the company sells into
foodservice, not specifically to US Foods.

## Retailers (US and worldwide)

`off_retailers.py` adds which stores carry each brand by matching USDA barcodes
to the crowd-sourced `stores` field in the Open Food Facts dump
(https://static.openfoodfacts.org/data/en.openfoodfacts.org.products.csv.gz, ~1.5 GB).

```bash
python3 off_retailers.py --usda-zip FoodData_Central_branded_food_csv_YYYY-MM-DD.zip \
    --off en.openfoodfacts.org.products.csv.gz \
    --brands output/all_brand_owners.csv --out output/brand_owners_with_retailers.xlsx
```

Matching runs strongest first: exact barcode, then the same GS1 company
prefix, then the same brand-owner name, then the same brand name. Prefix and
brand matches are used only when exactly one USDA owner has that prefix or
brand, and only for owners with no barcode match. `match_confidence` shows
high, medium or low.

It adds `us_retailers` (US chains with canonical names, e.g. Walmart or H-E-B),
`other_retailers` (stores worldwide), `countries`, retailer and country counts,
and Y/N flags for Walmart, Kroger, Costco, Target and Whole Foods. Coverage is
thin, about 2,100 of 34K brand owners: a blank means nobody logged the store,
not that the brand isn't sold there.
