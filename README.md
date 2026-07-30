# SDS Parser

A rule-based (no AI/LLM) Python tool for bulk-extracting structured data
from Safety Data Sheets (SDS) — PDFs and DOCX files — across a wide range
of real-world vendor formats and layouts. Built and tested against 1,134
real SDS documents from multiple small-to-mid-size chemical distributors
and manufacturers (not just one company's clean, consistent template).

Deterministic and auditable by design: every extraction rule is a plain
regex or line-scan you can read and verify, not a model making a judgment
call. That matters for compliance work, where you need to know *why* a
field came out the way it did.

## Why this exists

Most SDS-parsing tools (there aren't many open-source ones) are built and
tested against one or two vendors' documents — often big-name chemical
suppliers with clean, consistent templates (Thermo Fisher, BASF, etc.).
That works fine until you need to process SDS documents from dozens of
smaller distributors, each with their own inconsistent formatting,
scanned pages, and vendor-specific quirks. This tool was built exactly
for that messier, more common situation.

## What it extracts

- All 16 GHS/OSHA-standard SDS sections as full text, with bullets/tables
  preserved
- Structured fields where possible: product name, company info, CAS
  numbers, hazard statements (H-codes), precautionary statements
  (P-codes), NFPA/HMIS ratings, dates
- Embedded images (pictograms, hazard diamonds), mapped to the section
  they actually appear in — not just dumped in a pile
- A per-file triage status (`OK`, `MISSING DATE`, `INCOMPLETE`, `PARSE
  ERROR`, etc.) so you can filter a batch of thousands down to exactly
  the files that need a human look, without eyeballing every row

## How it handles format variation

Real SDS documents use "SECTION 1", "SECTION – 1", "Section 1.",
"1. Identification" (no "SECTION" word at all), and inconsistent
capitalization — sometimes multiple styles within the *same* document.
The section-boundary detection is deliberately permissive about all of
this, while still protecting against the failure mode of a mid-sentence
cross-reference ("see Section 11 for symptoms...") being mistaken for a
real header.

OCR fallback (via `pytesseract` + `poppler`) kicks in automatically for:
- Scanned PDFs with no text layer at all
- PDFs with a broken font/encoding table that produces garbled text
  despite a normal-looking character count
- PDFs where the text layer is fine but the section header *banners*
  specifically were rendered as images with no accompanying text

Legacy `.doc` files are auto-converted via macOS's built-in `textutil`
before parsing.

## Installation

```bash
pip install pdfplumber python-docx Pillow pytesseract pdf2image
```

OCR support also requires two system binaries (not pip packages):
```bash
# macOS
brew install tesseract poppler

# Ubuntu/Debian
apt install tesseract-ocr poppler-utils
```

Without these, the parser still works — it just flags scanned/garbled
PDFs as needing manual entry instead of attempting OCR.

`.doc` auto-conversion uses `textutil`, which is **macOS-only**. On other
platforms, `.doc` files will need manual conversion to `.docx` first (the
tool tells you this clearly rather than failing silently).

## Usage

```bash
python3 sds_parser.py /path/to/folder/of/sds/files
```

Recursively scans the given folder (including nested subfolders — useful
if your files are organized by vendor). Produces:
- `SDS_Sections.csv` — one row per file, ~60 columns of extracted data
- `SDS_Needs_Review.csv` — just the rows flagged for manual attention
- `extracted_images/` — images pulled from each document, organized by
  source file and section

## Honest limitations

This is not a universal solution. Here's what it doesn't handle, based
on what actually showed up processing 1,134 real files:

- **Non-English documents.** The header detection is hardcoded to the
  word "SECTION" (and tolerant of OCR misspellings/kerning artifacts of
  *that specific word*). A Spanish-language SDS using "SECCIÓN" will not
  be recognized. No other language support exists yet.
- **Documents that aren't actually 16-section SDS documents.** Technical
  Data Sheets, Product Specification Sheets, and similar documents
  sometimes get mixed into SDS folders (often due to filename confusion,
  not a parser problem) and will correctly fail to parse as SDS — because
  they aren't one.
- **Severely broken table/column layouts.** A few real documents had
  content scrambled by how their original PDF or DOCX table structure
  extracted — words and numbers interleaved in ways no reasonable regex
  fix addresses. These need either better table-aware extraction (a real,
  unbuilt feature) or manual entry.
- **Images containing critical text with no OCR available.** If a
  section's actual content — not just its pictogram — is embedded as an
  image and OCR dependencies aren't installed, that content is lost.
- **OCR accuracy is inherently unpredictable.** The same document was
  observed rendering the same dash character three different ways
  (equals sign, em-dash, plain hyphen) across three different section
  headers via OCR. The parser tolerates a wide range of OCR artifacts,
  but OCR'd content should always be spot-checked against the source
  before being trusted for anything compliance-critical.
- **This tool cannot catch data integrity problems in your source
  files.** If a file is misnamed or contains the wrong product's content
  under a different filename, the parser will faithfully extract exactly
  what's in the file — it has no way to know the filename and content
  don't match.

## What "done" looks like in practice

Tested against a real batch of 1,134 SDS documents across dozens of
small-to-mid-size chemical distributors and manufacturers — not one
company's clean, consistent template:

- **683 files (60%) extracted with zero issues.**
- **236 more (21%) needed only a single manual date entry** — the
  document was otherwise 16/16 sections, fully extracted; it just didn't
  state its revision date in a machine-extractable way.
- **That's 919 of 1,134 files (81%) usable with zero or trivial manual
  work**, on a genuinely messy, diverse real-world batch.
- The remaining 215 files (19%) needed real review: unrecognized vendor
  formats, missing CAS numbers, a handful of parse errors, and a few
  documents that turned out not to be real 16-section SDS files at all
  (Technical Data Sheets, Specification Sheets, and similar, misfiled
  into the same folders under SDS-sounding filenames).

That's the realistic bar to expect — not 100%, but a large majority
handled automatically with the rest clearly triaged into specific,
explainable categories rather than silently wrong or dumped in one
undifferentiated pile.

## Contributing

If you run this against your own batch and find a real vendor format it
doesn't handle, PRs are welcome — ideally with the specific text pattern
that broke it (redact any proprietary/confidential content first).

## License

MIT
