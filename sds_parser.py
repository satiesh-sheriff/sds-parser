"""
Rule-based SDS parser — section-focused extraction.

No AI, no field-level regex extraction. Purpose: capture the full content of
each of the 16 SDS sections as faithfully as possible, including:
  - bullet / numbered list structure (DOCX)
  - table content (DOCX) — previously silently dropped
  - embedded images, mapped to the section they appear in

Output: one row per file in SDS_Sections.csv with section_N_text and
section_N_images columns. Extracted images are saved to disk under
extracted_images/<filename_stem>/ and referenced by relative path.
"""
import re, csv, sys, subprocess, tempfile, shutil
from pathlib import Path

import pdfplumber
from docx import Document as DocxDoc
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

try:
    from PIL import Image
    import io
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

try:
    import pytesseract
    from pdf2image import convert_from_path
    # pytesseract needs the actual tesseract binary installed, not just the
    # Python wrapper — confirm it's actually callable, not just importable.
    pytesseract.get_tesseract_version()
    HAVE_OCR = True
except Exception:
    HAVE_OCR = False

FIELDS = (
    ["filename", "file_path", "relative_path", "file_link", "file_type",
     "source_folder", "format", "is_scanned", "ocr_used", "total_pages",
     "extraction_status", "extraction_notes"]
    # Dates — needed to pick the winning SDS per CAS group (newest wins)
    + ["sds_date", "sds_revision_date", "print_date"]
    # Section 1 — Identification, broken into fields for template formatting
    + ["product_name", "product_synonyms", "product_uses",
       "company_name", "company_address", "company_city_state_zip",
       "company_phone", "emergency_phone"]
    # Section 2 — Hazards Identification, broken into fields
    + ["signal_word", "hazard_statements", "precautionary_statements",
       "nfpa_health", "nfpa_flammability", "nfpa_reactivity",
       "hmis_health", "hmis_flammability", "hmis_reactivity"]
    # CAS — needed for winner-selection grouping (Section 3, best-effort)
    + ["cas_number_primary", "all_cas_numbers"]
    + [f"section_{n}_text" for n in range(1, 17)]
    + [f"section_{n}_images" for n in range(1, 17)]
)

DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ── Format detection (unchanged logic) ─────────────────────────────────────
def try_convert_doc_to_docx(path):
    """
    Attempts to auto-convert a legacy Word 97-2003 (.doc) file to modern
    .docx using macOS's built-in `textutil` command-line tool — no extra
    install needed on a Mac, unlike LibreOffice/antiword. Returns the path
    to a converted .docx in a temp directory, or None if conversion wasn't
    possible (not running on macOS, textutil missing, or the file is
    corrupt/unreadable) — caller falls back to the manual-conversion message.
    """
    try:
        tmp_dir = Path(tempfile.mkdtemp(prefix="doc2docx_"))
        out_path = tmp_dir / (path.stem + ".docx")
        result = subprocess.run(
            ["textutil", "-convert", "docx", str(path), "-output", str(out_path)],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0 and out_path.exists():
            return out_path, tmp_dir
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass
    return None, None


def detect_format(text):
    if "Americhem Sales Company" in text: return "AMERICHEM"
    if "Buckley Oil" in text:             return "BUCKLEY"
    if "Dakota Distributing" in text:     return "DAKOTA"
    if "AllChem Industries" in text:      return "ALLCHEM"
    if "Phillips 66" in text:             return "PHILLIPS66"
    # These vendors don't have their own field-extraction functions (yet) —
    # they fall through to the generic Section 1/2 parsers — but labeling
    # them distinctly (rather than lumping into UNKNOWN) means files that
    # extract cleanly aren't needlessly flagged for manual review, and
    # gives a concrete list of formats worth building dedicated parsers for
    # if they show up often enough. Matched on text actually present in the
    # document body, not logo images, which don't have a text layer.
    if "Gage Products Company" in text:  return "GAGE"
    if "Cellmark Basic Materials" in text: return "CELLMARK"
    if "Thermo Fisher Scientific" in text or "Alfa Aesar" in text: return "THERMOFISHER"
    if "BASF CORPORATION" in text:       return "BASF"
    if "12700 Park Central Drive" in text: return "ALONUSA"
    if "UMCO Umwelt Consult" in text:    return "BOTTCHER"
    # Fixed: previously only matched a period ("SECTION 1."), so vendors
    # using "SECTION 1:" (colon) — Gage/AlonUSA/CellMark/Böttcher above all
    # do this — fell through to UNKNOWN even though section-splitting
    # itself (which already handled colons) worked fine.
    if re.search(r"SECTION\s+\d+[.:]", text): return "CONE"
    return "UNKNOWN"


# Cross-checks the vendor subfolder a file was found in (e.g. "Cone",
# "Americhem") against the content-detected format. A mismatch is a strong
# signal the file is either misfiled in the wrong folder or format
# detection got it wrong — either way, worth a human look rather than
# silently trusting content-based detection alone.
#
# Dakota is deliberately excluded: per the team's own account (SDS project
# planning call), Dakota's folder legitimately contains other vendors'
# native SDS documents for products they resell — a Cone- or
# Buckley-format file showing up there is expected structure, not a
# misfile. Flagging it created ~180 false positives in the first full
# batch run with no real signal behind them.
_FOLDER_FORMAT_HINTS = {
    "cone": ("CONE",),
    "americhem": ("AMERICHEM", "BUCKLEY"),
    "allchem": ("ALLCHEM",),
    "phillips": ("PHILLIPS66",),
}

def _folder_format_mismatch(source_folder, fmt):
    folder_lower = (source_folder or "").lower()
    for hint, expected_fmts in _FOLDER_FORMAT_HINTS.items():
        if hint in folder_lower:
            return fmt not in expected_fmts
    return False  # folder name doesn't match any known vendor hint — nothing to check


# ══════════════════════════════════════════════════════════════════════════
# SECTION 1 & 2 FIELD EXTRACTION — company-specific
#
# The raw section_N_text capture above is format-agnostic (same code path
# for every company). This block is the one place formats genuinely diverge:
# each company labels identification/hazard fields differently, so each gets
# its own function. Shared GHS conventions (signal word, H-codes, P-codes,
# NFPA/HMIS diamonds) are standardized enough across formats to extract with
# one common helper, called from each company's section-2 function.
# ══════════════════════════════════════════════════════════════════════════

def _f(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""

def _fal(label, text, maxc=300):
    """Grab text after a label up to the next all-caps line or end of block."""
    pat = label + r"\s*[:\-]?\s*(.{1," + str(maxc) + r"}?)(?=\n[A-Z]|\Z)"
    m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
    if not m or m.group(1) is None:
        return ""
    return re.sub(r"\s+", " ", m.group(1)).strip()

def _cv(s):
    return re.sub(r"\s+", " ", s).strip()


# ---- Shared GHS hazard extraction (Section 2), used by every format ------
def parse_hazards_common(text):
    r = {}
    r["signal_word"] = _f(r"\b(DANGER|WARNING)\b", text)
    h_codes = re.findall(r"(H\d{3}[A-Z]?)\s+([A-Z][^H\n]{5,80}?)(?=H\d|\n|$)", text)
    if h_codes:
        r["hazard_statements"] = "; ".join(f"{c} {d.strip()}" for c, d in h_codes[:20])
    p_codes = re.findall(r"(P\d{3}(?:\+P\d{3})*)", text)
    if p_codes:
        r["precautionary_statements"] = "; ".join(dict.fromkeys(p_codes[:30]))
    r["nfpa_health"]       = _f(r"Health\s+(\d)\b", text)
    r["nfpa_flammability"] = _f(r"Flammabilit[yi]\s+(\d)\b", text)
    r["nfpa_reactivity"]   = _f(r"Reactivity\s+(\d)\b", text)
    hmis = re.findall(r"HMIS.*?Health\s+(\d).*?Flammabilit[yi]\s+(\d).*?Reactivity\s+(\d)", text, re.DOTALL)
    if hmis:
        r["hmis_health"], r["hmis_flammability"], r["hmis_reactivity"] = hmis[0]
    return {k: _cv(v) for k, v in r.items() if v}


# ---- NFPA/HMIS fallback — Cone/Dakota often place the diamond in Section 16 --
def parse_nfpa_fallback_section16(s16_text):
    r = {}
    nh = re.search(r"HEALTH\s*\(NFPA\)[:\s]*(\d)", s16_text, re.IGNORECASE)
    nf = re.search(r"FLAMMABILITY[:\s]*(\d)", s16_text, re.IGNORECASE)
    nr = re.search(r"PHYSICAL HAZARD[:\s]*(\d)", s16_text, re.IGNORECASE)
    hh = re.search(r"HEALTH\s*\(HMIS\)[:\s]*(\d)", s16_text, re.IGNORECASE)
    if nh: r["nfpa_health"] = nh.group(1)
    if nf: r["nfpa_flammability"] = nf.group(1)
    if nr: r["nfpa_reactivity"] = nr.group(1)
    if hh: r["hmis_health"] = hh.group(1)
    return r


# ── AMERICHEM (also covers Buckley — shares the same tabular header layout) ─
def parse_section1_americhem(text):
    r = {}
    r["product_name"]    = _f(r"Product Name\s+(.+?)(?:\s{3,}|Item|\n)", text)
    r["product_uses"]    = _f(r"Product Use\s+(.+?)(?:\n)", text)
    r["company_name"]    = _f(r"Company Name\s+(.+?)(?:\s+Office\b|\n)", text)
    r["company_address"] = _f(r"Address\s+(.+?)(?:\n)", text)
    r["company_phone"]   = _f(r"Office\s+([\d\-\(\)\s]+)", text)
    r["emergency_phone"] = _f(r"EMERGENCY TELEPHONE NUMBER\s+(.+?)(?:\n)", text)
    return {k: _cv(v) for k, v in r.items() if v}

def parse_section2_americhem(text, full_text_for_nfpa_fallback=""):
    r = parse_hazards_common(text)
    # Americhem sometimes puts the NFPA diamond in Section 5 instead of 2
    if not r.get("nfpa_health") and full_text_for_nfpa_fallback:
        r2 = parse_hazards_common(full_text_for_nfpa_fallback)
        for k in ["nfpa_health", "nfpa_flammability", "nfpa_reactivity",
                  "hmis_health", "hmis_flammability", "hmis_reactivity"]:
            if r2.get(k) and not r.get(k):
                r[k] = r2[k]
    return r

def parse_dates_americhem(text):
    r = {}
    r["sds_date"]   = _f(r"Revision Date\s+([\d/]+)", text)
    r["print_date"] = _f(r"Print Date\s+([\d/]+)", text)
    r["sds_revision_date"] = r.get("sds_date", "")
    return {k: v for k, v in r.items() if v}


# ── DAKOTA ───────────────────────────────────────────────────────────────────
def parse_section1_dakota(text):
    r = {}
    r["product_name"]     = _f(r"PRODUCT IDENTITY:\s*(.+?)(?:\n|SDS)", text)
    r["product_synonyms"] = _f(r"PRODUCT SYNONYMS:\s*(.+?)(?:\n)", text)
    r["product_uses"]     = _f(r"PRODUCT USES:\s*(.+?)(?:\n)", text)
    r["company_name"]     = _f(r"COMPANY IDENTITY:\s*(.+?)(?:\n)", text)
    r["company_address"]  = _f(r"COMPANY ADDRESS:\s*(.+?)(?:\n)", text)
    r["company_city_state_zip"] = _f(r"COMPANY CITY:\s*(.+?)(?:\n)", text)
    r["company_phone"]    = _f(r"COMPANY PHONE:\s*(.+?)(?:\n)", text)
    r["emergency_phone"]  = _f(r"EMERGENCY PHONES?:\s*(.+?)(?:\n)", text)
    return {k: _cv(v) for k, v in r.items() if v}

def parse_section2_dakota(text):
    return parse_hazards_common(text)

def parse_dates_dakota(text):
    r = {}
    r["sds_date"] = _f(r"SDS DATE:\s*([\d/]+)", text)
    r["sds_revision_date"] = r.get("sds_date", "")
    return {k: v for k, v in r.items() if v}


# ── CONE ─────────────────────────────────────────────────────────────────────
def parse_section1_cone(text):
    r = {}
    r["product_name"]     = _f(r"PRODUCT IDENTITY:\s*(.+?)(?:\n|SYNONYMS)", text)
    r["product_synonyms"] = _f(r"SYNONYMS?:\s*(.+?)(?:\n)", text)
    r["product_uses"]     = _f(r"PRODUCT USES?:\s*(.+?)(?:\n)", text)
    r["company_name"]     = _f(r"COMPANY IDENTITY:\s*(.+?)(?:\n)", text)
    r["company_address"]  = _f(r"COMPANY ADDRESS:\s*(.+?)(?:\n)", text)
    r["company_phone"]    = _f(r"COMPANY PHONE:\s*(.+?)(?:\n)", text)
    r["emergency_phone"]  = _f(r"EMERGENCY PHONES?:\s*(.+?)(?:\n)", text)
    return {k: _cv(v) for k, v in r.items() if v}

def parse_section2_cone(text):
    return parse_hazards_common(text)

def parse_dates_cone(text):
    r = {}
    r["sds_date"] = _f(r"SDS DATE:\s*([\d/]+)", text)
    r["sds_revision_date"] = r.get("sds_date", "")
    return {k: v for k, v in r.items() if v}


# ── PHILLIPS66 ───────────────────────────────────────────────────────────────
def parse_section1_phillips66(text):
    r = {}
    r["product_name"]     = _f(r"Product Identifier:\s*(.+?)(?:\n)", text)
    r["product_synonyms"] = _f(r"Other means of identification:\s*(.+?)(?:\n)", text)
    r["product_uses"]     = _f(r"Intended Use:\s*(.+?)(?:\n)", text)
    r["company_name"]     = _f(r"Manufacturer:\s*\n(.+?)(?:\n)", text)
    r["emergency_phone"]  = _f(r"Chemtrec:\s*(.+?)(?:\n|\()", text)
    return {k: _cv(v) for k, v in r.items() if v}

def parse_section2_phillips66(text):
    return parse_hazards_common(text)

def parse_dates_phillips66(text):
    r = {}
    r["sds_date"] = _f(r"Date of Issue:\s*([\d\-A-Za-z]+)", text)
    r["sds_revision_date"] = r.get("sds_date", "")
    return {k: v for k, v in r.items() if v}


# ── Generic fallback (ALLCHEM / UNKNOWN) — best effort only ────────────────
def parse_section1_generic(text):
    r = {}
    r["product_name"] = (
        _f(r"Product\s*(?:Name|Identity|Identifier)\s*[:\-]?\s*(.+?)(?:\n)", text)
    )
    r["company_name"] = (
        _f(r"Company\s*(?:Name|Identity)\s*[:\-]?\s*(.+?)(?:\n)", text)
    )
    return {k: _cv(v) for k, v in r.items() if v}

def parse_section2_generic(text):
    return parse_hazards_common(text)

_MONTH_NAMES = (r"(?:January|February|March|April|May|June|July|August|"
                r"September|October|November|December)")
_DATE_SHAPE = (r"(\d{1,4}[/\-]\d{1,2}[/\-]\d{1,4}|"
               rf"{_MONTH_NAMES}\s+\d{{1,2}},?\s+\d{{4}})")

def parse_dates_generic(text):
    r = {}
    m = re.search(
        r"(?:SDS DATE|Revision Date|Date of Issue|Date of compilation|Date of preparation)"
        r".{0,40}?" + _DATE_SHAPE,
        text, re.IGNORECASE
    )
    if m:
        r["sds_date"] = m.group(1)
    r["sds_revision_date"] = r.get("sds_date", "")
    return {k: v for k, v in r.items() if v}


# ── BASF ─────────────────────────────────────────────────────────────────────
def _value_on_next_nonblank_line(label_pattern, text):
    """
    BASF's template puts a field's label as its own heading line, with the
    actual value on the next non-blank line below it — e.g. "Company:" then
    "BASF CORPORATION" two lines later. The shared same-line "Label: value"
    regexes used everywhere else don't apply here.
    """
    lines = text.split("\n")
    label_re = re.compile(label_pattern, re.IGNORECASE)
    for i, line in enumerate(lines):
        if label_re.search(line.strip()):
            for j in range(i + 1, len(lines)):
                if lines[j].strip():
                    return lines[j].strip()
            break
    return ""

def parse_section1_basf(text):
    r = {}
    r["product_name"] = _value_on_next_nonblank_line(r"Product identifier used on the label", text)
    r["company_name"]  = _value_on_next_nonblank_line(r"^Company:?\s*$", text)
    return {k: _cv(v) for k, v in r.items() if v}


# ── Dispatch ─────────────────────────────────────────────────────────────────
_SECTION1_PARSERS = {
    "AMERICHEM": parse_section1_americhem,
    "BUCKLEY":   parse_section1_americhem,
    "DAKOTA":    parse_section1_dakota,
    "CONE":      parse_section1_cone,
    "PHILLIPS66": parse_section1_phillips66,
    "BASF":      parse_section1_basf,
}
_DATE_PARSERS = {
    "AMERICHEM": parse_dates_americhem,
    "BUCKLEY":   parse_dates_americhem,
    "DAKOTA":    parse_dates_dakota,
    "CONE":      parse_dates_cone,
    "PHILLIPS66": parse_dates_phillips66,
}

def parse_section1(text, fmt):
    fn = _SECTION1_PARSERS.get(fmt, parse_section1_generic)
    # Field regexes look for a trailing "\n" (or a specific next-label) after
    # the value to know where it ends. If the field happens to be the very
    # last line of the section, that terminator was stripped off along with
    # everything else — pad one back so the same patterns still match.
    return fn(text + "\n")

def parse_section2(text, fmt, full_text=""):
    text = text + "\n"
    if fmt in ("AMERICHEM", "BUCKLEY"):
        return parse_section2_americhem(text, full_text)
    if fmt == "DAKOTA":
        return parse_section2_dakota(text)
    if fmt == "CONE":
        return parse_section2_cone(text)
    if fmt == "PHILLIPS66":
        return parse_section2_phillips66(text)
    return parse_section2_generic(text)

def parse_dates(text, fmt):
    fn = _DATE_PARSERS.get(fmt, parse_dates_generic)
    return fn(text)

def parse_cas(text):
    r = {}
    cas_matches = re.findall(r"\b(\d{2,7}-\d{2}-\d)\b", text)
    cas_list = list(dict.fromkeys(cas_matches))
    if cas_list:
        r["cas_number_primary"] = cas_list[0]
        r["all_cas_numbers"] = ", ".join(cas_list)
    return r


# ── Section boundary detection, now line-based (robust for image mapping) ──
# Tolerant of stray spaces PDF font kerning sometimes inserts mid-word
# (e.g. "S ECTION 1" instead of "SECTION 1") — seen in real files where
# some lines on a page extract cleanly and others from the same page
# don't, depending on which glyph run pdfplumber is reading.
_SECTION_WORD = r"S\s?E\s?C\s?T\s?I\s?O\s?N"

# Tolerant of stray spaces PDF font kerning sometimes inserts mid-word
# (e.g. "S ECTION 1" instead of "SECTION 1") — seen in real files where
# some lines on a page extract cleanly and others from the same page
# don't, depending on which glyph run pdfplumber is reading.
_SECTION_WORD = r"S\s?E\s?C\s?T\s?I\s?O\s?N"

def _section_pattern(fmt):
    """
    One unified pattern regardless of vendor. Originally this dispatched
    AMERICHEM/BUCKLEY to a dash-required pattern ("SECTION – 1") and
    everything else to a period/colon pattern ("SECTION 1." / "SECTION
    1:") — but real Dakota files use the dash style too (their folder
    mixes multiple source templates), so a format-based dispatch guesses
    wrong for exactly the vendors it's meant to help. Making the dash
    optional handles both conventions unconditionally and safely, since
    it only ever makes matching MORE permissive, never introduces a new
    false-positive path.

    Also case-insensitive now: the line-start anchor in
    _match_section_header already excludes prose cross-references (every
    real example found — "See Section 11 for...", "refer to Section 8...")
    started with a word other than "Section", so that anchor alone was
    sufficient protection. Requiring ALL-CAPS on top of it was redundant
    and incorrectly excluded legitimate Title-Case headers ("Section 1.
    Identification") used by several real vendors (Cal-Lite, AllChem).

    Separator character class widened beyond en-dash/hyphen: confirmed
    real case (OCR'd Americhem file) where the SAME document rendered the
    same intended dash three different ways across three section headers —
    "SECTION =1" (equals sign), "SECTION — 2" (em-dash), "SECTION - 3"
    (plain hyphen) — because OCR doesn't reliably read a stylized dash
    glyph the same way twice. All three needed to be accepted.

    Separator also allowed to repeat (not just zero-or-one): another real
    case from the same file, "SECTION -—6" — OCR rendered one intended
    dash as TWO adjacent characters (hyphen immediately followed by
    em-dash, no space). A single optional separator char couldn't match
    that; a run of any combination of whitespace/dash-like characters can.

    Terminator broadened three ways from real files:
    - middle-dot "·" as trailing separator ("SECTION 2· HAZARDS...",
      Colonial Chemical Solutions' template)
    - hyphen/dash as trailing separator too, not just leading
      ("Section 1- Product Information", an OCR'd file)
    - end-of-line accepted as a valid terminator: some vendors put the
      section number alone on its own line with the title on the NEXT
      line ("SECTION 2" / "HAZARD IDENTIFICATION" as two separate lines,
      DE FO 60.docx) — there's no terminator character at all in that
      case since the line just ends.
    """
    return re.compile(_SECTION_WORD + r"[\s–—\-=]*(\d+)(?:[.\s:·\-–—]|$)", re.IGNORECASE)


def _match_section_header(line, pattern):
    """
    A genuine section header is the ONLY thing on its line. Prose
    cross-references ("See Section 11 for symptoms...", "refer to Section
    8...") appear mid-line, never at the start — anchoring the match to
    the start of the (stripped) line is what excludes those, regardless of
    case. (Case-insensitivity is needed too: several real vendors use
    Title-Case "Section 1. Identification" as their actual header, not
    ALL-CAPS "SECTION".)

    One more check: a cross-reference can coincidentally wrap to the start
    of a new line ("...take note of any information in\nSection 8 on
    suitable materials...") — anchoring alone doesn't catch this. A real
    section title always starts with a capital letter right after the
    number; a wrapped sentence continuation is grammatically lowercase
    ("on suitable...", "for additional..."). Reject the match if what
    follows starts lowercase.
    """
    stripped = line.strip()
    m = pattern.match(stripped)
    if not m:
        return None
    rest = stripped[m.end():].strip()
    if rest and rest[0].islower():
        return None
    return m


# Some vendors (Thermo Fisher/Alfa Aesar, BASF) number sections as bare
# "1. Identification", "10. Stability and Reactivity" — no "SECTION" word at
# all. Matching on canonical GHS section titles (rather than just "N.")
# keeps this safe from false-positiving on ordinary numbered list items
# elsewhere in the document, which don't carry these specific titles.
_GENERIC_TITLE_PATTERNS = {
    1: re.compile(r"identification", re.IGNORECASE),
    2: re.compile(r"hazard", re.IGNORECASE),
    3: re.compile(r"composition", re.IGNORECASE),
    4: re.compile(r"first[\s-]*aid", re.IGNORECASE),
    5: re.compile(r"fire", re.IGNORECASE),
    6: re.compile(r"accidental release", re.IGNORECASE),
    7: re.compile(r"handling and storage", re.IGNORECASE),
    8: re.compile(r"exposure control", re.IGNORECASE),
    9: re.compile(r"physical and chemical", re.IGNORECASE),
    10: re.compile(r"stability and reactivit", re.IGNORECASE),
    11: re.compile(r"toxicolog", re.IGNORECASE),
    12: re.compile(r"ecological", re.IGNORECASE),
    13: re.compile(r"disposal", re.IGNORECASE),
    14: re.compile(r"transport", re.IGNORECASE),
    15: re.compile(r"regulatory", re.IGNORECASE),
    16: re.compile(r"other information", re.IGNORECASE),
}
_GENERIC_NUM_PREFIX = re.compile(r"^(\d{1,2})\.\s*(.{0,80})")

def _match_generic_numbered_header(line):
    m = _GENERIC_NUM_PREFIX.match(line.strip())
    if not m:
        return None
    num = int(m.group(1))
    title_pattern = _GENERIC_TITLE_PATTERNS.get(num)
    if title_pattern and title_pattern.search(m.group(2)):
        return num
    return None


def split_sections_by_line(lines, fmt):
    """Returns {section_num: (start_line_idx, end_line_idx_exclusive)}."""
    pattern = _section_pattern(fmt)
    boundaries = {}
    for i, line in enumerate(lines):
        m = _match_section_header(line, pattern)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 16 and n not in boundaries:
                boundaries[n] = i

    # Always fill any GAPS with the bare-numbered fallback ("1. Identification",
    # no "SECTION" word) — not just when the primary pattern found nothing at
    # all. Some documents mix conventions (e.g. an Americhem-format vendor doc
    # using "SECTION" for later sections but bare numbers for early ones), and
    # an all-or-nothing fallback would miss exactly those early sections.
    for i, line in enumerate(lines):
        n = _match_generic_numbered_header(line)
        if n and n not in boundaries:
            boundaries[n] = i

    # Last resort for Section 1 specifically: confirmed real cases
    # (Hamposyl O, Amsol 400) where Section 1 has NO number and NO
    # "SECTION" word at all — just a bare title line like
    # "Identification" or "Identification of the Substance/Preparation
    # and the Company/Undertaking" — while every other section DOES have
    # a number. Only scan the preamble before the earliest other found
    # boundary, since scanning the whole document risks matching the word
    # "identification" in unrelated body text (e.g. a cross-reference).
    if 1 not in boundaries and boundaries:
        earliest_other = min(boundaries.values())
        title_pattern = _GENERIC_TITLE_PATTERNS[1]
        for i in range(earliest_other):
            stripped = lines[i].strip()
            if stripped and len(stripped) < 100 and title_pattern.search(stripped):
                boundaries[1] = i
                break

    # Sort by DOCUMENT POSITION, not section number. A section's end is
    # whatever the next header actually encountered in reading order is —
    # assuming section numbers always increase monotonically with position
    # is exactly what let one bad boundary corrupt every range after it.
    ordered = sorted(boundaries.items(), key=lambda kv: kv[1])
    ranges = {}
    for idx, (num, start) in enumerate(ordered):
        end = ordered[idx + 1][1] if idx + 1 < len(ordered) else len(lines)
        ranges[num] = (start, end)
    return ranges


def _section_for_line(ranges, line_idx):
    for n, (s, e) in ranges.items():
        if s <= line_idx < e:
            return n
    # fallback: last section that had started by this line
    candidates = [n for n, (s, e) in ranges.items() if s <= line_idx]
    return max(candidates, key=lambda n: ranges[n][0]) if candidates else None


# ── DOCX extraction: paragraphs + tables in true document order ───────────
def _blob_ext(rel):
    ct = rel.target_part.content_type
    if "png" in ct:  return "png"
    if "jpeg" in ct or "jpg" in ct: return "jpg"
    if "gif" in ct:  return "gif"
    if "bmp" in ct:  return "bmp"
    if "tiff" in ct: return "tiff"
    if "emf" in ct or "wmf" in ct:
        # Vector formats. Pillow/Word can't reliably rasterize these from raw
        # bytes, and mislabeling them "png" (the old behavior) produces a
        # broken/corrupt image in the generated doc with no error at all —
        # worse than skipping, since it looks like success until someone
        # opens the file. Flag for manual handling instead.
        return None
    return "png"


def _paragraph_images(p_elm, doc):
    """
    Returns (images, n_skipped_vector).
    images: list of (bytes, ext) for formats we can safely re-embed.
    n_skipped_vector: count of EMF/WMF images found but not extracted —
    surfaced in extraction_notes so these don't fail silently.
    """
    out = []
    n_skipped = 0
    for blip in p_elm.findall(f".//{{{DRAWING_NS}}}blip"):
        rId = blip.get(f"{{{REL_NS}}}embed")
        if rId and rId in doc.part.rels:
            rel = doc.part.rels[rId]
            if "image" in rel.reltype:
                ext = _blob_ext(rel)
                if ext is None:
                    n_skipped += 1
                else:
                    out.append((rel.target_part.blob, ext))
    return out, n_skipped


def _list_info(paragraph):
    """
    Returns (is_list, indent_level). Checks two independent ways Word marks
    a paragraph as a list item:
      1. Direct numPr on the paragraph's own pPr (manual/direct formatting)
      2. A "List ..." named style (List Bullet, List Number, List Paragraph,
         etc.) — numbering lives on the style definition, not the paragraph,
         in this case, so direct-numPr-only detection misses it entirely.
    """
    pPr = paragraph._p.pPr
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            ilvl_el = numPr.find(qn("w:ilvl"))
            ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
            return True, ilvl

    style_name = (paragraph.style.name or "") if paragraph.style else ""
    if style_name.lower().startswith("list"):
        m = re.search(r"(\d+)\s*$", style_name)
        ilvl = (int(m.group(1)) - 1) if m else 0
        return True, max(ilvl, 0)

    return False, 0


def extract_docx_blocks(path):
    """
    Returns (lines, image_map, n_skipped_vector_images) where:
      lines      -- ordered list of text lines reconstructing the doc,
                     with bullet/numbered items prefixed "- " (indented per level)
                     and table rows rendered as "| cell | cell | ..."
      image_map  -- list of (line_index, image_bytes, ext) — line_index is the
                     line the image is anchored to, for section bucketing.
      n_skipped_vector_images -- count of EMF/WMF images found but not
                     extracted (unsupported vector format).
    """
    doc = DocxDoc(str(path))
    lines = []
    image_map = []
    skipped_total = 0

    def handle_paragraph(para):
        nonlocal skipped_total
        txt = para.text.strip()
        imgs, n_skipped = _paragraph_images(para._p, doc)
        skipped_total += n_skipped
        is_list, ilvl = _list_info(para)
        if txt or imgs:
            if is_list:
                lines.append(("  " * ilvl) + "- " + txt)
            else:
                lines.append(txt)
            for img_bytes, ext in imgs:
                image_map.append((len(lines) - 1, img_bytes, ext))

    def handle_table(table):
        nonlocal skipped_total
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
            row_line_idx = len(lines) - 1
            for cell in row.cells:
                for p in cell.paragraphs:
                    imgs, n_skipped = _paragraph_images(p._p, doc)
                    skipped_total += n_skipped
                    for img_bytes, ext in imgs:
                        image_map.append((row_line_idx, img_bytes, ext))

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            handle_paragraph(Paragraph(child, doc))
        elif child.tag == qn("w:tbl"):
            handle_table(Table(child, doc))

    return lines, image_map, skipped_total


# ── PDF extraction ──────────────────────────────────────────────────────────
def _extract_pdf_image_bytes(img):
    """Best-effort raw bytes + extension for a pdfplumber page image dict."""
    stream = img.get("stream")
    if stream is None:
        return None, None
    try:
        filt = stream.get("Filter")
    except Exception:
        filt = None
    filt_str = str(filt) if filt else ""

    try:
        data = stream.get_data()
    except Exception:
        return None, None

    if "DCT" in filt_str:  # already-encoded JPEG
        return data, "jpg"

    # Fallback: attempt raw-bitmap reconstruction via Pillow
    if HAVE_PIL:
        try:
            width = int(img.get("width", 0))
            height = int(img.get("height", 0))
            if width and height and len(data) >= width * height:
                mode = "RGB" if len(data) >= width * height * 3 else "L"
                pil_img = Image.frombytes(mode, (width, height), data)
                buf = io.BytesIO()
                pil_img.save(buf, format="PNG")
                return buf.getvalue(), "png"
        except Exception:
            pass
    return None, None


# ── OCR fallback for scanned/image-only PDFs ────────────────────────────────
def ocr_pdf_pages(path, dpi=300):
    """
    Rasterizes each page and runs Tesseract OCR — used only when the PDF has
    no native text layer at all (confirmed via pdfplumber first, not
    guessed). Returns (lines, page_starts, page_count).

    Known limitation: since the source is a flat page image with no
    structural markup, we can't separate individual pictograms/images
    within a section the way the text-layer path does — this path gets you
    text only. If a scanned SDS's pictograms/logos matter for the generated
    document, they'll still need manual insertion.
    """
    images = convert_from_path(str(path), dpi=dpi)
    lines, page_starts = [], []
    for img in images:
        page_starts.append(len(lines))
        text = pytesseract.image_to_string(img) or ""
        for line in text.split("\n"):
            lines.append(line)
    return lines, page_starts, len(images)


def extract_pdf_blocks(path):
    """
    Returns (lines, page_starts, page_images, page_count, is_scanned, line_positions).
    page_images: list of (page_index, y_top, image_bytes, ext).
    line_positions: parallel to `lines` — (page_index, y_top or None) — used
    to map each image to the nearest preceding line by actual position on
    the page, rather than crudely attributing every image on a page to
    whichever section happens to own that page's very first line. A single
    PDF page routinely spans multiple SDS sections, so page-level granularity
    silently misattributes images to the wrong section.
    """
    lines, page_starts, page_images, line_positions = [], [], [], []
    with pdfplumber.open(str(path)) as pdf:
        page_count = len(pdf.pages)
        total_text_len = 0
        for pi, page in enumerate(pdf.pages):
            page_starts.append(len(lines))
            try:
                text_lines = page.extract_text_lines()
            except Exception:
                text_lines = None
            if text_lines:
                for tl in text_lines:
                    txt = tl.get("text", "")
                    total_text_len += len(txt.strip())
                    lines.append(txt)
                    line_positions.append((pi, tl.get("top")))
            else:
                t = page.extract_text() or ""
                for line in t.split("\n"):
                    lines.append(line)
                    line_positions.append((pi, None))
                    total_text_len += len(line.strip())
            for img in page.images:
                data, ext = _extract_pdf_image_bytes(img)
                if data:
                    page_images.append((pi, img.get("top", 0), data, ext))
    is_scanned = total_text_len < 200
    return lines, page_starts, page_images, page_count, is_scanned, line_positions


# Some PDFs have a broken font/ToUnicode mapping — pdfplumber extracts a
# normal-looking CHARACTER COUNT, so is_scanned (length-based) doesn't catch
# it, but the actual characters are wrong (e.g. "CascadCe a1s0c0a0de 1000"
# instead of "Cascade 1000"). No amount of regex tolerance fixes wrong
# characters — this needs OCR instead, reading pixels rather than the
# corrupted embedded character codes. Detected by checking whether common
# English stopwords appear at a normal rate; garbled text won't have them
# as clean whole words even though the raw character count looks fine.
_COMMON_WORDS_RE = re.compile(r"\b(the|and|of|to|in|for|with|is|or|by|this)\b", re.IGNORECASE)

def _looks_garbled(text, min_hits_per_1000_chars=2):
    if len(text) < 500:
        return False  # too short to judge reliably either way
    hits = len(_COMMON_WORDS_RE.findall(text))
    rate = hits / (len(text) / 1000.0)
    return rate < min_hits_per_1000_chars


def _nearest_line_for_image(line_positions, page_starts, page_count, pi, img_top):
    """
    Finds the global line index of the text line immediately above (closest
    preceding) the image's Y position on its page — i.e. whatever content
    the image visually follows. Falls back to the page's first line if no
    positional data is available (e.g. extract_text_lines() unsupported).
    """
    start = page_starts[pi]
    end = page_starts[pi + 1] if pi + 1 < page_count else len(line_positions)
    best_idx, best_dist = start, None
    for i in range(start, end):
        _, top = line_positions[i]
        if top is None or top > img_top:
            continue
        dist = img_top - top
        if best_dist is None or dist < best_dist:
            best_dist, best_idx = dist, i
    return best_idx


# ── Assembly ─────────────────────────────────────────────────────────────────
def _save_images(images_dir, stem, section_num, images):
    """images: list of (bytes, ext). Returns list of relative path strings."""
    if not images:
        return []
    out_dir = images_dir / stem
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for i, (data, ext) in enumerate(images, start=1):
        fname = f"section_{section_num}_img{i}.{ext}"
        fpath = out_dir / fname
        fpath.write_bytes(data)
        paths.append(str(fpath.relative_to(images_dir.parent)))
    return paths


# ── Repeated page-letterhead/footer noise ──────────────────────────────────
# Many vendor templates repeat identifying text as literal body paragraphs at
# the top/bottom of every printed page (not a real Word header/footer, which
# would live in a separate XML part we never see via body paragraphs). Each
# vendor uses a different pattern:
#   Cone/Dakota:  "COMPANY IDENTITY: X   SDS DATE: Y" / "PRODUCT IDENTITY: X   ORIGINAL: Y"
#   Americhem:    "Page 2 of 5 Hexane Revision Date 2/2/2025"
# The legitimate Section 1 fields never take these exact forms, so each
# pattern safely distinguishes noise from real content.
_LETTERHEAD_NOISE_RE = re.compile(
    r"^\s*(COMPANY IDENTITY|PRODUCT IDENTITY)\s*:\s+.*\b(SDS DATE|ORIGINAL)\s*:",
    re.IGNORECASE
)
_AMERICHEM_FOOTER_RE = re.compile(
    r"^\s*Page\s+\d+\s+of\s+\d+\b.*\bRevision Date\b",
    re.IGNORECASE
)

def _strip_letterhead_noise(lines):
    """
    Blanks out (doesn't remove — preserves line indices for image mapping)
    repeated-letterhead/footer noise lines, EXCEPT the first occurrence,
    which is kept intact. Some documents only ever state the SDS date
    inside this repeating banner line — no separate standalone "SDS DATE:"
    field exists anywhere else in them — so blanking every occurrence would
    silently destroy the only copy of the date in the whole document.
    Keeping just the first occurrence still eliminates the repetitive
    noise from every subsequent page, while leaving date/company parsing
    (which scans the full joined text, not just rendered section bodies)
    something to actually find.
    """
    kept_once = False
    out = []
    for ln in lines:
        is_noise = bool(_LETTERHEAD_NOISE_RE.match(ln) or _AMERICHEM_FOOTER_RE.match(ln))
        if is_noise:
            if not kept_once:
                out.append(ln)
                kept_once = True
            else:
                out.append("")
        else:
            out.append(ln)
    return out


def collapse_adjacent_duplicate_headers(lines, fmt):
    """
    Some vendors restate the current section's header as a running
    page-continuation banner on every new page (e.g. "Section 2. Hazards
    identification" reappearing verbatim at the top of page 2, page 3...)
    without adding "(CONTINUED)" — and some PDFs extract a header twice in
    immediate succession due to how pdfplumber reads colored banner
    graphics. Both are harmless: the defining property is that the SAME
    section number repeats with NO OTHER section's header appearing in
    between the two occurrences. That's different from genuinely dangerous
    scattered duplication (the original Cone PDF-conversion-artifact bug),
    where a section number resurfaced only AFTER several other real
    sections' headers had already appeared — that pattern still gets
    flagged by find_duplicate_section_headers, since this function only
    collapses same-number-with-nothing-else-in-between repeats.
    """
    pattern = _section_pattern(fmt)
    out = list(lines)
    last_seen_num = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        m = _match_section_header(line, pattern)
        if m:
            n = int(m.group(1))
            if n == last_seen_num:
                out[i] = ""  # same section restating itself — harmless
            else:
                last_seen_num = n
    return out


def find_duplicate_section_headers(lines, fmt):
    """
    Returns sorted list of section numbers whose header pattern (e.g.
    "SECTION 5.") appears more than once in the document. split_sections_by_line
    only ever uses the *first* occurrence as the boundary, so a duplicate
    doesn't cause a crash — but it's a strong signal that content got
    duplicated/scrambled upstream (classic PDF->DOCX conversion artifact),
    which can silently produce a corrupted merged document otherwise.
    Call collapse_adjacent_duplicate_headers on `lines` first so benign
    back-to-back PDF extraction artifacts aren't flagged here.
    """
    pattern = _section_pattern(fmt)
    counts = {}
    for line in lines:
        m = _match_section_header(line, pattern)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 16:
                if re.search(r"CONTINUED", line, re.IGNORECASE):
                    continue  # expected — multi-page section continuation, not a duplicate
                counts[n] = counts.get(n, 0) + 1
    return sorted(n for n, c in counts.items() if c > 1)


def parse_sds(path, images_dir=None, root=None):
    """
    root: the top-level folder the batch scan started from (e.g. the
    "SDS -12 July 2026" folder). When given, source_folder correctly
    identifies the top-level vendor folder (e.g. "SDS - Cone") no matter
    how deeply the actual file is nested under it — without root, we can
    only fall back to the immediate parent, which breaks once vendor
    folders have their own subfolders.
    """
    path = Path(path)
    if images_dir is None:
        images_dir = path.parent / "extracted_images"
    images_dir = Path(images_dir)

    row = {k: "" for k in FIELDS}
    row["filename"]  = path.name
    row["file_path"] = str(path.resolve())
    row["file_link"] = f'=HYPERLINK("{str(path.resolve())}","{path.name}")'
    row["file_type"] = path.suffix.upper().lstrip(".")

    if root is not None:
        try:
            rel = path.resolve().relative_to(Path(root).resolve())
            row["relative_path"] = str(rel)
            # rel.parts[0] is the top-level vendor folder (e.g. "SDS - Cone")
            # when the file is nested at all; if it's directly in root with
            # no subfolder, there's no vendor folder to report.
            row["source_folder"] = rel.parts[0] if len(rel.parts) > 1 else ""
        except ValueError:
            row["relative_path"] = str(path)
            row["source_folder"] = path.parent.name
    else:
        row["relative_path"] = str(path)
        row["source_folder"] = path.parent.name

    tmp_convert_dir = None
    try:
        if path.suffix.lower() == ".doc":
            converted_path, tmp_convert_dir = try_convert_doc_to_docx(path)
            if converted_path is None:
                # No reliable pure-Python reader exists for old .doc without
                # an external tool, and textutil either isn't available
                # (not macOS) or failed on this specific file.
                row["extraction_status"] = "LEGACY .DOC FORMAT — needs manual conversion to .docx"
                row["extraction_notes"] = (
                    "This is an old binary Word 97-2003 (.doc) file. Automatic conversion via "
                    "textutil failed or isn't available. Open it in Word and use 'Save As' → "
                    ".docx, then re-run the parser on the converted copy."
                )
                return row
            # Conversion succeeded — parse the converted .docx, but keep all
            # the original file's identity (filename/path/etc, already set
            # above) so it's clear which source file this row is about.
            docx_working_path = converted_path
        else:
            docx_working_path = path

        if path.suffix.lower() == ".pdf":
            lines, page_starts, page_images, page_count, is_scanned, line_positions = extract_pdf_blocks(path)
            row["total_pages"] = str(page_count)
            row["is_scanned"] = "YES" if is_scanned else "NO"
            row["ocr_used"] = "NO"
            is_garbled = (not is_scanned) and _looks_garbled("\n".join(lines))
            if is_scanned or is_garbled:
                if not HAVE_OCR:
                    if is_garbled:
                        row["extraction_status"] = "GARBLED TEXT — NEEDS MANUAL/OCR"
                        row["extraction_notes"] = (
                            "PDF has a text layer but it appears corrupted (broken font/encoding "
                            "table — common words aren't appearing normally). OCR dependencies "
                            "(tesseract/poppler) are not installed on this machine."
                        )
                    else:
                        row["extraction_status"] = "SCANNED — NEEDS MANUAL/OCR"
                        row["extraction_notes"] = (
                            "Scanned PDF — no text layer, and OCR dependencies "
                            "(tesseract/poppler) are not installed on this machine."
                        )
                    return row
                try:
                    ocr_lines, ocr_page_starts, ocr_page_count = ocr_pdf_pages(path)
                except Exception as e:
                    row["extraction_status"] = "SCANNED — OCR FAILED"
                    row["extraction_notes"] = f"OCR attempt failed: {type(e).__name__}: {e}"
                    return row
                ocr_text_len = sum(len(l.strip()) for l in ocr_lines)
                if ocr_text_len < 200:
                    row["extraction_status"] = "SCANNED — OCR INSUFFICIENT"
                    row["extraction_notes"] = (
                        "Scanned PDF — OCR ran but produced little/no usable text "
                        "(likely poor scan quality). Manual entry required."
                    )
                    return row
                # OCR succeeded — fall through into the normal pipeline below,
                # same as a native text-layer PDF, just with OCR'd lines.
                lines, page_starts, page_count = ocr_lines, ocr_page_starts, ocr_page_count
                page_images, line_positions = [], [(0, None)] * len(lines)
                row["total_pages"] = str(page_count)
                row["ocr_used"] = "YES"

            lines = _strip_letterhead_noise(lines)
            full_text = "\n".join(lines)
            fmt = detect_format(full_text)
            row["format"] = fmt
            lines = collapse_adjacent_duplicate_headers(lines, fmt)
            full_text = "\n".join(lines)
            ranges = split_sections_by_line(lines, fmt)
            duplicate_sections = find_duplicate_section_headers(lines, fmt)

            # Last resort: native text extraction found real body content
            # (not scanned, not garbled) but ZERO section boundaries
            # anywhere. Confirmed real case: some vendors render the
            # "SECTION – N" banners as pure image graphics with no
            # accompanying text layer at all, even though the surrounding
            # body text extracts fine. OCR reads pixels, not embedded text
            # objects, so it can sometimes see header text native
            # extraction never had access to. Only actually switch to the
            # OCR result if it finds MORE sections than native did —
            # otherwise the native result stands, since OCR isn't
            # guaranteed to be more accurate for the body text itself.
            if not ranges and row["ocr_used"] == "NO" and HAVE_OCR:
                try:
                    ocr_lines, ocr_page_starts, ocr_page_count = ocr_pdf_pages(path)
                    if sum(len(l.strip()) for l in ocr_lines) >= 200:
                        ocr_lines = _strip_letterhead_noise(ocr_lines)
                        ocr_fmt = detect_format("\n".join(ocr_lines))
                        ocr_lines = collapse_adjacent_duplicate_headers(ocr_lines, ocr_fmt)
                        ocr_ranges = split_sections_by_line(ocr_lines, ocr_fmt)
                        if ocr_ranges:
                            lines, fmt, ranges = ocr_lines, ocr_fmt, ocr_ranges
                            full_text = "\n".join(lines)
                            row["format"] = fmt
                            row["total_pages"] = str(ocr_page_count)
                            row["ocr_used"] = "YES"
                            duplicate_sections = find_duplicate_section_headers(lines, fmt)
                            page_images, line_positions = [], [(0, None)] * len(lines)
                except Exception:
                    pass  # native extraction result stands if this attempt fails

            section_images = {n: [] for n in range(1, 17)}
            unmapped_images = 0
            skipped_vector_images = 0
            for pi, img_top, data, ext in page_images:
                line_idx = _nearest_line_for_image(line_positions, page_starts, page_count, pi, img_top)
                n = _section_for_line(ranges, line_idx)
                if n:
                    section_images[n].append((data, ext))
                else:
                    unmapped_images += 1

        else:  # DOCX (or a .doc successfully auto-converted to .docx above)
            lines, image_map, skipped_vector_images = extract_docx_blocks(docx_working_path)
            row["total_pages"] = "0"
            row["is_scanned"] = "NO"
            row["ocr_used"] = "NO"
            lines = _strip_letterhead_noise(lines)
            full_text = "\n".join(lines)
            fmt = detect_format(full_text)
            row["format"] = fmt
            lines = collapse_adjacent_duplicate_headers(lines, fmt)
            full_text = "\n".join(lines)
            ranges = split_sections_by_line(lines, fmt)
            duplicate_sections = find_duplicate_section_headers(lines, fmt)

            section_images = {n: [] for n in range(1, 17)}
            unmapped_images = 0
            for line_idx, data, ext in image_map:
                n = _section_for_line(ranges, line_idx)
                if n:
                    section_images[n].append((data, ext))
                else:
                    unmapped_images += 1

        section_text = {}
        for n in range(1, 17):
            if n in ranges:
                s, e = ranges[n]
                section_text[n] = "\n".join(lines[s:e]).strip()
                row[f"section_{n}_text"] = section_text[n]
            saved = _save_images(images_dir, path.stem, n, section_images.get(n, []))
            row[f"section_{n}_images"] = "; ".join(saved)

        # Dates — parsed against full text since headers/footers can fall
        # outside the section-1 boundary depending on layout
        row.update(parse_dates(full_text, fmt))

        # Section 1 & 2 structured fields, company-specific
        if 1 in section_text:
            row.update(parse_section1(section_text[1], fmt))
        if 2 in section_text:
            row.update(parse_section2(section_text[2], fmt, full_text=full_text))

        # CAS — needed for winner-selection grouping. Section 3 is the
        # canonical location; fall back to full text if section 3 wasn't found.
        cas_source = section_text.get(3, "") or full_text
        row.update(parse_cas(cas_source))

        missing = sorted(set(range(1, 17)) - set(ranges.keys()))
        folder_mismatch = _folder_format_mismatch(row.get("source_folder", ""), fmt)
        notes = []
        if row.get("ocr_used") == "YES":
            notes.append("Text extracted via OCR (no native text layer) — verify against "
                          "source before trusting exact wording/numbers; pictograms/images "
                          "in this document were not extracted")
        if missing:
            notes.append(f"Missing sections: {missing}")
        if duplicate_sections:
            notes.append(f"Duplicate section header(s) found: {duplicate_sections} — "
                          f"likely content duplication/scrambling from source conversion")
        if folder_mismatch:
            notes.append(f"Folder/format mismatch: found in '{row.get('source_folder')}' "
                          f"folder but detected format is '{fmt}'")
        if unmapped_images:
            notes.append(f"{unmapped_images} image(s) found outside any detected section")
        if skipped_vector_images:
            notes.append(f"{skipped_vector_images} vector image(s) (EMF/WMF) could not be "
                          f"extracted — needs manual insertion")
        if not row.get("sds_date"):
            notes.append("No SDS date found — winner selection for this file needs manual date entry")
        row["extraction_notes"] = " | ".join(notes)

        # ── Triage status — worst issue wins, so sorting/filtering the CSV
        # by this one column tells you exactly what needs a human look ──
        critical_missing = sorted(set([1, 2, 3]) - set(ranges.keys()))
        if duplicate_sections:
            # Prioritized above missing-section checks: a duplicated file can
            # still show all 16 sections "found" while containing scrambled,
            # repeated content in between — worse than merely incomplete.
            row["extraction_status"] = (
                f"DUPLICATE CONTENT — section(s) {duplicate_sections} appear "
                f"more than once, manual review required"
            )
        elif critical_missing:
            # An unrecognized vendor AND missing sections is worth flagging
            # distinctly — could be a genuinely new format needing a parser
            # fix, not just a normal incomplete-document case.
            prefix = "UNKNOWN FORMAT" if fmt == "UNKNOWN" else "INCOMPLETE"
            row["extraction_status"] = f"{prefix} — critical section(s) {critical_missing} not found"
        elif len(missing) > 2:
            prefix = "UNKNOWN FORMAT" if fmt == "UNKNOWN" else "INCOMPLETE"
            row["extraction_status"] = f"{prefix} — {len(missing)} sections not found"
        elif fmt == "UNKNOWN":
            # Vendor not recognized, but extraction actually succeeded (all
            # sections found, nothing scrambled) — the generic Section 1/2
            # field parsers likely handled it fine too. Not worth a manual
            # review flag on its own; noted for future dedicated-parser work.
            row["extraction_status"] = "OK (unrecognized vendor — generic field extraction used)"
        elif folder_mismatch:
            row["extraction_status"] = (
                f"FOLDER MISMATCH — in '{row.get('source_folder')}' folder but detected as {fmt}"
            )
        elif not row.get("cas_number_primary"):
            row["extraction_status"] = "NO CAS FOUND — cannot group for winner selection"
        elif not row.get("sds_date"):
            row["extraction_status"] = "MISSING DATE — needs manual entry before winner selection"
        else:
            row["extraction_status"] = "OK"

        if tmp_convert_dir is not None:
            row["extraction_notes"] = (
                "Auto-converted from legacy .doc to .docx via macOS textutil. "
                + row.get("extraction_notes", "")
            ).strip(" |")

    except Exception as e:
        row["extraction_status"] = "PARSE ERROR"
        row["extraction_notes"] = f"Parse error: {type(e).__name__}: {e}"
    finally:
        if tmp_convert_dir is not None:
            shutil.rmtree(tmp_convert_dir, ignore_errors=True)
    return row


if __name__ == "__main__":
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
    supported = {".pdf", ".docx", ".doc"}
    images_dir = folder / "extracted_images"
    # Recursive: real batches are organized into per-vendor subfolders
    # (Cone/, Americhem/, Dakota/, ...) rather than flat in one directory.
    # Excludes extracted_images/ so a re-run doesn't try to "parse" the
    # image files a prior run already extracted there.
    files = sorted(
        f for f in folder.rglob("*")
        if f.is_file()
        and f.suffix.lower() in supported
        and images_dir not in f.parents
    )
    print(f"Parsing {len(files)} files...")
    rows = []
    for fi in files:
        print(f"  {fi.relative_to(folder)}...", end=" ", flush=True)
        try:
            row = parse_sds(fi, images_dir=images_dir, root=folder)
        except Exception as e:
            # Belt-and-suspenders: parse_sds already catches its own errors,
            # but with 3,000 files we do not want one truly unexpected crash
            # (e.g. a corrupt file the libraries choke on before entering the
            # try block) to kill the entire batch. Record it and keep going.
            row = {k: "" for k in FIELDS}
            row["filename"] = fi.name
            row["file_path"] = str(fi.resolve())
            row["file_type"] = fi.suffix.upper().lstrip(".")
            row["extraction_status"] = "PARSE ERROR"
            row["extraction_notes"] = f"Unhandled error: {type(e).__name__}: {e}"
        rows.append(row)
        n_imgs = sum(1 for n in range(1, 17) if row.get(f"section_{n}_images"))
        print(f"{row.get('format','?')} | status: {row.get('extraction_status','?')} | "
              f"sections found: {sum(1 for n in range(1,17) if row.get(f'section_{n}_text'))}/16 | "
              f"sections w/ images: {n_imgs}")

    out = folder / "SDS_Sections.csv"
    with open(out, "w", newline="", encoding="utf-8") as fout:
        writer = csv.DictWriter(fout, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)

    # ── Triage summary — the point of this with 3,000 incoming files is that
    # you never have to eyeball every row to know what needs a human look ──
    from collections import Counter
    status_counts = Counter(r.get("extraction_status", "") for r in rows)
    needs_review = [r for r in rows if r.get("extraction_status") != "OK"]

    review_fields = ["filename", "source_folder", "relative_path", "format", "extraction_status",
                      "extraction_notes", "product_name", "cas_number_primary", "sds_date"]
    review_out = folder / "SDS_Needs_Review.csv"
    with open(review_out, "w", newline="", encoding="utf-8") as fout:
        writer = csv.DictWriter(fout, fieldnames=review_fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(needs_review)

    print(f"\nSaved to {out}")
    print(f"Images saved under {images_dir}")
    print(f"\n=== TRIAGE SUMMARY ({len(rows)} files) ===")
    for status, count in status_counts.most_common():
        print(f"  {count:5} | {status}")
    print(f"\n{len(needs_review)} file(s) need review — see {review_out}")
    print(f"Images saved under {images_dir}")
