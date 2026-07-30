"""
debug_lines.py

Quick diagnostic tool: dumps the raw extracted text (before section
splitting) for one or more SDS files, so it can be copy-pasted into chat
instead of re-uploading the file. Replicates the EXACT SAME extraction and
OCR-fallback chain as sds_parser.py's parse_sds() (scanned -> OCR, garbled
-> OCR, zero-sections-found -> last-resort OCR), so what you see here is
always what the real parser actually produces, not a simplified version
that can drift out of sync.

Usage:
    python3 debug_lines.py "path/to/file1.pdf" "path/to/file2.docx" ...
        Prints the first 60 lines of raw extracted text with line numbers,
        plus detected format, OCR status, and which sections were found.

    python3 debug_lines.py --grep "Section 8" "path/to/file.pdf"
        Searches the ENTIRE document (not just the first 60 lines) for
        lines containing the given text (case-insensitive), printing each
        match with 2 lines of context before/after and its line number.
"""
import sys, re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import sds_parser as sp


def _get_lines(path):
    """
    Returns (lines, fmt, ranges, ocr_used) — replicates parse_sds()'s full
    chain: native extraction -> scanned/garbled OCR fallback -> zero-
    sections-found last-resort OCR fallback -> format detection -> header
    collapsing -> section splitting.
    """
    path = Path(path)
    if path.suffix.lower() == ".doc":
        converted, tmp_dir = sp.try_convert_doc_to_docx(path)
        if converted is None:
            raise RuntimeError("Could not convert .doc — textutil unavailable or failed.")
        path_to_read = converted
    else:
        path_to_read = path

    if path.suffix.lower() == ".pdf":
        lines, page_starts, page_images, page_count, is_scanned, line_positions = sp.extract_pdf_blocks(path_to_read)
        is_garbled = (not is_scanned) and sp._looks_garbled("\n".join(lines))
        ocr_used = False

        if is_scanned or is_garbled:
            if not sp.HAVE_OCR:
                raise RuntimeError("Scanned/garbled, and OCR not available in this environment.")
            lines, page_starts, page_count = sp.ocr_pdf_pages(path_to_read)
            ocr_used = True

        lines = sp._strip_letterhead_noise(lines)
        fmt = sp.detect_format("\n".join(lines))
        lines = sp.collapse_adjacent_duplicate_headers(lines, fmt)
        ranges = sp.split_sections_by_line(lines, fmt)

        # Same last-resort fallback as parse_sds: native (or already-OCR'd)
        # extraction found ZERO section boundaries anywhere -> try OCR,
        # only switch over if OCR actually finds more than native did.
        if not ranges and not ocr_used and sp.HAVE_OCR:
            try:
                ocr_lines, ocr_page_starts, ocr_page_count = sp.ocr_pdf_pages(path_to_read)
                if sum(len(l.strip()) for l in ocr_lines) >= 200:
                    ocr_lines = sp._strip_letterhead_noise(ocr_lines)
                    ocr_fmt = sp.detect_format("\n".join(ocr_lines))
                    ocr_lines = sp.collapse_adjacent_duplicate_headers(ocr_lines, ocr_fmt)
                    ocr_ranges = sp.split_sections_by_line(ocr_lines, ocr_fmt)
                    if ocr_ranges:
                        lines, fmt, ranges = ocr_lines, ocr_fmt, ocr_ranges
                        ocr_used = True
            except Exception:
                pass
        return lines, fmt, ranges, ocr_used

    else:
        lines, image_map, skipped = sp.extract_docx_blocks(path_to_read)
        lines = sp._strip_letterhead_noise(lines)
        fmt = sp.detect_format("\n".join(lines))
        lines = sp.collapse_adjacent_duplicate_headers(lines, fmt)
        ranges = sp.split_sections_by_line(lines, fmt)
        return lines, fmt, ranges, False


def dump_file(path):
    path = Path(path)
    print(f"\n{'='*70}")
    print(f"FILE: {path.name}")
    print('='*70)
    try:
        lines, fmt, ranges, ocr_used = _get_lines(path)
        print(f"  detected format: {fmt}")
        print(f"  ocr_used: {ocr_used}")
        print(f"  sections found: {sorted(ranges.keys())} ({len(ranges)}/16)")
        print(f"  total lines extracted: {len(lines)}")
        print(f"  --- first 60 lines ---")
        for i, line in enumerate(lines[:60]):
            print(f"  {i:3} | {line}")
    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")


def grep_file(path, term, context=2):
    path = Path(path)
    print(f"\n{'='*70}")
    print(f"FILE: {path.name}  (searching for: {term!r})")
    print('='*70)
    try:
        lines, fmt, ranges, ocr_used = _get_lines(path)
        print(f"  detected format: {fmt} | ocr_used: {ocr_used} | sections found: {sorted(ranges.keys())}")
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        found = False
        for i, line in enumerate(lines):
            if pattern.search(line):
                found = True
                start = max(0, i - context)
                end = min(len(lines), i + context + 1)
                print(f"  --- match at line {i} ---")
                for j in range(start, end):
                    marker = ">>" if j == i else "  "
                    print(f"  {marker} {j:3} | {lines[j]}")
                print()
        if not found:
            print(f"  No matches for {term!r} anywhere in the document ({len(lines)} lines total).")
    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 debug_lines.py file1.pdf [file2.docx ...]")
        print('   or: python3 debug_lines.py --grep "search term" file1.pdf [file2.pdf ...]')
        sys.exit(1)

    if sys.argv[1] == "--grep":
        if len(sys.argv) < 4:
            print('Usage: python3 debug_lines.py --grep "search term" file1.pdf [file2.pdf ...]')
            sys.exit(1)
        term = sys.argv[2]
        for arg in sys.argv[3:]:
            grep_file(arg, term)
    else:
        for arg in sys.argv[1:]:
            dump_file(arg)
