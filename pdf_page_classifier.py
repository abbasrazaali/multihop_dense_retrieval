"""
pdf_page_classifier.py
======================
Deterministic, rule-based classification of PDF pages using PyMuPDF (pymupdf / fitz).

Each page is classified as one of:

    scanned                 - page is essentially one big raster image, no real text
    scanned_with_text_layer - full-page image WITH text (usually an invisible OCR layer)
    text_only               - real extractable text, no meaningful raster images
    text_and_images         - real text plus meaningful embedded pictures
    text_and_vector         - real text plus heavy vector graphics (charts/diagrams drawn
                              as paths, which get_images() cannot see)
    image_only              - meaningful images but NOT full-page coverage, and no text
                              (e.g. a standalone figure page)
    empty                   - nothing meaningful on the page

Signals used (all concrete, all inspectable):
    1. Extractable text length            -> page.get_text("text")
    2. Rendered raster images + coverage  -> page.get_image_info()  (rendered bboxes,
                                             NOT get_images(), which lists referenced
                                             XObjects that may never be drawn)
    3. Invisible text render mode (OCR)   -> page.get_texttrace()  (span "type" == 3)
    4. Vector drawing count               -> page.get_drawings()

NOTE: PDF has no "this page is a scan" flag. This classifier is deterministic
(same input -> same output, rules stated exactly), which is as close to
"without any doubt" as the format allows. Tune the thresholds for your corpus.

Usage:
    python pdf_page_classifier.py document.pdf
    python pdf_page_classifier.py document.pdf --json
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, asdict, field

import fitz  # PyMuPDF


# ---------------------------------------------------------------------------
# Tunable thresholds
# ---------------------------------------------------------------------------

@dataclass
class Thresholds:
    # Minimum characters of extracted text for a page to count as "has text".
    # Filters out stray artifacts like a lone page number.
    text_len: int = 30

    # Fraction of the page area covered by raster images at/above which the
    # page is treated as a full-page image (i.e. a scan candidate).
    full_page_coverage: float = 0.80

    # Images smaller than this fraction of the page area are treated as
    # decorative (logos, header rules exported as tiny images) and ignored.
    min_image_frac: float = 0.01

    # Number of vector drawing objects above which the page is considered to
    # contain "real" vector graphics (borders/rules create a handful on
    # perfectly normal text pages, so keep this comfortably above zero).
    vector_drawings: int = 20


# ---------------------------------------------------------------------------
# Per-page result
# ---------------------------------------------------------------------------

@dataclass
class PageReport:
    page_number: int                 # 1-based
    label: str
    text_chars: int
    has_text: bool
    raster_images_total: int         # all rendered raster images
    raster_images_meaningful: int    # after decorative-size filtering
    image_coverage: float            # meaningful image area / page area (0..1)
    is_full_page_image: bool
    has_invisible_ocr_text: bool
    vector_drawing_count: int
    has_vector_graphics: bool
    notes: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Signal extraction helpers
# ---------------------------------------------------------------------------

def _extract_text_length(page: fitz.Page) -> int:
    """Signal 1: length of real extractable text on the page."""
    return len(page.get_text("text").strip())


def _meaningful_image_coverage(page: fitz.Page, th: Thresholds):
    """
    Signals 2 & 3: rendered raster images and how much of the page they cover.

    Uses get_image_info() because it reflects what is actually DRAWN and where.
    get_images() lists image XObjects referenced by the page - an image can be
    referenced without being drawn, drawn multiple times, or drawn tiny.

    Returns (total_count, meaningful_count, coverage_fraction).
    Coverage is a simple area sum clipped to the page; overlapping images can
    in theory push it slightly above what a union would give, which is fine
    for thresholding purposes.
    """
    page_area = abs(page.rect)
    if page_area == 0:
        return 0, 0, 0.0

    infos = page.get_image_info()
    meaningful = 0
    area = 0.0
    for info in infos:
        bbox = fitz.Rect(info["bbox"])
        bbox &= page.rect  # clip images that bleed off-page
        a = abs(bbox)
        if a >= th.min_image_frac * page_area:
            meaningful += 1
            area += a

    return len(infos), meaningful, area / page_area


def _has_invisible_ocr_text(page: fitz.Page) -> bool:
    """
    OCR-layer signature: text rendered in invisible mode (Tr 3).
    page.get_texttrace() returns one dict per span; span["type"] == 3 means
    the text is drawn invisibly - the hallmark of a searchable scan produced
    by Acrobat OCR, ABBYY, ocrmypdf, etc.
    """
    try:
        return any(span.get("type") == 3 for span in page.get_texttrace())
    except Exception:
        # get_texttrace() can fail on exotic/broken content streams;
        # fail open (assume no invisible text) rather than crash.
        return False


def _vector_drawing_count(page: fitz.Page) -> int:
    """
    Vector graphics (charts, diagrams drawn as lines/paths) are invisible to
    get_images()/get_image_info(). get_drawings() surfaces them.
    """
    try:
        return len(page.get_drawings())
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def classify_page(page: fitz.Page, th: Thresholds | None = None) -> PageReport:
    th = th or Thresholds()

    text_chars = _extract_text_length(page)
    has_text = text_chars >= th.text_len

    total_imgs, meaningful_imgs, coverage = _meaningful_image_coverage(page, th)
    is_full_page_image = coverage >= th.full_page_coverage

    invisible_ocr = _has_invisible_ocr_text(page) if has_text else False

    n_drawings = _vector_drawing_count(page)
    has_vector = n_drawings > th.vector_drawings

    notes: list[str] = []

    # --- Decision tree (order matters; rules are exact) -------------------
    if is_full_page_image and not has_text:
        label = "scanned"

    elif is_full_page_image and has_text:
        label = "scanned_with_text_layer"
        if invisible_ocr:
            notes.append(
                "Invisible text render mode (Tr 3) detected: text layer is "
                "almost certainly machine OCR over a scan."
            )
        else:
            notes.append(
                "Full-page image plus VISIBLE text: could be a design-heavy "
                "page with a background image rather than a scan. Inspect if "
                "this matters for your pipeline."
            )

    elif has_text and meaningful_imgs > 0:
        label = "text_and_images"

    elif has_text and has_vector:
        label = "text_and_vector"
        notes.append(
            f"{n_drawings} vector drawing objects: page likely contains "
            "charts/diagrams drawn as paths (not raster images)."
        )

    elif has_text:
        label = "text_only"

    elif meaningful_imgs > 0:
        label = "image_only"

    else:
        label = "empty"
        if total_imgs > meaningful_imgs:
            notes.append(
                f"{total_imgs - meaningful_imgs} decorative-size image(s) "
                "were ignored."
            )

    if total_imgs > meaningful_imgs and label not in ("empty",):
        notes.append(
            f"Ignored {total_imgs - meaningful_imgs} decorative image(s) "
            f"(< {th.min_image_frac:.0%} of page area)."
        )

    return PageReport(
        page_number=page.number + 1,
        label=label,
        text_chars=text_chars,
        has_text=has_text,
        raster_images_total=total_imgs,
        raster_images_meaningful=meaningful_imgs,
        image_coverage=round(coverage, 4),
        is_full_page_image=is_full_page_image,
        has_invisible_ocr_text=invisible_ocr,
        vector_drawing_count=n_drawings,
        has_vector_graphics=has_vector,
        notes=notes,
    )


def classify_pdf(path: str, th: Thresholds | None = None) -> list[PageReport]:
    """Classify every page of a PDF. Returns one PageReport per page."""
    th = th or Thresholds()
    reports: list[PageReport] = []
    with fitz.open(path) as doc:
        for page in doc:
            reports.append(classify_page(page, th))
    return reports


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _main() -> int:
    ap = argparse.ArgumentParser(
        description="Classify each PDF page as scanned / text-only / "
                    "text+images / etc. using PyMuPDF."
    )
    ap.add_argument("pdf", help="Path to the PDF file")
    ap.add_argument("--json", action="store_true",
                    help="Emit full per-page reports as JSON")
    ap.add_argument("--text-len", type=int, default=Thresholds.text_len,
                    help="Min extracted chars to count as 'has text' "
                         "(default: %(default)s)")
    ap.add_argument("--coverage", type=float,
                    default=Thresholds.full_page_coverage,
                    help="Image coverage fraction to count as full-page "
                         "(default: %(default)s)")
    ap.add_argument("--min-image-frac", type=float,
                    default=Thresholds.min_image_frac,
                    help="Images below this fraction of page area are "
                         "treated as decorative (default: %(default)s)")
    ap.add_argument("--vector-drawings", type=int,
                    default=Thresholds.vector_drawings,
                    help="Drawing-object count above which the page counts "
                         "as having vector graphics (default: %(default)s)")
    args = ap.parse_args()

    th = Thresholds(
        text_len=args.text_len,
        full_page_coverage=args.coverage,
        min_image_frac=args.min_image_frac,
        vector_drawings=args.vector_drawings,
    )

    try:
        reports = classify_pdf(args.pdf, th)
    except Exception as e:
        print(f"Error opening/reading '{args.pdf}': {e}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps([asdict(r) for r in reports], indent=2))
        return 0

    # Human-readable table
    print(f"{'page':>5}  {'label':<24} {'text':>6}  {'imgs':>4}  "
          f"{'cover':>6}  {'ocr':>3}  {'vect':>4}")
    print("-" * 64)
    for r in reports:
        print(f"{r.page_number:>5}  {r.label:<24} {r.text_chars:>6}  "
              f"{r.raster_images_meaningful:>4}  "
              f"{r.image_coverage:>6.1%}  "
              f"{'yes' if r.has_invisible_ocr_text else ' - ':>3}  "
              f"{r.vector_drawing_count:>4}")
        for note in r.notes:
            print(f"       note: {note}")

    # Summary
    counts: dict[str, int] = {}
    for r in reports:
        counts[r.label] = counts.get(r.label, 0) + 1
    print("-" * 64)
    print("summary:", ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    return 0


if __name__ == "__main__":
    sys.exit(_main())
