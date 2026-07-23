"""
overlay.py — PDF Overlay / Annotation Module (Module 3, v3.1 — text-render fix)
------------------------------------------------------------------------------
Stamps site-survey values on top of the original Fenesta WCS Report PDF using
PyMuPDF (fitz).

Public API:
    overlay_survey_data(pdf_bytes, rows, surveyor_name="") -> bytes

⚠️  Why the X-coordinates are hardcoded:
    Fenesta's data-entry cells are drawn as long horizontal underlines,
    not as closed rectangles. get_drawings() returns each stroke as a
    separate path, so there's no reliable "cell shape" to pick up. The
    visual columns are fixed in the template — measured once, they stay
    stable across every page of every order. If the template changes,
    re-measure the CELL_X_* constants.

🩹  v3.1 fix (2026-07-23):
    The "Aperture Size" row in the Fenesta template is only ~6-8 pt tall
    between two closely-spaced horizontal rules. PyMuPDF's insert_textbox()
    silently fails when the rect is shorter than one line of text at the
    minimum font, so the previous version rendered the coloured box but
    NO TEXT INSIDE IT.

    Fix:
      1) Enforce MIN_MAIN_CELL_HEIGHT so we always have room to draw.
      2) Use insert_text() (point-based, always renders) for the main
         cell instead of insert_textbox() (rect-based, height-sensitive).
      3) Auto-fit font size by measuring text width first.
      4) Remarks row still uses insert_textbox — it has a fixed 12pt height
         and always fits.
"""

from __future__ import annotations

import io
from typing import Any, Iterable, Optional

import fitz  # PyMuPDF

from utils import row_tolerance

__version__ = "3.1.0"


# =============================================================================
# CONFIGURABLE CONSTANTS (tune these against your source template)
# =============================================================================

# ---- Colors (RGB, 0..1) — keyed to tolerance status ------------------------
STATUS_COLORS: dict[str, dict[str, tuple[float, float, float]]] = {
    "ok":     {"border": (0.07, 0.72, 0.42), "fill": (0.87, 0.96, 0.90)},  # green
    "warn":   {"border": (0.97, 0.57, 0.04), "fill": (1.00, 0.94, 0.83)},  # amber
    "danger": {"border": (0.94, 0.27, 0.22), "fill": (1.00, 0.90, 0.89)},  # red
    "empty":  {"border": (0.20, 0.47, 0.87), "fill": (0.90, 0.95, 1.00)},  # blue
}
TEXT_COLOR: tuple[float, float, float] = (0.06, 0.09, 0.16)   # near-black
WHITE:      tuple[float, float, float] = (1.00, 1.00, 1.00)


# ---- Anchor labels used to LOCATE cells --------------------------------------
CELL_ANCHOR_LABEL = "Aperture Size"     # appears once per item on every page
SURVEYOR_NAME_ANCHOR = "Name"           # 2nd match on page 1 = Surveyor slot

# ---- Hardcoded X-coordinates for the survey-value cell ----------------------
# Measured from Fenesta WCS Report template (A4 portrait, ~595pt wide).
CELL_X_LEFT:  float = 78.25     # left edge of the data-entry cell
CELL_X_RIGHT: float = 555.0     # right edge (near page margin)
CELL_INSET:   float =   2.0     # inner padding so border doesn't touch grid

# ---- Vertical geometry constants -------------------------------------------
ANCHOR_SEARCH_UP_PT:   float =  6.0   # how far above the anchor to look
ANCHOR_SEARCH_DOWN_PT: float = 30.0   # how far below the anchor to look
FALLBACK_ROW_HEIGHT:   float = 14.0   # if no rules found, use this height
MIN_MAIN_CELL_HEIGHT:  float = 13.0   # 🩹 v3.1: enforce so text always renders

# ---- Remarks row (optional, drawn immediately below the main cell) ---------
REMARKS_ROW_HEIGHT: float = 12.0
REMARKS_GAP_PT:     float =  1.5

# ---- Text sizing ------------------------------------------------------------
MIN_FONT_SIZE: float = 5.5
MAX_FONT_SIZE: float = 10.0
FONT_NAME:      str  = "helv"     # Helvetica — built-in, no embedding
FONT_NAME_BOLD: str  = "hebo"     # Helvetica-Bold

# ---- Surveyor name stamp geometry ------------------------------------------
SURVEYOR_BOX_WIDTH:  float = 140.0
SURVEYOR_BOX_HEIGHT: float =  14.0
SURVEYOR_X_OFFSET:   float =  35.0    # push right of the "Name" label
SURVEYOR_Y_OFFSET:   float =  -2.0    # tiny nudge up to sit on the underline


# =============================================================================
# PUBLIC API
# =============================================================================

def overlay_survey_data(
    pdf_bytes: bytes,
    rows: list[dict[str, Any]],
    surveyor_name: str = "",
) -> bytes:
    """
    Stamp measured survey values on top of the original order PDF.

    Args:
        pdf_bytes:     Original order PDF (as bytes).
        rows:          List of row dicts from parse_survey_pdf(). Each row is
                       stamped on the page that contains its sales-line.
        surveyor_name: Optional. Stamped on page 1 near the "Surveyor Name"
                       slot (the 2nd "Name" occurrence).

    Returns:
        bytes: The annotated PDF, ready for st.download_button().
    """
    # ---- Defensive guards --------------------------------------------------
    if not pdf_bytes or len(pdf_bytes) < 100:
        raise ValueError("overlay_survey_data: pdf_bytes is empty or truncated.")
    if rows is None:
        rows = []
    if not isinstance(rows, list):
        raise TypeError("overlay_survey_data: rows must be a list of dicts.")
    surveyor_name = (surveyor_name or "").strip()

    # Build lookup: sales_line -> row dict
    by_sales_line: dict[str, dict[str, Any]] = {
        str(r.get("sales_line", "")).strip(): r
        for r in rows
        if isinstance(r, dict) and r.get("sales_line")
    }

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Could not open PDF for overlay: {e}") from e

    try:
        if getattr(doc, "needs_pass", False) or getattr(doc, "is_encrypted", False):
            raise ValueError(
                "PDF is password-protected. Remove the password before overlaying."
            )

        # ---- (1) Surveyor name stamp on page 1 -----------------------------
        if surveyor_name and len(doc) > 0:
            _stamp_surveyor_name(doc[0], surveyor_name)

        # ---- (2) Per-row overlays ------------------------------------------
        for page_index in range(len(doc)):
            page = doc[page_index]
            _annotate_page(page, by_sales_line)

        # ---- (3) Serialize to bytes ----------------------------------------
        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        return buf.getvalue()
    finally:
        try:
            doc.close()
        except Exception:
            pass


# =============================================================================
# INTERNALS
# =============================================================================

def _stamp_surveyor_name(page: fitz.Page, name: str) -> None:
    """Whites out the surveyor "Name" slot on page 1 and inserts the name."""
    hits = page.search_for(SURVEYOR_NAME_ANCHOR)
    if not hits:
        return
    anchor_rect = hits[1] if len(hits) >= 2 else hits[0]

    box = fitz.Rect(
        anchor_rect.x1 + SURVEYOR_X_OFFSET,
        anchor_rect.y0 + SURVEYOR_Y_OFFSET,
        anchor_rect.x1 + SURVEYOR_X_OFFSET + SURVEYOR_BOX_WIDTH,
        anchor_rect.y0 + SURVEYOR_Y_OFFSET + SURVEYOR_BOX_HEIGHT,
    )
    page.draw_rect(box, color=WHITE, fill=WHITE, overlay=True)
    _insert_single_line(
        page, box, text=name.strip(),
        fontname=FONT_NAME_BOLD, color=TEXT_COLOR, align="left",
    )


def _annotate_page(page: fitz.Page, by_sales_line: dict[str, dict[str, Any]]) -> None:
    """Find each "Aperture Size" anchor and stamp the matching sales-line."""
    anchor_hits = page.search_for(CELL_ANCHOR_LABEL)
    if not anchor_hits:
        return
    sales_line_rects = _find_sales_line_rects_on_page(page, by_sales_line.keys())

    for anchor_rect in anchor_hits:
        row = _pick_row_for_anchor(anchor_rect, sales_line_rects, by_sales_line)
        if row is None:
            continue
        _draw_row_overlay(page, anchor_rect, row)


def _find_sales_line_rects_on_page(
    page: fitz.Page,
    known_sales_lines: Iterable[str],
) -> list[tuple[str, fitz.Rect]]:
    found: list[tuple[str, fitz.Rect]] = []
    for code in known_sales_lines:
        for rect in page.search_for(code):
            found.append((code, rect))
    return found


def _pick_row_for_anchor(
    anchor_rect: fitz.Rect,
    sales_line_rects: list[tuple[str, fitz.Rect]],
    by_sales_line: dict[str, dict[str, Any]],
) -> Optional[dict[str, Any]]:
    """Pick the closest sales-line ABOVE this "Aperture Size" anchor."""
    candidates = [
        (code, rect) for code, rect in sales_line_rects
        if rect.y1 <= anchor_rect.y0
    ]
    if not candidates:
        return None
    code, _ = min(candidates, key=lambda cr: anchor_rect.y0 - cr[1].y1)
    return by_sales_line.get(code)


def _draw_row_overlay(
    page: fitz.Page,
    anchor_rect: fitz.Rect,
    row: dict[str, Any],
) -> None:
    """Draw the coloured cell + text for one survey row against its anchor."""
    cell_top, cell_bot = _find_cell_bounds(page, anchor_rect)

    # 🩹 v3.1: enforce minimum cell height so text always has room to render
    if (cell_bot - cell_top) < MIN_MAIN_CELL_HEIGHT:
        cell_bot = cell_top + MIN_MAIN_CELL_HEIGHT

    main_cell = fitz.Rect(
        CELL_X_LEFT + CELL_INSET,
        cell_top + CELL_INSET,
        CELL_X_RIGHT - CELL_INSET,
        cell_bot - CELL_INSET,
    )

    status = row_tolerance(
        row.get("order_width"), row.get("order_height"),
        row.get("survey_width"), row.get("survey_height"),
    )
    colors = STATUS_COLORS.get(status, STATUS_COLORS["empty"])

    # Filled + bordered rectangle covers the pre-printed content
    page.draw_rect(
        main_cell,
        color=colors["border"],
        fill=colors["fill"],
        width=1.1,
        overlay=True,
    )

    # 🩹 v3.1: use point-based insert_text (always renders) instead of
    #        rect-based insert_textbox (silently fails on tiny rects).
    text = _format_survey_text(row)
    _insert_single_line(
        page, main_cell, text=text,
        fontname=FONT_NAME_BOLD, color=TEXT_COLOR, align="center",
    )

    # Optional remarks row underneath (uses insert_textbox — has enough height)
    remarks = str(row.get("remarks", "") or "").strip()
    if remarks:
        rem_rect = fitz.Rect(
            main_cell.x0,
            main_cell.y1 + REMARKS_GAP_PT,
            main_cell.x1,
            main_cell.y1 + REMARKS_GAP_PT + REMARKS_ROW_HEIGHT,
        )
        page.draw_rect(
            rem_rect, color=colors["border"], fill=WHITE, width=0.7, overlay=True,
        )
        _insert_single_line(
            page, rem_rect, text=f"Remarks: {remarks}",
            fontname=FONT_NAME, color=TEXT_COLOR, align="left",
        )


def _find_cell_bounds(
    page: fitz.Page,
    anchor_rect: fitz.Rect,
) -> tuple[float, float]:
    """Locate the horizontal cell-boundary lines around the anchor."""
    anchor_top    = anchor_rect.y0
    anchor_bottom = anchor_rect.y1

    horiz_rules_above: list[float] = []
    horiz_rules_below: list[float] = []

    for d in page.get_drawings():
        rect: fitz.Rect = d.get("rect")
        if rect is None:
            continue
        if rect.height >= 3.0 or rect.width <= 100.0:
            continue

        rule_y = (rect.y0 + rect.y1) / 2.0
        if (anchor_top - ANCHOR_SEARCH_UP_PT) <= rule_y <= anchor_top:
            horiz_rules_above.append(rule_y)
        elif anchor_bottom <= rule_y <= (anchor_bottom + ANCHOR_SEARCH_DOWN_PT):
            horiz_rules_below.append(rule_y)

    top = max(horiz_rules_above) if horiz_rules_above else anchor_top - 1.0
    bot = min(horiz_rules_below) if horiz_rules_below else anchor_bottom + FALLBACK_ROW_HEIGHT

    if bot <= top:
        bot = top + FALLBACK_ROW_HEIGHT

    return top, bot


# ---- Text helpers -----------------------------------------------------------

def _format_survey_text(row: dict[str, Any]) -> str:
    """Build the '{room} : {surveyed_W} x {surveyed_H}' string."""
    room = str(row.get("room", "") or "").strip()
    if not room:
        room = str(row.get("reference", "") or "").strip() or "—"

    sw = row.get("survey_width")
    sh = row.get("survey_height")

    def _fmt(v: Any) -> str:
        if v is None:
            return "___"
        try:
            import math as _math
            fv = float(v)
            if _math.isnan(fv):
                return "___"
            return f"{int(round(fv))}"
        except (TypeError, ValueError):
            return "___"

    return f"{room} : {_fmt(sw)} x {_fmt(sh)}"


def _insert_single_line(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    fontname: str = FONT_NAME,
    color: tuple[float, float, float] = TEXT_COLOR,
    align: str = "center",  # "left" | "center" | "right"
) -> None:
    """
    🩹 v3.1: Guaranteed-render single-line text.

    Uses fitz.get_text_length() to measure the string, picks a font size that
    fits inside `rect`, then places it with insert_text() at an explicit
    (x, y) baseline. Unlike insert_textbox(), insert_text() has no rect-height
    check, so text ALWAYS renders — even on very narrow cells.
    """
    if not text:
        return

    # Start with a font size that would fit the height comfortably (~72%).
    # Never exceed MAX or drop below a hard 4pt floor.
    fontsize = max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, rect.height * 0.72))

    # Measure text width; shrink font until it fits the rect width (minus pad).
    usable_width = max(rect.width - 4.0, 8.0)
    text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
    if text_width > usable_width:
        shrunk = fontsize * (usable_width / text_width)
        fontsize = max(4.0, shrunk)   # hard floor to always render *something*
        text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)

    # Horizontal alignment
    if align == "left":
        x = rect.x0 + 2.0
    elif align == "right":
        x = rect.x1 - 2.0 - text_width
    else:  # center
        x = rect.x0 + (rect.width - text_width) / 2.0

    # Vertical baseline (~75% of rect height from top gives visually centered text)
    y = rect.y0 + rect.height * 0.78

    page.insert_text(
        fitz.Point(x, y),
        text,
        fontname=fontname,
        fontsize=fontsize,
        color=color,
        overlay=True,
    )
