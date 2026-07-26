"""
overlay.py — PDF Overlay / Annotation Module (Module 3, v3.2 — geometry + style parity with app.py)
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

🩹  v3.2 fix (2026-07-26):
    Two regressions vs. the original monolithic app.py were traced back to
    THIS module and fixed:

    (a) Border placement was off. `_find_cell_bounds` capped the search for
        the surrounding horizontal ruling lines to a narrow window (6pt
        above / 30pt below the anchor) and, when nothing was found in that
        window, fell back to `anchor_top - 1.0` — almost no gap above the
        label. app.py searched the WHOLE page for ruling lines (no distance
        cap, just direction) and fell back to `hit.y0 - 13` / `hit.y1 + 3`
        when none were found. This version now matches app.py exactly:
        unbounded search, same fallback offsets. This is very likely the
        actual cause of "the border placement was better in app.py."

    (b) Color/remarks styling didn't match app.py's look:
          - app.py: WHITE fill, border AND text both colored per tolerance
            status (green/amber/red/blue) — high-contrast, severity is
            visible at a glance in both the box outline and the text color.
          - This module (pre-v3.2): pastel colored FILL, border colored,
            but text was always near-black — status was only visible in the
            box outline, not the text.
          - app.py's remarks row sits DIRECTLY below the main cell with the
            SAME height (row_h) and no "Remarks:" label prefix — just the
            remark text, colored the same as the status.
          - This module (pre-v3.2) used a fixed 12pt remarks row with a
            small gap, white fill, and a "Remarks: " prefix.
        v3.2 reverts to app.py's scheme for both, since that's the style
        that was preferred. The safer point-based/auto-fit text rendering
        from v3.1 is KEPT (it's a correctness fix, not a style choice) —
        it just now renders in app.py's colors and geometry.
"""

from __future__ import annotations

import io
from typing import Any, Iterable, Optional

import fitz  # PyMuPDF

from utils import row_tolerance

__version__ = "3.2.0"


# =============================================================================
# CONFIGURABLE CONSTANTS (tune these against your source template)
# =============================================================================

# ---- Colors (RGB, 0..1) — keyed to tolerance status -------------------------
# Matches app.py: white fill, border AND text both colored per status.
STATUS_COLORS: dict[str, tuple[float, float, float]] = {
    "ok":     (0.0,  0.6,  0.2),    # green
    "warn":   (0.8,  0.5,  0.0),    # amber
    "danger": (0.91, 0.13, 0.18),   # red   — Fenesta Red #E8212E
    "empty":  (0.0,  0.36, 0.67),   # blue  — Fenesta Blue #005BAC
}
WHITE: tuple[float, float, float] = (1.0, 1.0, 1.0)


# ---- Anchor labels used to LOCATE cells --------------------------------------
CELL_ANCHOR_LABEL = "Aperture Size"     # appears once per item on every page
SURVEYOR_NAME_ANCHOR = "Name"           # 2nd match on page 1 = Surveyor slot

# ---- Hardcoded X-coordinates for the survey-value cell ----------------------
# Measured from Fenesta WCS Report template (A4 portrait, ~595pt wide).
# Reverted to app.py's measured value (459.65) — the wider 555.0 used
# previously extended past the data-entry cell into unrelated content.
CELL_X_LEFT:  float = 78.25     # left edge of the data-entry cell
CELL_X_RIGHT: float = 459.65    # right edge of the data-entry cell
CELL_INSET:   float =   0.5     # matches app.py's +0.5/-0.5 border inset

# ---- Vertical geometry constants -------------------------------------------
# app.py searched the WHOLE page for ruling lines near the anchor (direction
# only, no distance cap) and used fixed fallback offsets when none were
# found. Replicated exactly here — see v3.2 fix note above.
ANCHOR_RULE_BUFFER_PT: float =  2.0   # matches app.py's "+2 / -2" buffer
FALLBACK_TOP_OFFSET_PT: float = 13.0  # top = anchor_top - 13   (app.py)
FALLBACK_BOT_OFFSET_PT: float =  3.0  # bot = anchor_bottom + 3 (app.py)
MIN_MAIN_CELL_HEIGHT:  float = 13.0   # 🩹 v3.1: enforce so text always renders

# ---- Text sizing ------------------------------------------------------------
# app.py: fs = min(9.0, row_h * 0.62) — same formula used here as the
# starting point; the v3.1 width-based auto-shrink still applies underneath
# as a safety net for unusually long remarks/room strings.
FONT_SIZE_ROW_HEIGHT_RATIO: float = 0.62
MAX_FONT_SIZE: float = 9.0
MIN_FONT_SIZE: float = 5.5
FONT_NAME:      str  = "helv"     # Helvetica — built-in, no embedding
FONT_NAME_BOLD: str  = "hebo"     # Helvetica-Bold

# ---- Surveyor name stamp geometry ------------------------------------------
SURVEYOR_BOX_WIDTH:  float = 140.0
SURVEYOR_BOX_HEIGHT: float =  14.0
SURVEYOR_X_OFFSET:   float =  35.0    # push right of the "Name" label
SURVEYOR_Y_OFFSET:   float =  -2.0    # tiny nudge up to sit on the underline
SURVEYOR_TEXT_COLOR: tuple[float, float, float] = (0.06, 0.09, 0.16)  # near-black


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
        fontname=FONT_NAME_BOLD, color=SURVEYOR_TEXT_COLOR, align="left",
    )


def _annotate_page(page: fitz.Page, by_sales_line: dict[str, dict[str, Any]]) -> None:
    """Find each "Aperture Size" anchor and stamp the matching sales-line."""
    anchor_hits = page.search_for(CELL_ANCHOR_LABEL)
    if not anchor_hits:
        return
    sales_line_rects = _find_sales_line_rects_on_page(page, by_sales_line.keys())

    # Precompute the page's horizontal ruling lines once (matches app.py's
    # per-page h_lines pass) instead of re-scanning get_drawings() per row.
    h_lines = _find_horizontal_rules(page)

    for anchor_rect in anchor_hits:
        row = _pick_row_for_anchor(anchor_rect, sales_line_rects, by_sales_line)
        if row is None:
            continue
        _draw_row_overlay(page, anchor_rect, row, h_lines)


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
    h_lines: list[fitz.Rect],
) -> None:
    """Draw the coloured cell + text for one survey row against its anchor.

    Styling matches app.py: WHITE fill, border AND text both colored per
    tolerance status. The remarks row (if any) sits directly below the main
    cell with the SAME height, same styling, no label prefix — matching
    app.py's "Production Size row" behaviour exactly.
    """
    cell_top, cell_bot = _find_cell_bounds(anchor_rect, h_lines)

    # 🩹 v3.1: enforce minimum cell height so text always has room to render
    if (cell_bot - cell_top) < MIN_MAIN_CELL_HEIGHT:
        cell_bot = cell_top + MIN_MAIN_CELL_HEIGHT

    row_h = cell_bot - cell_top

    status = row_tolerance(
        row.get("order_width"), row.get("order_height"),
        row.get("survey_width"), row.get("survey_height"),
    )
    color = STATUS_COLORS.get(status, STATUS_COLORS["empty"])

    # ---- Main "Aperture Size" cell: white fill, colored border + text -----
    main_cell = fitz.Rect(
        CELL_X_LEFT + CELL_INSET,
        cell_top + CELL_INSET,
        CELL_X_RIGHT - CELL_INSET,
        cell_bot - CELL_INSET,
    )
    page.draw_rect(main_cell, color=color, fill=WHITE, width=1.5, overlay=True)

    text = _format_survey_text(row)
    _insert_single_line(
        page, main_cell, text=text, row_h=row_h,
        fontname=FONT_NAME, color=color, align="left",
    )

    # ---- Remarks row: directly below, SAME height, no gap, no label ------
    # (mirrors app.py's "Production Size row" — plain remark text, same
    # color as the status, same left-aligned layout as the main cell)
    remarks = str(row.get("remarks", "") or "").strip()
    if remarks:
        rem_top = cell_bot
        rem_bot = cell_bot + row_h
        rem_cell = fitz.Rect(
            CELL_X_LEFT + CELL_INSET,
            rem_top + CELL_INSET,
            CELL_X_RIGHT - CELL_INSET,
            rem_bot - CELL_INSET,
        )
        page.draw_rect(rem_cell, color=color, fill=WHITE, width=1.5, overlay=True)
        _insert_single_line(
            page, rem_cell, text=remarks, row_h=row_h,
            fontname=FONT_NAME, color=color, align="left",
        )


def _find_horizontal_rules(page: fitz.Page) -> list[fitz.Rect]:
    """
    Collect every thin, wide horizontal ruling line on the page — the same
    filter app.py used (height < 3pt, width > 100pt) — once per page.
    """
    rules: list[fitz.Rect] = []
    for d in page.get_drawings():
        rect: fitz.Rect = d.get("rect")
        if rect is None:
            continue
        if abs(rect.y1 - rect.y0) < 3 and (rect.x1 - rect.x0) > 100:
            rules.append(rect)
    return rules


def _find_cell_bounds(
    anchor_rect: fitz.Rect,
    h_lines: list[fitz.Rect],
) -> tuple[float, float]:
    """
    Locate the horizontal cell-boundary lines around the anchor.

    🩹 v3.2: matches app.py exactly — search the WHOLE set of ruling lines on
    the page (no distance cap, direction only), and use the same fallback
    offsets (-13 above / +3 below) when no ruling line is found. The
    previous version capped the search window and used a near-zero fallback
    gap, which misplaced the border whenever the true ruling line fell
    outside that window.
    """
    above = [r for r in h_lines if r.y1 <= anchor_rect.y0 + ANCHOR_RULE_BUFFER_PT]
    below = [r for r in h_lines if r.y0 >= anchor_rect.y1 - ANCHOR_RULE_BUFFER_PT]

    # app.py picks the CLOSEST line above (max y0) and CLOSEST line below
    # (min y0); fall back to the fixed offsets when no ruling line is found.
    top = max((r.y0 for r in above), default=anchor_rect.y0 - FALLBACK_TOP_OFFSET_PT)
    bot = min((r.y0 for r in below), default=anchor_rect.y1 + FALLBACK_BOT_OFFSET_PT)

    return top, bot


# ---- Text helpers -----------------------------------------------------------

def _format_survey_text(row: dict[str, Any]) -> str:
    """
    Build the '{room} : {surveyed_W} x {surveyed_H}' string.

    Matches app.py exactly: if room is blank, show the size alone (no
    placeholder room text); if only one dimension is measured, show '--'
    for the missing one; if neither is measured, show 'Not surveyed'.
    """
    room = str(row.get("room", "") or "").strip()

    sw = row.get("survey_width")
    sh = row.get("survey_height")

    def _has_value(v: Any) -> bool:
        if v is None:
            return False
        try:
            import math as _math
            return not _math.isnan(float(v))
        except (TypeError, ValueError):
            return False

    def _fmt(v: Any) -> str:
        return str(int(round(float(v))))

    if _has_value(sw) and _has_value(sh):
        size_txt = f"{_fmt(sw)} x {_fmt(sh)}"
    elif _has_value(sw):
        size_txt = f"{_fmt(sw)} x --"
    elif _has_value(sh):
        size_txt = f"-- x {_fmt(sh)}"
    else:
        size_txt = "Not surveyed"

    return f"{room} : {size_txt}" if room else size_txt


def _insert_single_line(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    row_h: Optional[float] = None,
    fontname: str = FONT_NAME,
    color: tuple[float, float, float] = STATUS_COLORS["empty"],
    align: str = "left",  # "left" | "center" | "right"
) -> None:
    """
    Guaranteed-render single-line text (🩹 v3.1 fix kept), sized using
    app.py's formula: fs = min(MAX_FONT_SIZE, row_h * 0.62).

    Uses fitz.get_text_length() to measure the string and shrinks further if
    needed so it never overflows the cell width, then places it with
    insert_text() at an explicit (x, y) baseline. Unlike insert_textbox(),
    insert_text() has no rect-height check, so text ALWAYS renders — even
    on the ~6-8pt tall cells this template uses.
    """
    if not text:
        return

    ref_height = row_h if row_h is not None else rect.height
    fontsize = max(
        MIN_FONT_SIZE,
        min(MAX_FONT_SIZE, ref_height * FONT_SIZE_ROW_HEIGHT_RATIO),
    )

    # Measure text width; shrink font until it fits the rect width (minus pad).
    usable_width = max(rect.width - 4.0, 8.0)
    text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
    if text_width > usable_width:
        shrunk = fontsize * (usable_width / text_width)
        fontsize = max(4.0, shrunk)   # hard floor to always render *something*
        text_width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)

    # Horizontal alignment
    if align == "left":
        x = rect.x0 + 4.0   # matches app.py's CELL_PAD
    elif align == "right":
        x = rect.x1 - 4.0 - text_width
    else:  # center
        x = rect.x0 + (rect.width - text_width) / 2.0

    # Vertical baseline — app.py: top_y + row_h * 0.73
    y = rect.y0 + ref_height * 0.73

    page.insert_text(
        fitz.Point(x, y),
        text,
        fontname=fontname,
        fontsize=fontsize,
        color=color,
        overlay=True,
    )
