"""
WCS Survey Editor — Main Application (Modules 4 + 5 + 6 Polish)
---------------------------------------------------------------
End-to-end workflow:
    upload PDFs → parse → edit survey data → live tolerance dashboard →
    download annotated PDFs → aggregate summary → combined Excel export.

This file also incorporates the Module 6 polish pass:
    • Robust error handling around parsing & overlay
    • Friendly empty-state screen before any upload
    • Consistent CSS spacing / branding across every custom component
    • Graceful handling of PDFs with zero detected rows / missing metadata
"""

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st

from pdf_parser import parse_survey_pdf
from overlay import overlay_survey_data
from utils import row_tolerance

# -----------------------------------------------------------------------------
# Conditional-formatting palette for the survey grid.
#
# NOTE ON HOW THIS WORKS (important):
#   st.data_editor only applies pandas-Styler styles to *non-editable* columns,
#   and its underlying grid (glide-data-grid) only honours background-color and
#   text color — CSS borders are ignored. So we do two things:
#     1. Tint the whole row's *read-only* cells (subtle background) by tolerance.
#     2. Add a coloured left "bar" column (█) that acts like a border strip.
#   The Survey W/H cells stay editable (so Excel copy/paste + keyboard nav keep
#   working); their neighbours carry the colour, which reads as a row highlight.
#
# Thresholds (per shop-floor SOP): diff > 75 mm -> amber, diff > 200 mm -> red.
# -----------------------------------------------------------------------------
TOL_WARN_MM = 75
TOL_DANGER_MM = 200

# Subtle row-background tint applied to the read-only cells of each row.
_STATUS_BG = {
    "ok":     "background-color: #e9f7ea;",   # soft green
    "warn":   "background-color: #fff6d6;",   # soft amber
    "danger": "background-color: #ffe0e0;",   # soft red
    "empty":  "",                              # not measured -> no tint
}
# Strong colour for the left "bar" column so it reads like a border strip.
_STATUS_BAR = {
    "ok":     "color: #2e7d32; font-weight: 700;",   # green
    "warn":   "color: #c98a00; font-weight: 700;",   # amber
    "danger": "color: #c62828; font-weight: 700;",   # red
    "empty":  "color: rgba(0,0,0,0.08);",             # faint placeholder
}
_BAR_GLYPH = "█"  # the character shown in the marker column


# =============================================================================
# Page configuration
# =============================================================================
st.set_page_config(
    page_title="WCS Survey Editor",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": "WCS Survey Editor — overlay site-survey dimensions on order "
                 "PDFs and flag discrepancies against tolerance thresholds."
    },
)


# =============================================================================
# Custom CSS — audited for consistent spacing, radius, colour, shadow tokens
# =============================================================================
# Design tokens (kept as CSS variables for one-line theming):
#   --wcs-primary  #0B3D91   deep brand blue
#   --wcs-accent   #1266C1   mid brand blue
#   --wcs-radius   8-10 px   consistent card corners
#   --wcs-gap      18px      consistent vertical rhythm
#   --wcs-shadow   subtle    for card lift
CUSTOM_CSS = """
<style>
    :root {
        --wcs-primary:  #0B3D91;
        --wcs-accent:   #1266C1;
        --wcs-accent2:  #2C93E8;
        --wcs-ink:      #101828;
        --wcs-muted:    #667085;
        --wcs-border:   #E4E9F0;
        --wcs-bg-soft:  #F8FAFC;
        --wcs-radius:   8px;
        --wcs-radius-lg:10px;
        --wcs-gap:      18px;
        --wcs-shadow:   0 1px 3px rgba(16,24,40,.05);
        --wcs-shadow-h: 0 4px 10px rgba(16,24,40,.08);
    }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1400px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', system-ui, sans-serif; }

    /* -------- Slim top bar (de-emphasized — the grid is the focus) -------- */
    .wcs-header {
        background: #fff; border: 1px solid var(--wcs-border);
        border-left: 3px solid var(--wcs-primary);
        color: var(--wcs-ink); padding: 8px 16px; border-radius: 6px;
        display: flex; align-items: center; justify-content: space-between;
        margin-bottom: 10px;
    }
    .wcs-header .wcs-title { display: flex; align-items: center; gap: 8px; }
    .wcs-header .wcs-title .wcs-logo { font-size: 16px; line-height: 1; }
    .wcs-header .wcs-title h1 { font-size: 14px; margin: 0; font-weight: 600; color: var(--wcs-ink); }
    .wcs-header .wcs-title p  { display: none; }  /* tagline dropped — noise, not needed every load */
    .wcs-header .wcs-meta     { text-align: right; font-size: 11px; color: var(--wcs-muted); }
    .wcs-header .wcs-meta strong { font-size: 11px; color: var(--wcs-muted); }

    /* -------- Section titles -------- */
    .section-title {
        font-size: 13px; font-weight: 600; color: var(--wcs-muted);
        text-transform: uppercase; letter-spacing: .5px;
        margin: 14px 0 6px 0;
        display: flex; align-items: center; gap: 6px;
    }
    .section-title::before {
        content: ""; width: 3px; height: 13px; background: var(--wcs-accent2);
        border-radius: 2px; display: inline-block;
    }
    /* The grid's own title gets to be the loudest thing on the page */
    .section-title.grid-title {
        font-size: 16px; font-weight: 700; color: var(--wcs-ink);
        text-transform: none; letter-spacing: 0; margin: 6px 0 8px 0;
    }
    .section-title.grid-title::before {
        width: 4px; height: 18px; background: var(--wcs-primary);
    }

    /* -------- Compact tolerance chips (replaces the old large metric cards
       for anything that isn't the primary grid) -------- */
    .wcs-chip-row { display: flex; gap: 8px; flex-wrap: wrap; margin: 4px 0 2px 0; }
    .wcs-chip {
        display: inline-flex; align-items: center; gap: 6px;
        background: var(--wcs-bg-soft); border: 1px solid var(--wcs-border);
        border-radius: 999px; padding: 3px 10px 3px 8px;
        font-size: 12px; color: #344054;
    }
    .wcs-chip .chip-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
    .wcs-chip .chip-val { font-weight: 700; color: var(--wcs-ink); }
    .wcs-chip.green .chip-dot { background: #12B76A; }
    .wcs-chip.amber .chip-dot { background: #F79009; }
    .wcs-chip.red   .chip-dot { background: #F04438; }
    .wcs-chip.blue  .chip-dot { background: var(--wcs-accent2); }
    .wcs-chip.grey  .chip-dot { background: #98A2B3; }

    /* -------- Legacy metric card (kept only for the collapsed aggregate
       expander at the bottom — no longer used inline in the main flow) -------- */
    .metric-card {
        background: #fff; border: 1px solid var(--wcs-border);
        border-left: 4px solid var(--wcs-accent);
        border-radius: var(--wcs-radius);
        padding: 12px 14px; box-shadow: var(--wcs-shadow);
        min-height: 76px;
    }
    .metric-card .metric-label {
        color: var(--wcs-muted); font-size: 11px;
        text-transform: uppercase; letter-spacing: .6px; font-weight: 600;
    }
    .metric-card .metric-value { color: var(--wcs-ink); font-size: 21px; font-weight: 700; margin-top: 2px; }
    .metric-card .metric-sub   { color: var(--wcs-muted); font-size: 11px; margin-top: 2px; }
    .metric-card.green  { border-left-color: #12B76A; }
    .metric-card.amber  { border-left-color: #F79009; }
    .metric-card.red    { border-left-color: #F04438; }
    .metric-card.blue   { border-left-color: var(--wcs-accent2); }
    .metric-card.grey   { border-left-color: #98A2B3; }

    /* -------- Tolerance legend — one quiet caption line, not a boxed card -------- */
    .wcs-legend {
        display: flex; gap: 16px; flex-wrap: wrap;
        padding: 0; margin: 0 0 10px 0;
        font-size: 11.5px; color: var(--wcs-muted);
    }
    .wcs-legend .legend-item { display: flex; align-items: center; gap: 5px; }
    .wcs-legend .dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
    .dot.green  { background: #12B76A; }
    .dot.amber  { background: #F79009; }
    .dot.red    { background: #F04438; }
    .dot.blue   { background: var(--wcs-accent2); }
    .dot.grey   { background: #98A2B3; }

    /* -------- Uploader — minimal strip, not a big hero card once files exist -------- */
    .upload-section {
        background: transparent; border: none; padding: 0; margin: 0 0 4px 0;
    }
    .upload-section h3 { margin: 0 0 6px 0; font-size: 16px; color: var(--wcs-ink); }
    .upload-section p  { margin: 0 0 10px 0; font-size: 13px; color: var(--wcs-muted); }

    /* -------- Order metadata strip -------- */
    .order-summary {
        background: var(--wcs-bg-soft); border: 1px solid var(--wcs-border);
        border-radius: var(--wcs-radius);
        padding: 12px 16px; margin-bottom: 10px;
        display: flex; gap: 30px; flex-wrap: wrap;
    }
    .order-summary .field-label {
        font-size: 11px; color: var(--wcs-muted);
        text-transform: uppercase; letter-spacing: .5px;
    }
    .order-summary .field-value {
        font-size: 14px; color: var(--wcs-ink);
        font-weight: 600; margin-top: 2px;
    }

    /* -------- Aggregate card (now lives inside a collapsed expander) -------- */
    .aggregate-card {
        background: var(--wcs-bg-soft);
        border: 1px solid var(--wcs-border); border-radius: var(--wcs-radius);
        padding: 14px 16px; margin-top: 4px;
    }
    .aggregate-card h3 { margin: 0 0 10px 0; font-size: 14px; color: var(--wcs-ink); }

    /* -------- Empty-state hero (Module 6 polish) -------- */
    .empty-hero {
        background: #F8FAFC;
        border: 1px solid var(--wcs-border); border-radius: var(--wcs-radius);
        padding: 32px 28px; text-align: center; margin: 6px 0 16px 0;
    }
    .empty-hero .hero-icon { font-size: 40px; line-height: 1; margin-bottom: 8px; }
    .empty-hero h2 { margin: 0 0 6px 0; color: var(--wcs-ink); font-size: 18px; font-weight: 600; }
    .empty-hero p  { margin: 0 auto; color: var(--wcs-muted); font-size: 13px; max-width: 600px; }
    .empty-steps {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
        gap: 10px; margin-top: 16px;
    }
    .empty-step {
        background: #fff; border: 1px solid var(--wcs-border);
        border-radius: var(--wcs-radius); padding: 12px 14px;
        text-align: left;
    }
    .empty-step .step-num {
        display: inline-block; width: 20px; height: 20px; line-height: 20px;
        text-align: center; background: var(--wcs-accent); color: #fff;
        border-radius: 50%; font-size: 11px; font-weight: 700; margin-right: 6px;
    }
    .empty-step .step-title { font-weight: 600; color: var(--wcs-ink); font-size: 12.5px; }
    .empty-step .step-body  { color: var(--wcs-muted); font-size: 11.5px; margin-top: 3px; }

    /* -------- Grid card: the visual anchor of the page -------- */
    .grid-card {
        background: #fff; border: 1px solid var(--wcs-border);
        border-radius: var(--wcs-radius); padding: 14px 16px 6px 16px;
        box-shadow: var(--wcs-shadow); margin-bottom: 6px;
    }

    /* -------- Footer -------- */
    .wcs-footer {
        margin-top: 28px; padding: 10px 4px; border-top: 1px solid var(--wcs-border);
        color: #98A2B3; font-size: 12px; text-align: center;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# =============================================================================
# Constants
# =============================================================================
STATUSES = ("ok", "warn", "danger", "empty")
STATUS_LABEL = {
    "ok":     "Within tolerance",
    "warn":   "Borderline",
    "danger": "Out of tolerance",
    "empty":  "Not measured",
}
EDIT_COLUMNS = ("survey_width", "survey_height", "room", "remarks")


# =============================================================================
# Header, legend, sidebar
# =============================================================================
def render_header() -> None:
    st.markdown(
        """
        <div class="wcs-header">
            <div class="wcs-title">
                <div class="wcs-logo">🪟</div>
                <div><h1>WCS Survey Editor</h1></div>
            </div>
            <div class="wcs-meta">Fenesta · v0.5.0</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_legend() -> None:
    st.markdown(
        """
        <div class="wcs-legend">
            <div class="legend-item"><span class="dot green"></span> ≤75mm OK</div>
            <div class="legend-item"><span class="dot amber"></span> ≤200mm review</div>
            <div class="legend-item"><span class="dot red"></span> &gt;200mm critical</div>
            <div class="legend-item"><span class="dot blue"></span> not measured</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown("### ⚙️ Settings")

        st.text_input(
            "Surveyor Name",
            value=st.session_state.get("surveyor_name", ""),
            help="Stamped on page 1 of every annotated PDF.",
            key="surveyor_name",
        )

        st.text_input(
            "Project / Lot name",
            value=st.session_state.get("project_name", ""),
            help="Used in the combined Excel filename (e.g. 'Prestige_T3').",
            key="project_name",
            placeholder="e.g. Prestige_T3_L34",
        )

        st.markdown("---")
        st.markdown("**Tolerance thresholds** *(shop-floor SOP)*")
        st.caption("• OK ≤ 75 mm\n\n• Warn ≤ 200 mm\n\n• Danger > 200 mm")

        st.markdown("---")
        if st.button("🔄 Clear all uploads & edits", use_container_width=True):
            for k in list(st.session_state.keys()):
                if k.startswith("edited_") or k == "wcs_pdf_uploader":
                    del st.session_state[k]
            st.rerun()

        st.markdown("---")
        st.caption("WCS Survey Editor · v0.5.0")


# =============================================================================
# Small helpers
# =============================================================================
def metric_card(col, label: str, value: str, sub: str = "", css_class: str = "") -> None:
    col.markdown(
        f"""
        <div class="metric-card {css_class}">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _pct(part: int, whole: int) -> str:
    if not whole:
        return "0%"
    return f"{(part / whole) * 100:.0f}%"


def _safe(value: Any) -> str:
    """Coerce metadata to display string, guarding against None / empty."""
    if value is None:
        return "—"
    s = str(value).strip()
    return s if s else "—"


def _sanitize_sheet_name(name: str, fallback: str) -> str:
    """
    Sanitize a sheet name to Excel's rules:
      - alphanumeric + underscore only
      - max 31 chars
      - never blank
    """
    if not name:
        name = fallback
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    if not cleaned:
        cleaned = fallback
    return cleaned[:31]


def _dedupe_sheet_names(names: list[str]) -> list[str]:
    """Make sheet names unique across the workbook."""
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            suffix = f"_{seen[n]}"
            trimmed = n[: 31 - len(suffix)]
            out.append(f"{trimmed}{suffix}")
    return out


# =============================================================================
# Metadata strip
# =============================================================================
def render_metadata(metadata: dict[str, Any]) -> None:
    fields = [
        ("Order No.",  _safe(metadata.get("order_number"))),
        ("MSC No.",    _safe(metadata.get("reference_number"))),
        ("Zone",       _safe(metadata.get("zone"))),
        ("Customer",   _safe(metadata.get("customer_name"))),
        ("Quote No.",  _safe(metadata.get("quote_number"))),
        ("Order Date", _safe(metadata.get("date"))),
    ]
    html = '<div class="order-summary">'
    for label, value in fields:
        html += (
            f'<div><div class="field-label">{label}</div>'
            f'<div class="field-value">{value}</div></div>'
        )
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# =============================================================================
# Data editor
# =============================================================================
EXPECTED_COLS = [
    "marker",                       # coloured left "border" strip (read-only)
    "sales_line", "reference", "location", "description", "system",
    "order_width", "order_height",
    "survey_width", "survey_height", "room", "remarks", "status",
    "flag",
]


def rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    """Convert parser rows to a DataFrame with a status column pre-computed."""
    if not rows:
        return pd.DataFrame(columns=EXPECTED_COLS)

    df = pd.DataFrame(rows)

    # Ensure edit columns exist even if the parser omitted them
    for col in EDIT_COLUMNS:
        if col not in df.columns:
            df[col] = None if "width" in col or "height" in col else ""

    # Coerce numeric columns for the editor
    for col in ("order_width", "order_height", "survey_width", "survey_height"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["status"] = df.apply(
        lambda r: row_tolerance(
            r.get("order_width"), r.get("order_height"),
            r.get("survey_width"), r.get("survey_height"),
        ),
        axis=1,
    )

    # The coloured left "bar" — just a glyph; its colour is set by the Styler.
    df["marker"] = _BAR_GLYPH

    # Surface parsing gaps instead of leaving Ref/Location/System silently
    # blank — see pdf_parser._extract_rows / _find_subfields_anchor.
    if "subfields_missing" not in df.columns:
        df["subfields_missing"] = False
    df["subfields_missing"] = df["subfields_missing"].fillna(False).astype(bool)
    df["flag"] = df["subfields_missing"].apply(lambda m: "⚠ Check" if m else "")

    return df


def _row_status(row: pd.Series) -> str:
    """Tolerance status for a single row, reusing the shared SOP thresholds."""
    return row_tolerance(
        row.get("order_width"), row.get("order_height"),
        row.get("survey_width"), row.get("survey_height"),
    )


def _build_row_styles(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a same-shaped DataFrame of CSS strings.

    • The 'marker' column gets a strong colour  -> reads like a left border.
    • Every OTHER read-only column gets a subtle background tint.
    Editable columns (survey_width/height, room, remarks) are left blank —
    st.data_editor ignores styles on editable columns anyway, and leaving them
    un-tinted keeps the typing cells visually clean.
    """
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    editable = {"survey_width", "survey_height", "room", "remarks"}
    for i, row in df.iterrows():
        status = _row_status(row) or "empty"
        bg = _STATUS_BG.get(status, "")
        bar = _STATUS_BAR.get(status, "")
        for col in df.columns:
            if col == "marker":
                styles.at[i, col] = bar
            elif col not in editable:
                styles.at[i, col] = bg
    return styles


def render_data_editor(df: pd.DataFrame, key: str) -> pd.DataFrame:
    column_order = [c for c in EXPECTED_COLS if c in df.columns]

    # Apply conditional formatting via a pandas Styler. Styles land on the
    # read-only cells (incl. the coloured 'marker' bar); the editable Survey
    # W/H cells stay plain so copy/paste + keyboard nav feel like Excel.
    styled = df.style.apply(_build_row_styles, axis=None)

    edited = st.data_editor(
        styled,
        key=key,
        use_container_width=True,
        hide_index=True,
        column_order=column_order,
        num_rows="fixed",
        height=520,   # the grid is the main event — give it real room to breathe
        column_config={
            "marker":        st.column_config.TextColumn(
                " ", disabled=True, width="small",
                help="Tolerance flag · green = OK (≤75 mm) · "
                     "amber = review (≤200 mm) · red = critical (>200 mm)."),
            "sales_line":    st.column_config.TextColumn("Sales Line", disabled=True, width="small"),
            "reference":     st.column_config.TextColumn("Ref",        disabled=True, width="small"),
            "location":      st.column_config.TextColumn("Location",   disabled=True, width="medium"),
            "description":   st.column_config.TextColumn("Config",     disabled=True, width="medium"),
            "system":        st.column_config.TextColumn("System",     disabled=True, width="large"),
            "order_width":   st.column_config.NumberColumn("Ord W (mm)", disabled=True, format="%d", width="small"),
            "order_height":  st.column_config.NumberColumn("Ord H (mm)", disabled=True, format="%d", width="small"),
            "survey_width":  st.column_config.NumberColumn(
                "Survey W (mm)", format="%d", width="small",
                min_value=0, max_value=9999,
                help="Site-measured width in mm."),
            "survey_height": st.column_config.NumberColumn(
                "Survey H (mm)", format="%d", width="small",
                min_value=0, max_value=9999,
                help="Site-measured height in mm."),
            "room":          st.column_config.TextColumn("Room", width="medium",
                help="Friendly room label (e.g. 'Master Bedroom')."),
            "remarks":       st.column_config.TextColumn("Remarks", width="large",
                help="Free-text notes — stamped on the annotated PDF."),
            "status":        st.column_config.TextColumn("Status", disabled=True, width="small"),
            "flag":          st.column_config.TextColumn(
                "Flag", disabled=True, width="small",
                help="⚠ Check = Ref/Location/System couldn't be confidently "
                     "parsed from the PDF for this row. Verify against the "
                     "source order sheet before surveying."),
        },
    )
    return edited


# =============================================================================
# Tolerance summary
# =============================================================================
def compute_counts(df: pd.DataFrame) -> dict[str, int]:
    counts = {s: 0 for s in STATUSES}
    if df is None or df.empty:
        return counts
    for _, r in df.iterrows():
        s = row_tolerance(
            r.get("order_width"), r.get("order_height"),
            r.get("survey_width"), r.get("survey_height"),
        )
        counts[s] = counts.get(s, 0) + 1
    return counts


def render_tolerance_metrics(counts: dict[str, int], total: int) -> None:
    surveyed = total - counts["empty"]
    pct = f"{(surveyed / total * 100):.0f}%" if total else "—"

    st.markdown(
        f"""
        <div class="wcs-chip-row">
            <span class="wcs-chip green"><span class="chip-dot"></span>OK <span class="chip-val">{counts['ok']}</span></span>
            <span class="wcs-chip amber"><span class="chip-dot"></span>Review <span class="chip-val">{counts['warn']}</span></span>
            <span class="wcs-chip red"><span class="chip-dot"></span>Critical <span class="chip-val">{counts['danger']}</span></span>
            <span class="wcs-chip blue"><span class="chip-dot"></span>Not measured <span class="chip-val">{counts['empty']}</span></span>
            <span class="wcs-chip grey">Surveyed <span class="chip-val">{pct}</span> of {total}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Per-file processing
# =============================================================================
def process_file(file, idx: int) -> dict[str, Any]:
    """
    Parse one file, render its expander UI, return a small result dict for the
    aggregate/Excel sections. Never raises — errors are surfaced inline.
    """
    empty_result = {
        "name": file.name,
        "total": 0,
        "counts": {s: 0 for s in STATUSES},
        "df": pd.DataFrame(columns=EXPECTED_COLS),
        "metadata": {},
        "pdf_bytes": None,
        "parsed_ok": False,
    }

    # ---- Read & parse safely ----------------------------------------------
    try:
        pdf_bytes = file.getvalue()
    except Exception as e:
        st.error(f"❌ Could not read **{file.name}**: {e}")
        return empty_result

    if not pdf_bytes:
        st.warning(f"⚠️ **{file.name}** is empty — skipping.")
        return empty_result

    try:
        metadata, rows = parse_survey_pdf(pdf_bytes)
    except Exception as e:
        st.error(f"❌ Failed to parse **{file.name}**: {e}")
        return empty_result

    order_no = _safe(metadata.get("order_number"))
    total_rows = len(rows)
    title = (
        f"📄 {file.name}  ·  Order {order_no}  ·  "
        f"{total_rows} opening{'s' if total_rows != 1 else ''}"
    )

    with st.expander(title, expanded=(idx == 0)):
        render_metadata(metadata)

        # ---- Zero-rows guard (Module 6) -----------------------------------
        if not rows:
            st.warning(
                "🕵️ No line items were detected in this PDF.\n\n"
                "This usually means the PDF layout differs from the tuned Fenesta "
                "WCS Report template. The regex constants in `parser.py` may need "
                "adjusting — share a `page.get_text('text')` dump and we can retune."
            )
            return {**empty_result, "metadata": metadata, "pdf_bytes": pdf_bytes}

        # ---- Editable data grid — the main event ---------------------------
        editor_key = (
            f"edited_{file.file_id if hasattr(file, 'file_id') else file.name}"
        )
        df_source = rows_to_dataframe(rows)

        # Preserve prior edits across reruns
        saved = st.session_state.get(editor_key + "_df")
        if isinstance(saved, pd.DataFrame) and len(saved) == len(df_source):
            for col in EDIT_COLUMNS:
                if col in saved.columns:
                    df_source[col] = saved[col].values

        st.markdown(
            '<div class="section-title grid-title">✎ Survey Grid — enter Survey W/H, Room, Remarks</div>',
            unsafe_allow_html=True,
        )
        edited_df = render_data_editor(df_source, key=editor_key)
        st.session_state[editor_key + "_df"] = edited_df

        # ---- Live tolerance dashboard -------------------------------------
        counts = compute_counts(edited_df)
        render_tolerance_metrics(counts, total_rows)

        # ---- Annotated PDF ------------------------------------------------
        st.markdown('<div class="section-title">⬇️ Annotated PDF</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns([1, 3])
        with c1:
            build_clicked = st.button(
                "Generate annotated PDF",
                key=f"build_{editor_key}",
                type="primary", use_container_width=True,
            )
        with c2:
            st.caption(
                "Runs the overlay engine with your latest edits. "
                "You'll get a download button once it's ready."
            )

        if build_clicked:
            with st.spinner("Stamping survey values on the PDF…"):
                try:
                    updated_rows = edited_df.to_dict(orient="records")
                    annotated = overlay_survey_data(
                        pdf_bytes,
                        updated_rows,
                        surveyor_name=st.session_state.get("surveyor_name", ""),
                    )
                    st.session_state[editor_key + "_pdf"] = annotated
                    st.success("✅ Annotated PDF ready.")
                except Exception as e:
                    st.error(f"Overlay failed: {e}")

        annotated_bytes = st.session_state.get(editor_key + "_pdf")
        if annotated_bytes:
            fname = f"annotated_{order_no if order_no != '—' else 'order'}.pdf"
            st.download_button(
                label=f"📥 Download {fname}",
                data=annotated_bytes,
                file_name=fname,
                mime="application/pdf",
                key=f"dl_{editor_key}",
                use_container_width=True,
            )

    return {
        "name": file.name,
        "total": total_rows,
        "counts": counts,
        "df": edited_df,
        "metadata": metadata,
        "pdf_bytes": pdf_bytes,
        "parsed_ok": True,
    }


# =============================================================================
# Aggregate summary
# =============================================================================
def render_aggregate(results: list[dict[str, Any]]) -> None:
    if not results:
        return

    total_rows = sum(r["total"] for r in results)
    agg = {s: 0 for s in STATUSES}
    for r in results:
        for s in STATUSES:
            agg[s] += r["counts"].get(s, 0)

    surveyed = total_rows - agg["empty"]
    pct = (surveyed / total_rows) if total_rows else 0.0
    files_processed = len(results)
    files_ok = sum(1 for r in results if r.get("parsed_ok"))

    st.markdown('<div class="aggregate-card">', unsafe_allow_html=True)

    cols = st.columns(5)
    metric_card(cols[0], "Files Processed", f"{files_ok}/{files_processed}",
                "successfully parsed", "")
    metric_card(cols[1], "Total Openings", str(total_rows),
                "across all orders", "")
    metric_card(cols[2], "Within Tolerance", str(agg["ok"]),
                _pct(agg["ok"], total_rows), "green")
    metric_card(cols[3], "Discrepancies",
                str(agg["warn"] + agg["danger"]),
                f"warn: {agg['warn']} · danger: {agg['danger']}", "amber")
    metric_card(cols[4], "Not Measured", str(agg["empty"]),
                _pct(agg["empty"], total_rows), "blue")

    st.markdown("&nbsp;", unsafe_allow_html=True)
    st.markdown(
        f"**Overall survey completion:** {surveyed} / {total_rows} openings "
        f"measured ({pct * 100:.1f}%)"
    )
    st.progress(pct, text=f"{pct * 100:.1f}% surveyed")

    st.markdown('</div>', unsafe_allow_html=True)


# =============================================================================
# Module 5 — Combined Excel export
# =============================================================================
def build_combined_workbook(results: list[dict[str, Any]]) -> bytes | None:
    """
    Combine every file's edited DataFrame into ONE Excel workbook.
        • openpyxl engine
        • one sheet per file, sanitized to alphanumeric, ≤ 31 chars, de-duplicated
        • plus a leading "Summary" sheet with metadata + tolerance counts

    Returns the workbook as bytes, or None if there is nothing to export.
    """
    exportable = [r for r in results if r.get("parsed_ok") and r["total"] > 0]
    if not exportable:
        return None

    # ---- Summary rows ------------------------------------------------------
    summary_rows: list[dict[str, Any]] = []
    for r in exportable:
        md = r.get("metadata", {}) or {}
        counts = r.get("counts", {})
        surveyed = r["total"] - counts.get("empty", 0)
        completion = (surveyed / r["total"]) if r["total"] else 0.0
        summary_rows.append({
            "File":                r["name"],
            "Order No.":           _safe(md.get("order_number")),
            "MSC No.":             _safe(md.get("reference_number")),
            "Zone":                _safe(md.get("zone")),
            "Customer":            _safe(md.get("customer_name")),
            "Quote No.":           _safe(md.get("quote_number")),
            "Order Date":          _safe(md.get("date")),
            "Openings":            r["total"],
            "OK":                  counts.get("ok", 0),
            "Warn":                counts.get("warn", 0),
            "Danger":              counts.get("danger", 0),
            "Not Measured":        counts.get("empty", 0),
            "Survey Completion %": round(completion * 100, 1),
        })
    summary_df = pd.DataFrame(summary_rows)

    # ---- Sheet-name planning ----------------------------------------------
    raw_names: list[str] = []
    for i, r in enumerate(exportable, start=1):
        md = r.get("metadata", {}) or {}
        base = md.get("order_number") or r["name"].rsplit(".", 1)[0]
        raw_names.append(_sanitize_sheet_name(base, fallback=f"Order_{i}"))
    sheet_names = _dedupe_sheet_names(raw_names)

    # ---- Write the workbook ------------------------------------------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Summary first
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # One sheet per file
        for r, sheet_name in zip(exportable, sheet_names):
            df: pd.DataFrame = r["df"].copy()

            # Order the columns for readability
            preferred = [c for c in EXPECTED_COLS if c in df.columns]
            df = df[preferred]

            # Recompute status column against the latest edits
            df["status"] = df.apply(
                lambda row: row_tolerance(
                    row.get("order_width"), row.get("order_height"),
                    row.get("survey_width"), row.get("survey_height"),
                ),
                axis=1,
            )

            # Prepend the order-header block as the first two rows
            md = r.get("metadata", {}) or {}
            header_df = pd.DataFrame([
                {"": "Order No.",   " ": _safe(md.get("order_number"))},
                {"": "MSC No.",     " ": _safe(md.get("reference_number"))},
                {"": "Zone",        " ": _safe(md.get("zone"))},
                {"": "Customer",    " ": _safe(md.get("customer_name"))},
                {"": "Quote No.",   " ": _safe(md.get("quote_number"))},
                {"": "Order Date",  " ": _safe(md.get("date"))},
            ])
            header_df.to_excel(
                writer, sheet_name=sheet_name, index=False, startrow=0, header=False,
            )
            df.to_excel(
                writer, sheet_name=sheet_name, index=False, startrow=len(header_df) + 1,
            )

            # Autosize columns for that sheet
            ws = writer.sheets[sheet_name]
            for column_cells in ws.columns:
                length = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in column_cells
                )
                ws.column_dimensions[column_cells[0].column_letter].width = min(
                    max(length + 2, 10), 40
                )

        # Autosize Summary too
        ws = writer.sheets["Summary"]
        for column_cells in ws.columns:
            length = max(
                (len(str(cell.value)) if cell.value is not None else 0)
                for cell in column_cells
            )
            ws.column_dimensions[column_cells[0].column_letter].width = min(
                max(length + 2, 12), 40
            )

    return buf.getvalue()


def render_excel_export(results: list[dict[str, Any]]) -> None:
    exportable = [r for r in results if r.get("parsed_ok") and r["total"] > 0]

    st.markdown(
        '<div class="section-title">📚 Combined Excel Export — All Files</div>',
        unsafe_allow_html=True,
    )

    if not exportable:
        st.info(
            "ℹ️ Nothing to export yet. Excel export becomes available once at "
            "least one PDF is parsed with detected line items."
        )
        return

    project = _sanitize_sheet_name(
        st.session_state.get("project_name", "") or "",
        fallback="WCS_Survey",
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{project}_survey_{ts}.xlsx"

    c1, c2 = st.columns([1, 3])
    with c1:
        build_xl = st.button(
            "🧮 Build combined workbook",
            key="build_combined_xlsx",
            type="primary",
            use_container_width=True,
        )
    with c2:
        st.caption(
            f"One sheet per file (+ Summary sheet). "
            f"Filename: **{filename}**"
        )

    if build_xl:
        with st.spinner("Assembling workbook…"):
            try:
                wb_bytes = build_combined_workbook(results)
                if wb_bytes is None:
                    st.warning("No exportable rows.")
                else:
                    st.session_state["combined_xlsx"] = wb_bytes
                    st.session_state["combined_xlsx_name"] = filename
                    st.success("✅ Workbook ready.")
            except Exception as e:
                st.error(f"Excel export failed: {e}")

    xl_bytes = st.session_state.get("combined_xlsx")
    if xl_bytes:
        st.download_button(
            label=f"📥 Download {st.session_state.get('combined_xlsx_name', filename)}",
            data=xl_bytes,
            file_name=st.session_state.get("combined_xlsx_name", filename),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_combined_xlsx",
            use_container_width=True,
        )


# =============================================================================
# Empty state (Module 6 polish)
# =============================================================================
def render_empty_state() -> None:
    st.markdown(
        """
        <div class="empty-hero">
            <div class="hero-icon">🪟</div>
            <h2>Welcome to WCS Survey Editor</h2>
            <p>
                Upload one or more <strong>Fenesta WCS Report PDFs</strong> to begin.
                The app parses every opening automatically, lets your site surveyors
                enter measured dimensions in a live grid, colour-codes discrepancies
                against tolerance, stamps an annotated PDF for the shop floor, and
                exports everything to a management-ready Excel workbook.
            </p>
            <div class="empty-steps">
                <div class="empty-step">
                    <div><span class="step-num">1</span><span class="step-title">Upload PDFs</span></div>
                    <div class="step-body">One or many order PDFs. Parsing is cached.</div>
                </div>
                <div class="empty-step">
                    <div><span class="step-num">2</span><span class="step-title">Edit the grid</span></div>
                    <div class="step-body">Fill in Survey W / H, Room, and Remarks.</div>
                </div>
                <div class="empty-step">
                    <div><span class="step-num">3</span><span class="step-title">Watch tolerance</span></div>
                    <div class="step-body">Green / amber / red counts update live.</div>
                </div>
                <div class="empty-step">
                    <div><span class="step-num">4</span><span class="step-title">Download</span></div>
                    <div class="step-body">Annotated PDFs + combined Excel workbook.</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Uploader + Footer
# =============================================================================
def render_uploader():
    return st.file_uploader(
        label="📤 Upload order PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
        help="You can select multiple PDFs. Only .pdf files are accepted.",
        key="wcs_pdf_uploader",
        label_visibility="visible",
    )


def render_footer():
    st.markdown(
        '<div class="wcs-footer">'
        'WCS Survey Editor · v0.5.0 · © Manufacturing Ops'
        '</div>',
        unsafe_allow_html=True,
    )


# =============================================================================
# Main
# =============================================================================
def main() -> None:
    render_sidebar()
    render_header()
    render_legend()

    uploaded = render_uploader()

    if not uploaded:
        render_empty_state()
        render_footer()
        return

    # ---- Loop through uploaded files --------------------------------------
    results: list[dict[str, Any]] = []
    for idx, file in enumerate(uploaded):
        results.append(process_file(file, idx))

    # ---- Aggregate summary + Excel export — collapsed, out of the grid's way
    with st.expander("📈 Overall progress & export (all files)", expanded=False):
        render_aggregate(results)
        render_excel_export(results)

    render_footer()


if __name__ == "__main__":
    main()
