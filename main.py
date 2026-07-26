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
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px; }
    html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', system-ui, sans-serif; }

    /* -------- Branded Header -------- */
    .wcs-header {
        background: linear-gradient(90deg, var(--wcs-primary) 0%, var(--wcs-accent) 60%, var(--wcs-accent2) 100%);
        color: #fff; padding: 18px 26px; border-radius: var(--wcs-radius-lg);
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 4px 14px rgba(11,61,145,.18);
        margin-bottom: var(--wcs-gap);
    }
    .wcs-header .wcs-title { display: flex; align-items: center; gap: 14px; }
    .wcs-header .wcs-title .wcs-logo { font-size: 34px; line-height: 1; }
    .wcs-header .wcs-title h1 { font-size: 22px; margin: 0; font-weight: 600; letter-spacing: .3px; }
    .wcs-header .wcs-title p  { margin: 2px 0 0 0; font-size: 13px; opacity: .88; }
    .wcs-header .wcs-meta     { text-align: right; font-size: 12px; opacity: .9; }
    .wcs-header .wcs-meta strong { font-size: 13px; }

    /* -------- Section titles (uniform across the page) -------- */
    .section-title {
        font-size: 15px; font-weight: 600; color: var(--wcs-ink);
        margin: var(--wcs-gap) 0 8px 0;
        display: flex; align-items: center; gap: 8px;
    }
    .section-title::before {
        content: ""; width: 4px; height: 18px; background: var(--wcs-accent);
        border-radius: 2px; display: inline-block;
    }

    /* -------- Metric Cards (uniform padding + min height) -------- */
    .metric-card {
        background: #fff; border: 1px solid var(--wcs-border);
        border-left: 4px solid var(--wcs-accent);
        border-radius: var(--wcs-radius);
        padding: 14px 16px; box-shadow: var(--wcs-shadow);
        transition: transform .15s ease, box-shadow .15s ease;
        min-height: 92px;
    }
    .metric-card:hover { transform: translateY(-1px); box-shadow: var(--wcs-shadow-h); }
    .metric-card .metric-label {
        color: var(--wcs-muted); font-size: 12px;
        text-transform: uppercase; letter-spacing: .6px; font-weight: 600;
    }
    .metric-card .metric-value { color: var(--wcs-ink); font-size: 26px; font-weight: 700; margin-top: 4px; }
    .metric-card .metric-sub   { color: var(--wcs-muted); font-size: 12px; margin-top: 2px; }
    .metric-card.green  { border-left-color: #12B76A; }
    .metric-card.amber  { border-left-color: #F79009; }
    .metric-card.red    { border-left-color: #F04438; }
    .metric-card.blue   { border-left-color: var(--wcs-accent2); }
    .metric-card.grey   { border-left-color: #98A2B3; }

    /* -------- Tolerance Legend -------- */
    .wcs-legend {
        display: flex; gap: 22px; flex-wrap: wrap;
        background: var(--wcs-bg-soft); border: 1px solid var(--wcs-border);
        border-radius: var(--wcs-radius);
        padding: 10px 16px; margin: 10px 0 var(--wcs-gap) 0;
        font-size: 13px; color: #344054;
    }
    .wcs-legend .legend-item { display: flex; align-items: center; gap: 8px; }
    .wcs-legend .dot { width: 12px; height: 12px; border-radius: 50%;
        display: inline-block;
        box-shadow: 0 0 0 2px rgba(255,255,255,.9), 0 0 0 3px rgba(16,24,40,.06); }
    .dot.green  { background: #12B76A; }
    .dot.amber  { background: #F79009; }
    .dot.red    { background: #F04438; }
    .dot.blue   { background: var(--wcs-accent2); }
    .dot.grey   { background: #98A2B3; }

    /* -------- Uploader Section -------- */
    .upload-section {
        background: #fff; border: 1px dashed #B8C4D6;
        border-radius: var(--wcs-radius-lg);
        padding: 18px 20px; margin-top: 6px;
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

    /* -------- Aggregate card -------- */
    .aggregate-card {
        background: linear-gradient(135deg, #EEF4FF 0%, var(--wcs-bg-soft) 100%);
        border: 1px solid #B8C4D6; border-radius: var(--wcs-radius-lg);
        padding: 18px 22px; margin-top: 24px;
    }
    .aggregate-card h3 { margin: 0 0 12px 0; font-size: 16px; color: var(--wcs-primary); }

    /* -------- Empty-state hero (Module 6 polish) -------- */
    .empty-hero {
        background: linear-gradient(135deg, #F8FAFC 0%, #EEF4FF 100%);
        border: 1px solid var(--wcs-border); border-radius: var(--wcs-radius-lg);
        padding: 42px 32px; text-align: center; margin: 8px 0 22px 0;
    }
    .empty-hero .hero-icon { font-size: 56px; line-height: 1; margin-bottom: 10px; }
    .empty-hero h2 { margin: 0 0 6px 0; color: var(--wcs-primary); font-size: 22px; }
    .empty-hero p  { margin: 0 auto; color: var(--wcs-muted); font-size: 14px; max-width: 620px; }
    .empty-steps {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
        gap: 14px; margin-top: 22px;
    }
    .empty-step {
        background: #fff; border: 1px solid var(--wcs-border);
        border-radius: var(--wcs-radius); padding: 14px 16px;
        text-align: left; box-shadow: var(--wcs-shadow);
    }
    .empty-step .step-num {
        display: inline-block; width: 22px; height: 22px; line-height: 22px;
        text-align: center; background: var(--wcs-accent); color: #fff;
        border-radius: 50%; font-size: 12px; font-weight: 700; margin-right: 6px;
    }
    .empty-step .step-title { font-weight: 600; color: var(--wcs-ink); font-size: 13px; }
    .empty-step .step-body  { color: var(--wcs-muted); font-size: 12px; margin-top: 4px; }

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
                <div>
                    <h1>WCS Survey Editor</h1>
                    <p>Overlay site-survey dimensions on order PDFs · Flag discrepancies against tolerances</p>
                </div>
            </div>
            <div class="wcs-meta">
                <div><strong>Fenesta</strong> · Manufacturing Ops</div>
                <div>v0.5.0 · Polish pass</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_legend() -> None:
    st.markdown(
        """
        <div class="wcs-legend">
            <div class="legend-item"><span class="dot green"></span> Within tolerance (≤ 75 mm)</div>
            <div class="legend-item"><span class="dot amber"></span> Borderline (≤ 200 mm)</div>
            <div class="legend-item"><span class="dot red"></span> Out of tolerance (&gt; 200 mm)</div>
            <div class="legend-item"><span class="dot blue"></span> Not yet measured</div>
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

    # Surface parsing gaps instead of leaving Ref/Location/System silently
    # blank — see pdf_parser._extract_rows / _find_subfields_anchor.
    if "subfields_missing" not in df.columns:
        df["subfields_missing"] = False
    df["subfields_missing"] = df["subfields_missing"].fillna(False).astype(bool)
    df["flag"] = df["subfields_missing"].apply(lambda m: "⚠ Check" if m else "")

    return df


def render_data_editor(df: pd.DataFrame, key: str) -> pd.DataFrame:
    column_order = [c for c in EXPECTED_COLS if c in df.columns]
    edited = st.data_editor(
        df,
        key=key,
        use_container_width=True,
        hide_index=True,
        column_order=column_order,
        num_rows="fixed",
        column_config={
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
    st.markdown('<div class="section-title">📊 Tolerance Summary</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    surveyed = total - counts["empty"]
    pct = f"{(surveyed / total * 100):.0f}%" if total else "—"

    metric_card(cols[0], "Within Tolerance", str(counts["ok"]),
                f"of {total} openings", "green")
    metric_card(cols[1], "Borderline", str(counts["warn"]),
                "review required", "amber")
    metric_card(cols[2], "Out of Tolerance", str(counts["danger"]),
                "action required", "red")
    metric_card(cols[3], "Not Measured", str(counts["empty"]),
                f"surveyed: {pct}", "blue")


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

        # ---- Editable data grid -------------------------------------------
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

    st.markdown(
        '<div class="section-title">📈 Aggregate Summary — All Uploaded Files</div>',
        unsafe_allow_html=True,
    )
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
    st.markdown('<div class="section-title">📤 Upload Order PDFs</div>',
                unsafe_allow_html=True)
    st.markdown(
        """
        <div class="upload-section">
            <h3>Drop one or more order PDFs</h3>
            <p>Each PDF is parsed automatically. Fill in the surveyed
               width / height / room / remarks in the data grid — the tolerance
               summary, annotated PDF, and Excel export all update live.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    return st.file_uploader(
        label="Select PDF file(s)",
        type=["pdf"],
        accept_multiple_files=True,
        help="You can select multiple PDFs. Only .pdf files are accepted.",
        key="wcs_pdf_uploader",
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

    st.success(
        f"✅ {len(uploaded)} file{'s' if len(uploaded) != 1 else ''} received. "
        "Parsing…"
    )

    # ---- Loop through uploaded files --------------------------------------
    results: list[dict[str, Any]] = []
    for idx, file in enumerate(uploaded):
        results.append(process_file(file, idx))

    # ---- Aggregate summary ------------------------------------------------
    render_aggregate(results)

    # ---- Module 5 — Combined Excel export ---------------------------------
    render_excel_export(results)

    render_footer()


if __name__ == "__main__":
    main()
