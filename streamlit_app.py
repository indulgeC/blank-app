# app_streamlit_build_based.py
# Front-end (Streamlit) built ON TOP OF build_crash_list.py
#
# Usage:
#   1) Put this file and build_crash_list.py in the SAME folder
#   2) pip install -r requirements_frontend.txt
#   3) streamlit run app_streamlit_build_based.py
#
# What you get:
# - Upload one/many Crash Report PDFs
# - Upload event file (csv/xlsx/xlsm)
# - Edit mappings (Manner Of Collision, Injury Severity) — NO rank, NO templates
# - Generate output with EXACT "Crash_List.xlsx" structure (same as CLI script)
# - Sort by PDF Date ascending; Crash Number restarts at 1 when year changes
# - Download formatted Excel (.xlsx)

import io
import tempfile
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import streamlit as st

# Import the existing CLI script as a module
import build_crash_list as bcl


# -----------------------------
# Helpers for event file (xlsx/xlsm)
# -----------------------------
def _col_letter_to_index(col: str) -> int:
    return bcl.excel_col_to_index(col)

def _should_skip_key(k: str) -> bool:
    if not k:
        return True
    lk = str(k).strip().lower()
    if lk in {"report number", "hsmv crash report number", "crash report number"}:
        return True
    if lk == "nan":
        return True
    return False

def load_event_lookup_any(uploaded_bytes: bytes, filename: str, sheet: Optional[str] = None) -> Dict[str, Dict[str, str]]:
    """
    Supports:
      - CSV: uses bcl.load_event_lookup logic (the same as your CLI baseline)
      - XLSX/XLSM: reads grid-like and maps by Excel letters A,B,C,AF,AG,AH,BH,BI
    """
    name = filename.lower()
    if name.endswith(".csv"):
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tf:
            tf.write(uploaded_bytes)
            tmp_path = Path(tf.name)
        try:
            return bcl.load_event_lookup(tmp_path)
        finally:
            try: tmp_path.unlink()
            except Exception: pass

    # Excel (xlsx/xlsm): read raw, map by positions
    xl = pd.ExcelFile(io.BytesIO(uploaded_bytes))
    use_sheet = sheet or xl.sheet_names[0]
    df = xl.parse(use_sheet, header=None, dtype=str)

    # Map to the same output keys as bcl.load_event_lookup returns
    mapping = {
        "Crash Year": "B",
        "Report Number": "A",
        "Date": "C",
        "Light": "AF",
        "Weather": "AG",
        "Surface Condition": "AH",
        "Crash Type": "BH",
        "Severity": "BI",
    }
    idx = {k: _col_letter_to_index(v) for k, v in mapping.items()}

    lookup: Dict[str, Dict[str, str]] = {}
    for i in range(len(df)):
        rn = df.iat[i, idx["Report Number"]] if idx["Report Number"] < df.shape[1] else None
        rn = "" if rn is None else str(rn).strip()
        if _should_skip_key(rn):
            continue

        row = {}
        for k, j in idx.items():
            if j < df.shape[1]:
                v = df.iat[i, j]
                v = "" if v is None else str(v).strip()
                if k != "Date":
                    v = bcl.normalize_condition(bcl.strip_leading_code(v))
                row[k] = v

        # Normalize Date to match bcl output format (m/d/YYYY HH:MM) if possible
        dt = bcl.parse_event_datetime(row.get("Date", ""))
        row["Date"] = dt.strftime("%-m/%-d/%Y %H:%M") if dt else row.get("Date", "")

        # Keep first occurrence
        if rn not in lookup:
            lookup[rn] = row

    return lookup

# -----------------------------
# Sorting + Crash Number rules (your latest requirement)
# -----------------------------
def build_dataframe_sorted(pdf_rows, event_lookup):
    """
    - Sort by PDF Date ascending (true datetime).
    - Crash Number resets each Crash Year.
    - Append event fields for comparison.
    - Column order matches the CLI output.
    """
    # Sort by datetime; push missing dates to the end
    def _key(r):
        d = r.get("Date")
        return (0, d) if d is not None else (1, pd.Timestamp.max.to_pydatetime())

    pdf_rows = sorted(pdf_rows, key=_key)

    # Crash Number per year (year from Crash Year field)
    counters = {}
    for r in pdf_rows:
        y = int(r["Crash Year"]) if r.get("Crash Year") else 0
        counters.setdefault(y, 0)
        counters[y] += 1
        r["Crash Number"] = counters[y]

    out_rows = []
    for r in pdf_rows:
        base = {
            "Crash Year": r.get("Crash Year", ""),
            "Crash Number": r.get("Crash Number", ""),
            "Report Number": r.get("Report Number", ""),
            "Date": bcl.format_dt_out(r.get("Date")),
            "Crash Type": r.get("Crash Type", ""),
            "Severity": r.get("Severity", ""),
            "Light": r.get("Light", ""),
            "Weather": r.get("Weather", ""),
            "Surface Condition": r.get("Surface Condition", ""),
        }

        ev = event_lookup.get(str(base["Report Number"]).strip())
        if ev:
            base.update({
                "Crash Year (event)": ev.get("Crash Year", ""),
                "Report Number (event)": ev.get("Report Number", ""),
                "Date (event)": ev.get("Date", ""),
                "Crash Type (event)": ev.get("Crash Type", ""),
                "Severity (event)": ev.get("Severity", ""),
                "Light (event)": ev.get("Light", ""),
                "Weather (event)": ev.get("Weather", ""),
                "Surface Condition (event)": ev.get("Surface Condition", ""),
            })
        else:
            base.update({
                "Crash Year (event)": "",
                "Report Number (event)": "",
                "Date (event)": "",
                "Crash Type (event)": "",
                "Severity (event)": "",
                "Light (event)": "",
                "Weather (event)": "",
                "Surface Condition (event)": "",
            })
        out_rows.append(base)

    return pd.DataFrame(out_rows)


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Crash List Builder (based on build_crash_list.py)", layout="wide")

# HBC logo at top
_logo_path = Path(__file__).parent / "assets" / "hbc-logo.png"
if _logo_path.exists():
    st.image(str(_logo_path), use_container_width=False, width=460)

st.title("Crash List Builder")

st.caption("Upload PDF + event files, generate Crash_List.xlsx, sort by Date, and reset Crash Number annually.")

col1, col2 = st.columns(2, gap="large")
with col1:
    if "crash_uploader_key" not in st.session_state:
        st.session_state.crash_uploader_key = 0
    crash_files = st.file_uploader(
        "Crash Report PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"crash_pdfs_{st.session_state.crash_uploader_key}",
    )
    if crash_files and st.button("Clear all PDFs", key="clear_crash_pdfs"):
        st.session_state.crash_uploader_key += 1
        st.rerun()

with col2:
    event_file = st.file_uploader("event file（csv/xlsx/xlsm）", type=["csv", "xlsx", "xlsm"], accept_multiple_files=False)

    sheet_name = None
    if event_file and event_file.name.lower().endswith((".xlsx", ".xlsm")):
        try:
            xl = pd.ExcelFile(io.BytesIO(event_file.getvalue()))
            sheet_name = st.selectbox("Select the sheet to read", xl.sheet_names, index=0)
        except Exception:
            sheet_name = None

st.divider()

st.subheader("Type mapping (editable）")

left, right = st.columns(2, gap="large")

with left:
    st.markdown("**Manner Of Collision (Code → Label)**")
    default_collision = [{"Code": k, "Label": v} for k, v in bcl.COLLISION_MAP.items()]
    df_collision = pd.DataFrame(default_collision).sort_values("Code")
    df_collision_edit = st.data_editor(df_collision, use_container_width=True, num_rows="dynamic")

with right:
    st.markdown("**Injury Severity (Code → Label)**")
    # Default: mirror the labels from the CLI's injury_simplify mapping (expanded)
    default_sev = [
        {"Code": 1, "Label": "No Injury"},
        {"Code": 2, "Label": "Possible Injury"},
        {"Code": 3, "Label": "Non‐incapacitating Injury"},
        {"Code": 4, "Label": "Incapacitating Injury"},
        {"Code": 5, "Label": "Fatal (within 30 days)"},
        {"Code": 6, "Label": "Non‐Traffic Fatality"},
    ]
    df_sev = pd.DataFrame(default_sev)
    df_sev_edit = st.data_editor(df_sev, use_container_width=True, num_rows="dynamic")

st.divider()

btn = st.button("Build & Download Crash_List.xlsx", type="primary", use_container_width=True)

if btn:
    if not crash_files:
        st.error("Please upload at least one Crash Report PDF.")
        st.stop()
    if not event_file:
        st.error("Please upload the event file (csv/xlsx/xlsm).")
        st.stop()

    # Patch the underlying module mappings to guarantee same parsing path + your custom labels
    new_collision = {}
    for _, r in df_collision_edit.iterrows():
        try:
            new_collision[int(r["Code"])] = str(r["Label"]).strip()
        except Exception:
            continue
    if new_collision:
        bcl.COLLISION_MAP = new_collision

    sev_map = {}
    for _, r in df_sev_edit.iterrows():
        try:
            sev_map[int(r["Code"])] = str(r["Label"]).strip()
        except Exception:
            continue

    def injury_simplify_patched(code: int) -> str:
        return sev_map.get(int(code), "") if code is not None else ""

    bcl.injury_simplify = injury_simplify_patched

    # Load event lookup
    with st.spinner("Reading event file..."):
        event_lookup = load_event_lookup_any(event_file.getvalue(), event_file.name, sheet=sheet_name)

    # Parse PDFs using the SAME extract_one_pdf logic (write to temp file so bcl works unchanged)
    pdf_rows = []
    debug = []
    with st.spinner(f"Analyze {len(crash_files)} PDF files..."):
        for f in crash_files:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                tf.write(f.getvalue())
                tmp_pdf = Path(tf.name)

            try:
                row = bcl.extract_one_pdf(tmp_pdf)
                # remove internal source column
                row.pop("_src_pdf", None)
                pdf_rows.append(row)
                debug.append({"pdf": f.name, "report": row.get("Report Number", ""), "date": str(row.get("Date", ""))})
            except Exception as e:
                st.warning(f"Analysis failed: {f.name} | {e}")
            finally:
                try: tmp_pdf.unlink()
                except Exception: pass

    if not pdf_rows:
        st.error("No PDFs were successfully analyzed. You can expand the Debug section to review the extraction status.")
        st.stop()

    # Build final DataFrame with your sorting + crash number rule
    df_out = build_dataframe_sorted(pdf_rows, event_lookup)

    st.success(f"Completed: Analysed {len(pdf_rows)} PDF files, outputting {len(df_out)} rows (sorted by Date, Crash Number reset annually)")
    st.dataframe(df_out.head(50), use_container_width=True)

    # Write formatted xlsx using the SAME writer style as CLI
    with st.spinner("Generate Excel (with formatting)..."):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            out_path = Path(tf.name)
        try:
            bcl.write_xlsx(df_out, out_path)
            xlsx_bytes = out_path.read_bytes()
        finally:
            try: out_path.unlink()
            except Exception: pass

    st.download_button(
        "Download Crash_List.xlsx",
        data=xlsx_bytes,
        file_name="Crash_List.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    with st.expander("Debug (Check whether Report Number/Date was successfully retrieved)"):
        st.dataframe(pd.DataFrame(debug), use_container_width=True)
