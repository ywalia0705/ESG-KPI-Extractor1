import streamlit as st
import pandas as pd
import pdfplumber
import docx2txt
import tempfile
import re
import io
from pathlib import Path
from typing import Tuple, Dict, List, Any

# Optional: python-docx for extracting tables from .docx (fallback)
try:
    from docx import Document
    HAS_PYDOCX = True
except Exception:
    HAS_PYDOCX = False

st.set_page_config(page_title="ESG KPI Extraction Tool — v2", layout="wide")
st.title("📊 ESG KPI Extraction Tool (MVP) — v2")
st.caption("Upload a sustainability/climate/GRI/TCFD/TNFD report → select framework → extract KPIs → download table.")

# ---------------------------
# KPI dictionaries (starter set; expand anytime)
# ---------------------------
KPI_DICTIONARIES = {
    "TCFD": {
        "Scope 1 emissions": [r"\bScope\s*1\b.*?(?:GHG|emissions)", r"\bdirect\s+emissions\b"],
        "Scope 2 emissions": [r"\bScope\s*2\b.*?(?:GHG|emissions)", r"\bindirect\s+emissions\b"],
        "Scope 3 emissions": [r"\bScope\s*3\b.*?(?:GHG|emissions)"],
        "Board oversight of climate": [r"\bBoard\b.*\boversight\b.*\bclimate\b", r"\bgovernance\b.*\bclimate\b"],
        "Scenario analysis disclosed": [r"\bscenario analysis\b", r"\bclimate scenarios?\b"],
    },
    "GRI": {
        "Energy consumption": [r"\benergy\s+consumption\b", r"\btotal\s+energy\s+(?:use|consumed)\b"],
        "Water withdrawal": [r"\bwater\s+withdrawal\b", r"\bwater\s+use(?:age)?\b"],
        "Waste generated": [r"\bwaste\s+generated\b", r"\btotal\s+waste\b"],
        "GHG emissions (total)": [r"\bGHG\s+emissions\b", r"\bgreenhouse\s+gas\b.*\bemissions\b"],
    },
    "TNFD": {
        "Biodiversity impact": [r"\bbiodiversity\s+impact\b", r"\becosystem\b.*\bimpact\b"],
        "Land use / habitat": [r"\bland\s+use\b", r"\bhabitat\s+(?:loss|conversion)\b"],
        "Natural capital dependency": [r"\bnatural\s+capital\s+dependenc(?:y|ies)\b", r"\bnature[-\s]?related\s+dependencies\b"],
        "Water stress exposure": [r"\bwater\s+stress\b", r"\bhigh\s+baseline\s+water\s+stress\b"],
    }
}

# Units we try to capture after numbers (non-exhaustive; add as needed)
UNIT_LIST = [
    "tCO2e", "CO2e", "kgCO2e", "kt", "t", "tonnes", "tonne", "mt",
    "m3", "m³", "kWh", "MWh", "GWh", "GJ", "TJ", "%", "ppm",
    "ha", "km2", "km²", "L", "litres", "liters", "USD", "EUR",
    "GBP", "₹", "€", "£"
]
UNIT_REGEX = r"(?:" + r"|".join(re.escape(u) for u in UNIT_LIST) + r")"
NUMBER_PATTERN = rf"(-?\d{{1,3}}(?:,\d{{3}})*|\d+)(?:\.\d+)?\s*(?P<unit>{UNIT_REGEX})?"

# Some words we treat as 'negative context' to reduce confidence (e.g., in refs/appendix)
NEGATIVE_SECTIONS = ["references", "bibliography", "appendix", "annex", "table of contents", "contents"]

# ---------------------------
# Helpers
# ---------------------------

def _save_to_tmp(uploaded_file, suffix: str):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.flush()
    return tmp.name


def normalize_text(text: str) -> str:
    if not text:
        return ""
    # Remove hyphenation at line breaks (e.g. 'emiss-\nions' -> 'emissions')
    text = re.sub(r"-\n\s*", "", text)
    # Replace newlines with spaces but keep paragraph breaks short
    text = re.sub(r"\n+", " ", text)
    # Normalize unicode minus and multiple spaces
    text = text.replace('\u2212', '-')
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


@st.cache_data(show_spinner=False)
def extract_text_pdf(uploaded_file) -> Tuple[Dict[int, str], Dict[int, List[pd.DataFrame]]]:
    path = _save_to_tmp(uploaded_file, ".pdf")
    text_by_page = {}
    tables_by_page = {}
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages):
                raw = page.extract_text() or ""
                text_by_page[i + 1] = normalize_text(raw)

                # try to extract tables (simple approach)
                tables = []
                try:
                    raw_tables = page.extract_tables() or []
                    for t in raw_tables:
                        df = pd.DataFrame(t)
                        # drop empty rows/cols
                        if df.dropna(how='all').shape[0] > 0:
                            tables.append(df)
                except Exception:
                    tables = []
                tables_by_page[i + 1] = tables
    finally:
        Path(path).unlink(missing_ok=True)
    return text_by_page, tables_by_page


@st.cache_data(show_spinner=False)
def extract_text_docx(uploaded_file) -> Tuple[Dict[int, str], Dict[int, List[pd.DataFrame]]]:
    path = _save_to_tmp(uploaded_file, ".docx")
    text = docx2txt.process(path) or ""
    pages_text = {1: normalize_text(text)}

    tables_by_page = {1: []}
    # Try python-docx for extracting tables if available
    if HAS_PYDOCX:
        try:
            doc = Document(path)
            tables = []
            for tbl in doc.tables:
                data = [[cell.text for cell in row.cells] for row in tbl.rows]
                df = pd.DataFrame(data)
                if df.dropna(how='all').shape[0] > 0:
                    tables.append(df)
            tables_by_page[1] = tables
        except Exception:
            tables_by_page[1] = []
    Path(path).unlink(missing_ok=True)
    return pages_text, tables_by_page


def find_values_near(text: str, match_start: int, match_end: int, window_after=220, window_before=80):
    """Find nearest number to the matched KPI phrase by absolute character distance.
    Returns (num_str, unit_str, raw_match_span_tuple)
    """
    left_start = max(0, match_start - window_before)
    right_end = min(len(text), match_end + window_after)
    window = text[left_start:right_end]
    matches = list(re.finditer(NUMBER_PATTERN, window, flags=re.IGNORECASE))
    if not matches:
        return None, None, None
    match_center = (match_start + match_end) / 2
    # compute absolute distance from the match center to the number center
    best = None
    best_dist = None
    for m in matches:
        num_start = left_start + m.start()
        num_end = left_start + m.end()
        num_center = (num_start + num_end) / 2
        dist = abs(num_center - match_center)
        if best is None or dist < best_dist:
            best = m
            best_dist = dist
    value = best.group(0).strip()
    unit = best.group('unit') or ""
    # remove unit from numeric string for cleanliness
    num = value
    if unit:
        # strip trailing unit
        num = re.sub(re.escape(unit) + r"\s*$", "", num, flags=re.IGNORECASE).strip()
    return num, unit, (left_start + best.start(), left_start + best.end())


def compute_confidence(evidence: str, number_found: bool, unit_found: bool, table_hit: bool) -> int:
    score = 25
    if number_found:
        score += 40
    if unit_found:
        score += 15
    if table_hit:
        score += 10
    # short heuristics: presence of strong keywords
    if re.search(r"\b(emissions|energy|GHG|tonnes|tCO2e|water|waste|biodiversity)\b", evidence, flags=re.IGNORECASE):
        score += 10
    # penalize if in references/appendix-like context
    if any(ns in evidence.lower() for ns in NEGATIVE_SECTIONS):
        score = max(0, score - 60)
    return max(0, min(100, score))


def search_kpis_in_text(text_by_page: Dict[int, str], framework: str) -> pd.DataFrame:
    out = []
    patterns = KPI_DICTIONARIES.get(framework, {})
    for page, text in text_by_page.items():
        if not text:
            continue
        for kpi_name, pat_list in patterns.items():
            for pat in pat_list:
                try:
                    for m in re.finditer(pat, text, flags=re.IGNORECASE | re.DOTALL):
                        val, unit, num_span = find_values_near(text, m.start(), m.end())
                        evidence = text[max(0, m.start()-80): m.end()+160]
                        # small negative filter to reduce false positives
                        if any(x in evidence.lower() for x in ["table of contents", "references", "bibliography"]):
                            continue
                        confidence = compute_confidence(evidence, bool(val), bool(unit), table_hit=False)
                        out.append({
                            "Framework": framework,
                            "KPI Name": kpi_name,
                            "Value": val or "",
                            "Unit": unit or "",
                            "Source Page": page,
                            "Source": "text",
                            "Evidence": evidence,
                            "Confidence": confidence
                        })
                except re.error:
                    # skip invalid regex
                    continue
    if out:
        df = pd.DataFrame(out)
        df = df.sort_values(["KPI Name", "Source Page", "Confidence"], ascending=[True, True, False])
        # keep the best per KPI+Page
        df = df.drop_duplicates(subset=["KPI Name", "Source Page"], keep="first")
        return df
    return pd.DataFrame(columns=["Framework","KPI Name","Value","Unit","Source Page","Source","Evidence","Confidence"])


def search_kpis_in_tables(tables_by_page: Dict[int, List[pd.DataFrame]], framework: str) -> pd.DataFrame:
    out = []
    patterns = KPI_DICTIONARIES.get(framework, {})
    for page, tables in tables_by_page.items():
        for ti, df in enumerate(tables):
            # normalize all cells to string
            df_str = df.fillna("").astype(str)
            # create a flattened text blob for fuzzy searching
            table_text = " ".join(df_str.values.ravel())
            for kpi_name, pat_list in patterns.items():
                for pat in pat_list:
                    for m in re.finditer(pat, table_text, flags=re.IGNORECASE | re.DOTALL):
                        # Try to find exact cell containing match
                        found_cell = None
                        for r in range(df_str.shape[0]):
                            for c in range(df_str.shape[1]):
                                if re.search(pat, df_str.iat[r, c], flags=re.IGNORECASE):
                                    found_cell = (r, c)
                                    break
                            if found_cell:
                                break
                        # Look for numeric nearby (same row, rightwards up to 3 cols)
                        val = None
                        unit = None
                        if found_cell is not None:
                            r, c = found_cell
                            for c2 in range(c, min(df_str.shape[1], c+4)):
                                cell = df_str.iat[r, c2]
                                mnum = re.search(NUMBER_PATTERN, cell, flags=re.IGNORECASE)
                                if mnum:
                                    val = mnum.group(0).strip()
                                    unit = mnum.group('unit') or ""
                                    break
                        # fallback: search entire table blob
                        if not val:
                            mnum = re.search(NUMBER_PATTERN, table_text, flags=re.IGNORECASE)
                            if mnum:
                                val = mnum.group(0).strip()
                                unit = mnum.group('unit') or ""

                        evidence = df_str.head(6).to_csv(index=False)  # simple snippet of table for evidence
                        confidence = compute_confidence(evidence, bool(val), bool(unit), table_hit=True)
                        out.append({
                            "Framework": framework,
                            "KPI Name": kpi_name,
                            "Value": val or "",
                            "Unit": unit or "",
                            "Source Page": page,
                            "Source": f"table_{ti}",
                            "Evidence": evidence,
                            "Confidence": confidence
                        })
    if out:
        df = pd.DataFrame(out)
        df = df.sort_values(["KPI Name","Source Page","Confidence"], ascending=[True, True, False])
        df = df.drop_duplicates(subset=["KPI Name","Source Page","Source"], keep="first")
        return df
    return pd.DataFrame(columns=["Framework","KPI Name","Value","Unit","Source Page","Source","Evidence","Confidence"])


# ---------------------------
# UI
# ---------------------------
with st.sidebar:
    st.header("⚙️ Settings")
    framework = st.selectbox("Report framework", ["TCFD", "GRI", "TNFD"])
    min_confidence = st.slider("Minimum confidence to display", 0, 100, 35)
    extract_tables = st.checkbox("Extract tables from report (PDF/DOCX)", value=True)
    search_tables = st.checkbox("Search KPI patterns inside extracted tables", value=True)
    show_evidence = st.checkbox("Show evidence text", value=False)
    highlight = st.checkbox("Highlight match in evidence (if shown)", value=True)
    st.markdown("---")
    st.caption("Tip: For scanned PDFs (images), text extraction may be blank; OCR (pytesseract) can be added later.")

uploaded = st.file_uploader("Upload report (PDF or DOCX)", type=["pdf", "docx"]) 

if uploaded is not None:
    ext = uploaded.name.lower()
    if ext.endswith('.pdf'):
        text_by_page, tables_by_page = extract_text_pdf(uploaded)
    else:
        text_by_page, tables_by_page = extract_text_docx(uploaded)

    text_df = search_kpis_in_text(text_by_page, framework)
    table_df = pd.DataFrame()
    if extract_tables and search_tables:
        table_df = search_kpis_in_tables(tables_by_page, framework)

    combined = pd.concat([text_df, table_df], ignore_index=True, sort=False) if (not text_df.empty or not table_df.empty) else pd.DataFrame()

    if combined.empty:
        st.warning("No KPI hits with the current patterns. Try another framework, expand the dictionary, or lower the confidence threshold.")
    else:
        # filter by confidence
        filtered = combined[combined.Confidence >= min_confidence].copy()
        st.success(f"Found {len(combined)} KPI candidates — {len(filtered)} meet the confidence threshold (>= {min_confidence})")

        display_cols = ["Framework","KPI Name","Value","Unit","Source Page","Source","Confidence"]
        if show_evidence:
            display_cols += ["Evidence"]

        # Show high and low confidence splits for easier triage
        high = combined[combined.Confidence >= min_confidence]
        low = combined[combined.Confidence < min_confidence]

        with st.expander(f"High-confidence hits ({len(high)})", expanded=True):
            if not high.empty:
                st.dataframe(high[display_cols], use_container_width=True)
            else:
                st.write("— none —")

        with st.expander(f"Lower-confidence hits ({len(low)}) — review manually", expanded=False):
            if not low.empty:
                st.dataframe(low[display_cols], use_container_width=True)
            else:
                st.write("— none —")

        # Download Excel (two sheets: combined + tables)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            combined.to_excel(writer, index=False, sheet_name="All_KPI_Candidates")
            high.to_excel(writer, index=False, sheet_name="High_Confidence")
            low.to_excel(writer, index=False, sheet_name="Low_Confidence")
        st.download_button(
            "📥 Download KPI Results (Excel)",
            data=buffer.getvalue(),
            file_name="kpi_results_v2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optionally show extracted tables so user can inspect
        if extract_tables and any(len(v) for v in tables_by_page.values()):
            st.markdown("---")
            st.header("🔎 Extracted tables preview")
            for page, tables in tables_by_page.items():
                if not tables:
                    continue
                st.subheader(f"Page {page} — {len(tables)} table(s)")
                for i, tbl in enumerate(tables):
                    with st.expander(f"Table {i} (Page {page}) — preview"):
                        st.dataframe(tbl.fillna(''), use_container_width=True)

else:
    st.info("Upload a report and pick the framework from the sidebar to start.")

# ---------------------------
# End of app
# ---------------------------
