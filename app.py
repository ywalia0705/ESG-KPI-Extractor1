import streamlit as st
import pandas as pd
import pdfplumber
import docx2txt
import tempfile
import re
import io
from pathlib import Path
st.set_page_config(page_title="ESG KPI Extraction Tool", layout="wide")
st.title("📊 ESG KPI Extraction Tool (MVP)")
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
        "Water withdrawal": [r"\bwater\s+withdrawal\b", r"\bwater\s+use(age)?\b"],
        "Waste generated": [r"\bwaste\s+generated\b", r"\btotal\s+waste\b"],
        "GHG emissions (total)": [r"\bGHG\s+emissions\b", r"\bgreenhouse\s+gas\b.*\bemissions\b"],
    },
    "TNFD": {
        "Biodiversity impact": [r"\bbiodiversity\s+impact\b", r"\becosystem\b.*\bimpact\b"],
        "Land use / habitat": [r"\bland\s+use\b", r"\bhabitat\s+(?:loss|conversion)\b"],
        "Natural capital dependency": [r"\bnatural\s+capital\s+dependenc(y|ies)\b", r"\bnature[-\s]?related\s+dependencies\b"],
        "Water stress exposure": [r"\bwater\s+stress\b", r"\bhigh\s+baseline\s+water\s+stress\b"],
    }
}

# Units we try to capture after numbers (non-exhaustive; add as needed)
UNIT_PATTERN = r"(?:tCO2e|CO2e|kgCO2e|kt|t|tonnes?|mt|m3|m³|kWh|MWh|GWh|GJ|TJ|%|ppm|ha|km2|km²|L|litres?|USD|EUR|GBP|₹|€|£)?"

NUMBER_PATTERN = rf"(-?\d{{1,3}}(?:,\d{{3}})*|\d+)(?:\.\d+)?\s*{UNIT_PATTERN}"

def _save_to_tmp(uploaded_file, suffix: str):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.flush()
    return tmp.name

def extract_text_pdf(uploaded_file) -> dict:
    path = _save_to_tmp(uploaded_file, ".pdf")
    text_by_page = {}
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text_by_page[i + 1] = page.extract_text() or ""
    Path(path).unlink(missing_ok=True)
    return text_by_page

def extract_text_docx(uploaded_file) -> dict:
    path = _save_to_tmp(uploaded_file, ".docx")
    text = docx2txt.process(path) or ""
    Path(path).unlink(missing_ok=True)
    return {1: text}

def find_values_near(text: str, start: int, end: int, window_after=220, window_before=50):
    """Search for the closest number+unit near a KPI mention."""
    left_start = max(0, start - window_before)
    right_end = min(len(text), end + window_after)
    window = text[left_start:right_end]
    matches = list(re.finditer(NUMBER_PATTERN, window, flags=re.IGNORECASE))
    if not matches:
        return None, None
    # Choose the first number after the KPI phrase if possible; otherwise the first in window
    after = [m for m in matches if (left_start + m.start()) >= end]
    chosen = after[0] if after else matches[0]
    value = chosen.group(0).strip()
    # Split out raw number and unit if present
    unit_match = re.search(r"[a-zA-Z%₹€£]+$", value)
    unit = unit_match.group(0) if unit_match else ""
    num = value.replace(unit, "").strip()
    return num, unit

def search_kpis(text_by_page: dict, framework: str):
    out = []
    patterns = KPI_DICTIONARIES.get(framework, {})
    for page, text in text_by_page.items():
        if not text:
            continue
        for kpi_name, pat_list in patterns.items():
            for pat in pat_list:
                for m in re.finditer(pat, text, flags=re.IGNORECASE | re.DOTALL):
                    val, unit = find_values_near(text, m.start(), m.end())
                    out.append({
                        "Framework": framework,
                        "KPI Name": kpi_name,
                        "Value": val or "",
                        "Unit": unit or "",
                        "Source Page": page,
                        "Evidence": text[max(0, m.start()-60): m.end()+120].replace("\n", " ")
                    })
    # De-duplicate by KPI+Page (keep first)
    if out:
        df = pd.DataFrame(out)
        df = df.sort_values(["KPI Name", "Source Page"])
        df = df.drop_duplicates(subset=["KPI Name", "Source Page"], keep="first")
        return df
    return pd.DataFrame(columns=["Framework","KPI Name","Value","Unit","Source Page","Evidence"])

# ---------------------------
# UI
# ---------------------------
with st.sidebar:
    st.header("⚙️ Settings")
    framework = st.selectbox("Report framework", ["TCFD", "GRI", "TNFD"])
    show_evidence = st.checkbox("Show evidence text", value=False)
    st.markdown("---")
    st.caption("Tip: For scanned PDFs (images), text extraction may be blank; OCR can be added later.")

uploaded = st.file_uploader("Upload report (PDF or DOCX)", type=["pdf", "docx"])

if uploaded is not None:
    if uploaded.name.lower().endswith(".pdf"):
        pages = extract_text_pdf(uploaded)
    else:
        pages = extract_text_docx(uploaded)

    df = search_kpis(pages, framework)

    if df.empty:
        st.warning("No KPI hits with the current patterns. Try another framework or expand the dictionary.")
    else:
        display_cols = ["Framework","KPI Name","Value","Unit","Source Page"]
        if show_evidence:
            display_cols += ["Evidence"]
        st.success(f"Found {len(df)} KPI entries")
        st.dataframe(df[display_cols], use_container_width=True)

        # Download Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="KPI Results")
        st.download_button(
            "📥 Download KPI Results (Excel)",
            data=buffer.getvalue(),
            file_name="kpi_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload a report and pick the framework from the sidebar to start.")
