import streamlit as st
import pandas as pd
import json
from pathlib import Path

# -------------------------------------------------
# Config
# -------------------------------------------------
st.set_page_config(page_title="Inactive Tutor Executive Dashboard", layout="wide")

EXCEL_FILE = "Inactive_Tutor_Executive_Report_v7_FULL_FINAL.xlsx"
JSON_FILE = "parsed_tutor_data.json"  # optional but enables tutor-level filtering

BAND_TO_GRADES = {
    "K-2nd": [0, 1, 2],
    "3rd-5th": [3, 4, 5],
    "6th-8th": [6, 7, 8],
    "9th-12th": [9, 10, 11, 12],
}

HS_SPECIALTIES = {
    "Algebra 1", "Algebra 2", "Geometry", "Calculus", "Statistics", "Trigonometry", "PSAT/ACT/SAT Prep"
}

def expand_grade_band(band: str):
    if band in BAND_TO_GRADES:
        return BAND_TO_GRADES[band]
    if band in HS_SPECIALTIES:
        # treat HS specialties as HS grades (customize as needed)
        return [9, 10, 11, 12]
    return []

@st.cache_data
def load_workbook(path: str):
    xls = pd.ExcelFile(path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

@st.cache_data
def load_parsed_json(path: str):
    with open(path, "r") as f:
        return json.load(f)

def build_tutor_long_from_json(tutors: list):
    """Create a tutor-level long table from parsed_tutor_data.json for filtering."""
    rows = []
    for t in tutors:
        tutor_id = t.get("tutor_id")
        name = t.get("name")
        langs = t.get("languages", []) or []
        langs_norm = [str(x).strip() for x in langs if str(x).strip()]
        langs_join = ", ".join(sorted(set(langs_norm)))

        # cert sets
        cert_subjects = set()
        for c in (t.get("certifications", []) or []) + (t.get("second_certifications", []) or []):
            subj = c.get("subject")
            if subj is None:
                continue
            subj = str(subj).strip()
            # Map Reading -> ELA as requested
            if subj.lower() == "reading":
                subj = "ELA"
            cert_subjects.add(subj)

        # grade bands = coverage
        for gb in t.get("grade_bands", []) or []:
            cov_subject = gb.get("subject")
            band = gb.get("grade_band")
            grades = expand_grade_band(band)

            # Specialty: only for Math when band isn't a grade range
            specialty = None
            if cov_subject == "Math" and band not in BAND_TO_GRADES:
                specialty = band

            for g in grades:
                rows.append({
                    "tutor_id": tutor_id,
                    "name": name,
                    "coverage_subject": cov_subject,
                    "grade": g,
                    "math_specialty": specialty,
                    "languages": langs_norm,
                    "languages_str": langs_join,
                    "has_ela_cert": ("ELA" in cert_subjects),
                    "has_math_cert": ("Math" in cert_subjects),
                    "has_sped_cert": ("SPED" in cert_subjects),
                    "has_spanish_cert": ("Spanish" in cert_subjects),
                    "has_ir_cert": ("IR" in cert_subjects),
                    "cert_subjects": sorted(cert_subjects),
                })
    df = pd.DataFrame(rows)
    # Ensure list columns are safe
    if not df.empty:
        df["languages"] = df["languages"].apply(lambda x: x if isinstance(x, list) else [])
        df["cert_subjects"] = df["cert_subjects"].apply(lambda x: x if isinstance(x, list) else [])
    return df

# -------------------------------------------------
# Sidebar
# -------------------------------------------------
st.title("Inactive Tutor Executive Dashboard")

st.sidebar.header("Data files")
st.sidebar.write(f"• Excel expected: **{EXCEL_FILE}**")
st.sidebar.write(f"• Optional JSON (enables tutor lookup & filters): **{JSON_FILE}**")

excel_exists = Path(EXCEL_FILE).exists()
json_exists = Path(JSON_FILE).exists()
import os

st.sidebar.markdown("### JSON Debug")
st.sidebar.write("json_exists:", json_exists)

if json_exists:
    # show file size
    size_mb = os.path.getsize(JSON_FILE) / (1024 * 1024)
    st.sidebar.write("json_size_mb:", round(size_mb, 2))

    # show first line (this catches Git LFS pointers instantly)
    with open(JSON_FILE, "r", encoding="utf-8", errors="replace") as f:
        first_line = f.readline().strip()
    st.sidebar.write("json_first_line:", first_line[:120])

    # try parsing
    try:
        import json
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        st.sidebar.write("json_type:", type(data).__name__)
        st.sidebar.write("json_len:", len(data) if hasattr(data, "__len__") else "n/a")
    except Exception as e:
        st.sidebar.error("JSON failed to load:")
        st.sidebar.exception(e)


if not excel_exists:
    st.sidebar.warning("Excel workbook not found yet. Add it to this folder and redeploy/restart.")
if not json_exists:
    st.sidebar.info("If you also add parsed_tutor_data.json, you’ll get tutor-level filtering + tutor lists.")

# Load workbook if present
sheets = {}
if excel_exists:
    try:
        sheets = load_workbook(EXCEL_FILE)
    except Exception as e:
        st.error(f"Error loading Excel workbook: {e}")
        st.stop()

# Load JSON long table if present
tutor_long = pd.DataFrame()
if json_exists:
    try:
        tutors = load_parsed_json(JSON_FILE)
        tutor_long = build_tutor_long_from_json(tutors)
    except Exception as e:
        st.warning(f"Could not load {JSON_FILE}: {e}")

# -------------------------------------------------
# Executive charts (workbook-driven)
# -------------------------------------------------
if sheets and "Coverage Matrix" in sheets and "Certified Coverage Matrix" in sheets:
    st.header("Coverage vs Certified Coverage")

    cov = sheets["Coverage Matrix"]
    cert = sheets["Certified Coverage Matrix"]

    subject_col = cov.columns[0]
    grade_cols = [c for c in cov.columns[1:]]

    colA, colB = st.columns([1, 2])
    with colA:
        subject = st.selectbox("Subject", sorted(cov[subject_col].dropna().unique().tolist()))
        grade = st.selectbox("Grade", grade_cols)

    total = int(cov.loc[cov[subject_col] == subject, grade].fillna(0).iloc[0])
    certified = 0
    if subject in cert[subject_col].values:
        certified = int(cert.loc[cert[subject_col] == subject, grade].fillna(0).iloc[0])
    gap = max(total - certified, 0)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total coverage", total)
    m2.metric("Certified coverage", certified)
    m3.metric("Gap", gap)

    # Chart by grade for selected subject
    cov_row = cov.loc[cov[subject_col] == subject, grade_cols].iloc[0].fillna(0).astype(int)
    cert_row = (
        cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0).astype(int)
        if subject in cert[subject_col].values else pd.Series([0]*len(grade_cols), index=grade_cols)
    )
    chart_df = pd.DataFrame({"Total": cov_row.values, "Certified": cert_row.values}, index=[str(x) for x in grade_cols])
    st.bar_chart(chart_df)

# Math Specialty Coverage
if sheets and "Math Specialty Coverage" in sheets:
    st.header("Math Specialty Coverage")
    ms = sheets["Math Specialty Coverage"]
    ms = ms.sort_values(ms.columns[-1], ascending=False)
    st.bar_chart(ms.set_index(ms.columns[0]))

# Special Certification Flags
if sheets and "Special Certification Flags" in sheets:
    st.header("Special Certification Flags")
    flags = sheets["Special Certification Flags"]
    st.bar_chart(flags.set_index(flags.columns[0]))

# State Certification Counts
if sheets and "State Certification Counts" in sheets:
    st.header("State Certification Counts")
    states = sheets["State Certification Counts"].head(25)
    st.bar_chart(states.set_index(states.columns[0]))

# -------------------------------------------------
# Tutor lookup & filters (JSON-driven)
# -------------------------------------------------
st.divider()
st.header("Tutor Lookup (Filters)")

if tutor_long.empty:
    st.info(
        "To enable tutor-level filtering and tutor lists, add **parsed_tutor_data.json** to this repo alongside app.py, "
        "then redeploy. (We can also adapt this to a tutor-detail tab in Excel if you prefer.)"
    )
    st.stop()

# Filters
with st.sidebar:
    st.subheader("Filters")

    all_subjects = sorted([x for x in tutor_long["coverage_subject"].dropna().unique().tolist()])
    f_subjects = st.multiselect("Coverage subject", all_subjects, default=all_subjects)

    all_grades = sorted([int(x) for x in tutor_long["grade"].dropna().unique().tolist()])
    f_grades = st.multiselect("Grade", all_grades, default=all_grades)

    # Math specialty filter (only relevant when Math selected)
    all_specs = sorted([x for x in tutor_long["math_specialty"].dropna().unique().tolist()])
    f_specs = st.multiselect("Math specialty (optional)", all_specs, default=[])

    # Language filter
    # Build a flattened language list
    lang_set = set()
    for langs in tutor_long["languages"].tolist():
        for l in (langs or []):
            lang_set.add(str(l).strip())
    all_langs = sorted([l for l in lang_set if l])
    f_langs = st.multiselect("Language spoken (optional)", all_langs, default=[])

    # Certification flags
    require_ela = st.checkbox("Require ELA cert")
    require_math = st.checkbox("Require Math cert")
    require_sped = st.checkbox("Require SPED cert")
    require_esl = st.checkbox("Require ESL cert (not in JSON)")  # placeholder
    require_ir = st.checkbox("Require IR cert")
    require_spanish = st.checkbox("Require Spanish cert")

# Apply filters
flt = tutor_long.copy()

flt = flt[flt["coverage_subject"].isin(f_subjects)]
flt = flt[flt["grade"].isin(f_grades)]

if f_specs:
    flt = flt[flt["math_specialty"].isin(f_specs)]

if f_langs:
    # Keep row if any selected language is in the tutor's language list
    flt = flt[flt["languages"].apply(lambda ls: any(l in (ls or []) for l in f_langs))]

if require_ela:
    flt = flt[flt["has_ela_cert"] == True]
if require_math:
    flt = flt[flt["has_math_cert"] == True]
if require_sped:
    flt = flt[flt["has_sped_cert"] == True]
if require_ir:
    flt = flt[flt["has_ir_cert"] == True]
if require_spanish:
    flt = flt[flt["has_spanish_cert"] == True]

# Aggregate to tutor list (unique tutors)
tutors_df = (
    flt.groupby(["tutor_id", "name"], as_index=False)
       .agg(
           subjects=("coverage_subject", lambda s: ", ".join(sorted(set(s)))),
           grades=("grade", lambda g: ", ".join(map(str, sorted(set(map(int, g)))))),
           specialties=("math_specialty", lambda s: ", ".join(sorted(set([x for x in s if pd.notna(x)])))),
           languages=("languages_str", "first"),
           certs=("cert_subjects", lambda c: ", ".join(sorted(set(sum(c, []))))),
       )
       .sort_values("name")
)

st.subheader("Results")
st.write(f"Matching tutor-grade rows: **{len(flt):,}**")
st.write(f"Unique tutors: **{len(tutors_df):,}**")

st.dataframe(tutors_df, use_container_width=True, hide_index=True)

st.caption("Tip: Add parsed_tutor_data.json to enable these filters. Excel-only deployments can still show executive charts.")
