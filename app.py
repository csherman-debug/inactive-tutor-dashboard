import streamlit as st
import pandas as pd
import json
import io
from pathlib import Path

# -----------------------------
# Files expected in repo root
# -----------------------------
EXCEL_FILE = "Inactive_Tutor_Executive_Report_v7_FULL_FINAL.xlsx"
JSON_FILE  = "parsed_tutor_data.json"

st.set_page_config(page_title="Inactive Tutor Dashboard", layout="wide")
st.title("Inactive Tutor Dashboard (Inactive Pool)")

# -----------------------------
# Helpers
# -----------------------------
@st.cache_data
def load_sheets(xlsx_path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(xlsx_path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

@st.cache_data
def load_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def clean_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    # Drop typical Excel artifact columns and dedupe headers
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed", case=False, regex=True)]
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def pick_count_column(df: pd.DataFrame, preferred_names=("Inactive Tutor Count","Count","Tutor Count")) -> str | None:
    # Prefer named columns
    cols = [str(c) for c in df.columns]
    for name in preferred_names:
        if name in cols:
            return name
    # Otherwise pick the last numeric-ish column
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        return numeric_cols[-1]
    # Try coercing last column
    if len(df.columns) >= 2:
        return df.columns[-1]
    return None

def normalize_cert_subject(subj: str) -> str:
    if subj is None:
        return ""
    s = str(subj).strip()
    if s.lower() == "reading":
        return "ELA"
    return s

# Grade-band expansion (for tutor-level coverage rows)
BAND_TO_GRADES = {
    "K-2nd": [0,1,2],
    "3rd-5th": [3,4,5],
    "6th-8th": [6,7,8],
    "9th-12th": [9,10,11,12],
}
HS_SPECIALTIES = {"Algebra 1","Algebra 2","Geometry","Calculus","Statistics","Trigonometry","PSAT/ACT/SAT Prep"}

def expand_grade_band(band: str):
    if band is None:
        return []
    b = str(band).strip()
    # handle "Math: Algebra 1"
    if ":" in b:
        b = b.split(":", 1)[1].strip()
    if b in BAND_TO_GRADES:
        return BAND_TO_GRADES[b]
    if b in HS_SPECIALTIES:
        return [9,10,11,12]
    return []

def build_tutor_rows(tutors: list[dict]) -> pd.DataFrame:
    rows = []
    for t in tutors:
        tutor_id = t.get("tutor_id")
        name = t.get("name")

        langs = t.get("languages", []) or []
        langs_norm = sorted({str(x).strip() for x in langs if str(x).strip()})
        langs_str = ", ".join(langs_norm)

        cert_subjects = set()
        for c in (t.get("certifications", []) or []) + (t.get("second_certifications", []) or []):
            cert_subjects.add(normalize_cert_subject(c.get("subject")))

        for gb in (t.get("grade_bands", []) or []):
            cov_subject = gb.get("subject")
            band_raw = gb.get("grade_band")
            band_str = str(band_raw).strip() if band_raw is not None else ""
            band_norm = band_str.split(":", 1)[1].strip() if ":" in band_str else band_str

            grades = expand_grade_band(band_str)
            specialty = None
            if cov_subject == "Math" and band_norm and band_norm not in BAND_TO_GRADES:
                specialty = band_norm

            for g in grades:
                rows.append({
                    "tutor_id": tutor_id,
                    "name": name,
                    "coverage_subject": cov_subject,
                    "grade": int(g),
                    "math_specialty": specialty,
                    "languages": langs_norm,
                    "languages_str": langs_str,
                    "cert_subjects": sorted({c for c in cert_subjects if c}),
                    "has_ela_cert": ("ELA" in cert_subjects),
                    "has_math_cert": ("Math" in cert_subjects),
                    "has_sped_cert": ("SPED" in cert_subjects),
                    "has_ir_cert": ("IR" in cert_subjects),
                    "has_spanish_cert": ("Spanish" in cert_subjects),
                })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["languages"] = df["languages"].apply(lambda x: x if isinstance(x, list) else [])
    df["cert_subjects"] = df["cert_subjects"].apply(lambda x: x if isinstance(x, list) else [])
    return df

# -----------------------------
# Sidebar: status + filters
# -----------------------------
st.sidebar.header("Data status")

excel_exists = Path(EXCEL_FILE).exists()
json_exists  = Path(JSON_FILE).exists()

st.sidebar.write("Excel:", "✅" if excel_exists else "❌", EXCEL_FILE)
st.sidebar.write("JSON:",  "✅" if json_exists else "❌", JSON_FILE)

# Load sources
sheets = {}
if excel_exists:
    try:
        sheets = load_sheets(EXCEL_FILE)
    except Exception as e:
        st.error(f"Failed to load Excel workbook: {e}")
        sheets = {}

tutor_long = pd.DataFrame()
if json_exists:
    try:
        tutor_long = build_tutor_rows(load_json(JSON_FILE))
    except Exception as e:
        st.error(f"Failed to load JSON: {e}")
        tutor_long = pd.DataFrame()

# -----------------------------
# Executive summary (Excel-driven)
# -----------------------------
st.header("Executive Summary")

c1, c2, c3, c4 = st.columns(4)
if not tutor_long.empty:
    c1.metric("Inactive tutors (unique)", f"{tutor_long['tutor_id'].nunique():,}")
    c2.metric("Math coverage (unique)", f"{tutor_long.loc[tutor_long['coverage_subject']=='Math','tutor_id'].nunique():,}")
    c3.metric("ELA/Lit coverage (unique)", f"{tutor_long.loc[tutor_long['coverage_subject']=='ELA/Literacy','tutor_id'].nunique():,}")
    c4.metric("SPED certified (unique)", f"{tutor_long.loc[tutor_long['has_sped_cert']==True,'tutor_id'].nunique():,}")
else:
    c1.metric("Inactive tutors (unique)", "—")
    c2.metric("Math coverage (unique)", "—")
    c3.metric("ELA/Lit coverage (unique)", "—")
    c4.metric("SPED certified (unique)", "—")

if sheets and "Special Certification Flags" in sheets:
    st.subheader("Special Certification Flags")
    flags = clean_excel_df(sheets["Special Certification Flags"])
    st.dataframe(flags, use_container_width=True, hide_index=True)

# -----------------------------
# Coverage vs Certified (Excel-driven)
# -----------------------------
st.header("Coverage vs Certified Coverage")

if sheets and "Coverage Matrix" in sheets and "Certified Coverage Matrix" in sheets:
    cov = clean_excel_df(sheets["Coverage Matrix"])
    cert = clean_excel_df(sheets["Certified Coverage Matrix"])

    subject_col = cov.columns[0]
    grade_cols = [c for c in cov.columns[1:]]

    left, right = st.columns([1, 2])
    with left:
        subject = st.selectbox("Subject", sorted(cov[subject_col].dropna().unique().tolist()))
    # build a chart df by grade
    cov_row = cov.loc[cov[subject_col] == subject, grade_cols].iloc[0].fillna(0)
    cov_row = pd.to_numeric(cov_row, errors="coerce").fillna(0).astype(int)

    if subject in cert[subject_col].values:
        cert_row = cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0)
        cert_row = pd.to_numeric(cert_row, errors="coerce").fillna(0).astype(int)
    else:
        cert_row = pd.Series([0]*len(grade_cols), index=grade_cols)

    gap_row = (cov_row - cert_row).clip(lower=0)

    with right:
        chart_df = pd.DataFrame(
            {"Total": cov_row.values, "Certified": cert_row.values, "Gap": gap_row.values},
            index=[str(x) for x in grade_cols],
        )
        st.bar_chart(chart_df)

else:
    st.info("Coverage matrices not found in the workbook (Coverage Matrix / Certified Coverage Matrix).")

# -----------------------------
# Math Specialty Coverage (Excel-driven, robust)
# -----------------------------
st.header("Math Specialty Coverage")

if sheets and "Math Specialty Coverage" in sheets:
    ms = clean_excel_df(sheets["Math Specialty Coverage"])

    if ms.shape[1] < 2:
        st.warning("Math Specialty Coverage sheet doesn't have at least 2 columns.")
    else:
        specialty_col = ms.columns[0]
        count_col = pick_count_column(ms)

        if count_col is None:
            st.warning("Could not determine the count column for Math Specialty Coverage.")
            st.dataframe(ms.head(30), use_container_width=True, hide_index=True)
        else:
            # Coerce count column to numeric safely
            ms[count_col] = pd.to_numeric(ms[count_col], errors="coerce").fillna(0).astype(int)

            plot_df = (
                ms[[specialty_col, count_col]]
                .dropna(subset=[specialty_col])
                .groupby(specialty_col, as_index=False)[count_col].sum()
                .sort_values(count_col, ascending=False)
            )

            st.bar_chart(plot_df.set_index(specialty_col)[count_col])
            st.caption("Y-axis = unique inactive tutor count per math specialty.")
else:
    st.info("Math Specialty Coverage sheet not found in the workbook.")

# -----------------------------
# Tutor Lookup (JSON-driven) + export
# -----------------------------
st.header("Tutor Lookup")

if tutor_long.empty:
    st.info("Tutor lookup is disabled until parsed_tutor_data.json is present and readable.")
else:
    with st.sidebar:
        st.subheader("Lookup filters")

        search = st.text_input("Search tutor name", value="").strip()

        subjects = sorted(tutor_long["coverage_subject"].dropna().unique().tolist())
        f_subjects = st.multiselect("Coverage subject", subjects, default=subjects)

        grades = sorted([int(x) for x in tutor_long["grade"].dropna().unique().tolist()])
        f_grades = st.multiselect("Grade", grades, default=grades)

        specs = sorted(tutor_long["math_specialty"].dropna().unique().tolist())
        f_specs = st.multiselect("Math specialty (optional)", specs, default=[])

        # Language filter
        lang_set = set()
        for ls in tutor_long["languages"].tolist():
            for l in (ls or []):
                lang_set.add(str(l).strip())
        all_langs = sorted([l for l in lang_set if l])
        f_langs = st.multiselect("Language spoken (optional)", all_langs, default=[])

        st.caption("Certification requirements")
        require_ela = st.checkbox("Require ELA cert")
        require_math = st.checkbox("Require Math cert")
        require_sped = st.checkbox("Require SPED cert")
        require_ir = st.checkbox("Require IR cert")
        require_spanish = st.checkbox("Require Spanish cert")

    flt = tutor_long.copy()
    flt = flt[flt["coverage_subject"].isin(f_subjects)]
    flt = flt[flt["grade"].isin(f_grades)]

    if f_specs:
        flt = flt[flt["math_specialty"].isin(f_specs)]
    if f_langs:
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

    if search:
        flt = flt[flt["name"].fillna("").str.contains(search, case=False)]

    # District staffing summary (based on current filter)
    st.subheader("District Staffing Mode")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Matching tutors (unique)", f"{flt['tutor_id'].nunique():,}")
    s2.metric("Math (unique)", f"{flt.loc[flt['coverage_subject']=='Math','tutor_id'].nunique():,}")
    s3.metric("ELA/Lit (unique)", f"{flt.loc[flt['coverage_subject']=='ELA/Literacy','tutor_id'].nunique():,}")
    s4.metric("Spanish-certified (unique)", f"{flt.loc[flt['has_spanish_cert']==True,'tutor_id'].nunique():,}")

    # Aggregate to tutor list
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

    # Export
    st.subheader("Export")
    out = io.BytesIO()
    tutors_df.to_excel(out, index=False, engine="openpyxl")
    st.download_button(
        "Download filtered tutors (Excel)",
        data=out.getvalue(),
        file_name="filtered_inactive_tutors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
