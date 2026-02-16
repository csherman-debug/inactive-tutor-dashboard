import streamlit as st
import pandas as pd
import json
import io
from pathlib import Path

EXCEL_FILE = "Inactive_Tutor_Executive_Report_v7_FULL_FINAL.xlsx"
JSON_FILE  = "parsed_tutor_data.json"

st.set_page_config(page_title="Inactive Tutor Dashboard", layout="wide")
st.title("Inactive Tutor Dashboard (Inactive Pool)")

@st.cache_data
def load_sheets(xlsx_path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(xlsx_path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

@st.cache_data
def load_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def clean_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed", case=False, regex=True)]
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def pick_label_and_count_columns(df: pd.DataFrame):
    df = df.copy()
    label_col = None
    for c in df.columns:
        if not pd.api.types.is_numeric_dtype(df[c]):
            label_col = c
            break
    if label_col is None:
        label_col = df.columns[0]

    preferred = ["Inactive Tutor Count", "Tutor Count", "Count"]
    cols_str = [str(c) for c in df.columns]
    count_col = None
    for name in preferred:
        if name in cols_str:
            count_col = name
            break
    if count_col is None:
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        if numeric_cols:
            count_col = numeric_cols[-1]
        else:
            count_col = df.columns[-1]
            df[count_col] = pd.to_numeric(df[count_col], errors="coerce")
    return label_col, count_col, df

def normalize_cert_subject(subj: str) -> str:
    if subj is None:
        return ""
    s = str(subj).strip()
    if s.lower() == "reading":
        return "ELA"
    return s

BAND_TO_GRADES = {"K-2nd":[0,1,2], "3rd-5th":[3,4,5], "6th-8th":[6,7,8], "9th-12th":[9,10,11,12]}
HS_SPECIALTIES = {"Algebra 1","Algebra 2","Geometry","Calculus","Statistics","Trigonometry","PSAT/ACT/SAT Prep"}

def expand_grade_band(band: str):
    if band is None:
        return []
    b = str(band).strip()
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

# Sidebar status
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

# Executive summary
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

# Coverage by grade (no gap)
st.header("Coverage by Grade Level")

if sheets and "Coverage Matrix" in sheets:
    cov = clean_excel_df(sheets["Coverage Matrix"])
    cert = clean_excel_df(sheets["Certified Coverage Matrix"]) if (sheets and "Certified Coverage Matrix" in sheets) else None

    subject_col = cov.columns[0]
    grade_cols = list(cov.columns[1:])

    subject = st.selectbox("Select subject area", sorted(cov[subject_col].dropna().unique().tolist()))

    total_row = cov.loc[cov[subject_col] == subject, grade_cols].iloc[0].fillna(0)
    total_row = pd.to_numeric(total_row, errors="coerce").fillna(0).astype(int)

    chart_df = pd.DataFrame({"Total Coverage": total_row.values}, index=[str(g) for g in grade_cols])

    show_cert = st.checkbox("Show certified coverage overlay", value=True)
    if show_cert and cert is not None and subject in cert[subject_col].values:
        cert_row = cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0)
        cert_row = pd.to_numeric(cert_row, errors="coerce").fillna(0).astype(int)
        chart_df["Certified Coverage"] = cert_row.values

    st.bar_chart(chart_df)
    st.caption("X-axis: grade level columns from the executive workbook. Y-axis: unique inactive tutor counts.")
else:
    st.info("Coverage Matrix sheet not found in the workbook.")

# Math specialty coverage (fix labels)
st.header("Math Specialty Coverage")

if sheets and "Math Specialty Coverage" in sheets:
    ms_raw = clean_excel_df(sheets["Math Specialty Coverage"])

    if ms_raw.shape[1] < 2:
        st.warning("Math Specialty Coverage sheet doesn't have at least 2 columns.")
    else:
        label_col, count_col, ms = pick_label_and_count_columns(ms_raw)
        ms[count_col] = pd.to_numeric(ms[count_col], errors="coerce").fillna(0).astype(int)
        ms[label_col] = ms[label_col].astype(str).str.strip()
        ms = ms[ms[label_col].ne("")]

        plot_df = (
            ms[[label_col, count_col]]
            .groupby(label_col, as_index=False)[count_col].sum()
            .sort_values(count_col, ascending=False)
        )

        st.bar_chart(plot_df.set_index(label_col)[count_col])
        st.caption("X-axis: Math specialty label. Y-axis: unique inactive tutor counts.")
else:
    st.info("Math Specialty Coverage sheet not found in the workbook.")

# Tutor lookup (no district staffing mode)
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

    st.subheader("Export")
    out = io.BytesIO()
    tutors_df.to_excel(out, index=False, engine="openpyxl")
    st.download_button(
        "Download filtered tutors (Excel)",
        data=out.getvalue(),
        file_name="filtered_inactive_tutors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
