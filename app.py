
import streamlit as st
import pandas as pd
import json
import io
import os
from pathlib import Path

# -------------------------------------------------
# Optional plotting (heatmap). If matplotlib isn't available,
# we gracefully skip the heatmap instead of crashing.
# -------------------------------------------------
try:
    import matplotlib.pyplot as plt
    HAS_MPL = True
except Exception:
    HAS_MPL = False

# -------------------------------------------------
# Config
# -------------------------------------------------
st.set_page_config(page_title="Inactive Tutor Executive Dashboard", layout="wide")

EXCEL_FILE = "Inactive_Tutor_Executive_Report_v7_FULL_FINAL.xlsx"
JSON_FILE = "parsed_tutor_data.json"

BAND_TO_GRADES = {
    "K-2nd": [0, 1, 2],
    "3rd-5th": [3, 4, 5],
    "6th-8th": [6, 7, 8],
    "9th-12th": [9, 10, 11, 12],
}

HS_SPECIALTIES = {
    "Algebra 1", "Algebra 2", "Geometry", "Calculus",
    "Statistics", "Trigonometry", "PSAT/ACT/SAT Prep"
}


def expand_grade_band(band: str):
    if band is None:
        return []
    band = str(band).strip()
    if ":" in band:
        band = band.split(":", 1)[1].strip()

    if band in BAND_TO_GRADES:
        return BAND_TO_GRADES[band]
    if band in HS_SPECIALTIES:
        return [9, 10, 11, 12]
    return []


@st.cache_data
def load_workbook(path: str):
    xls = pd.ExcelFile(path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}


@st.cache_data
def load_parsed_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_tutor_long_from_json(tutors: list) -> pd.DataFrame:
    rows = []
    for t in tutors:
        tutor_id = t.get("tutor_id")
        name = t.get("name")

        langs = t.get("languages", []) or []
        langs_norm = [str(x).strip() for x in langs if str(x).strip()]
        langs_join = ", ".join(sorted(set(langs_norm)))

        cert_subjects = set()
        for c in (t.get("certifications", []) or []) + (t.get("second_certifications", []) or []):
            subj = c.get("subject")
            if subj is None:
                continue
            subj = str(subj).strip()
            if subj.lower() == "reading":
                subj = "ELA"
            cert_subjects.add(subj)

        for gb in t.get("grade_bands", []) or []:
            cov_subject = gb.get("subject")
            band_raw = gb.get("grade_band")
            band = str(band_raw).strip() if band_raw is not None else ""
            band_norm = band.split(":", 1)[1].strip() if ":" in band else band

            grades = expand_grade_band(band)

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
                    "languages_str": langs_join,
                    "cert_subjects": sorted(cert_subjects),
                    "has_ela_cert": ("ELA" in cert_subjects),
                    "has_math_cert": ("Math" in cert_subjects),
                    "has_sped_cert": ("SPED" in cert_subjects),
                    "has_spanish_cert": ("Spanish" in cert_subjects),
                    "has_ir_cert": ("IR" in cert_subjects),
                })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["languages"] = df["languages"].apply(lambda x: x if isinstance(x, list) else [])
        df["cert_subjects"] = df["cert_subjects"].apply(lambda x: x if isinstance(x, list) else [])
    return df


def safe_int(x):
    try:
        return int(x)
    except Exception:
        return 0


def make_gap_heatmap(cov_matrix, cert_matrix, subject):
    if not HAS_MPL:
        return None

    subject_col = cov_matrix.columns[0]
    grade_cols = cov_matrix.columns[1:]

    cov_row = cov_matrix.loc[cov_matrix[subject_col] == subject, grade_cols].iloc[0].fillna(0).apply(safe_int)
    if subject in cert_matrix[subject_col].values:
        cert_row = cert_matrix.loc[cert_matrix[subject_col] == subject, grade_cols].iloc[0].fillna(0).apply(safe_int)
    else:
        cert_row = pd.Series([0]*len(grade_cols), index=grade_cols)

    gap = (cov_row - cert_row).clip(lower=0)

    fig, ax = plt.subplots()
    ax.imshow(gap.values.reshape(1, -1), aspect="auto")

    ax.set_yticks([0])
    ax.set_yticklabels(["Gap"])
    ax.set_xticks(range(len(grade_cols)))
    ax.set_xticklabels(grade_cols)

    for j, val in enumerate(gap.values):
        ax.text(j, 0, str(int(val)), ha="center", va="center")

    ax.set_title(f"Certified Coverage Gap — {subject}")
    return fig


# -------------------------------------------------
# Load Data
# -------------------------------------------------
st.title("Inactive Tutor Executive Dashboard")

excel_exists = Path(EXCEL_FILE).exists()
json_exists = Path(JSON_FILE).exists()

sheets = load_workbook(EXCEL_FILE) if excel_exists else {}
tutor_long = build_tutor_long_from_json(load_parsed_json(JSON_FILE)) if json_exists else pd.DataFrame()


# -------------------------------------------------
# Math Specialty Coverage (clean version)
# -------------------------------------------------
if sheets and "Math Specialty Coverage" in sheets:
    st.header("Math Specialty Coverage")

    ms = sheets["Math Specialty Coverage"].copy()
    ms = ms.loc[:, ~ms.columns.astype(str).str.contains("^Unnamed")]

    specialty_col = ms.columns[0]
    numeric_cols = ms.select_dtypes(include="number").columns

    if len(numeric_cols) > 0:
        count_col = numeric_cols[0]
        ms = ms[[specialty_col, count_col]].dropna()
        ms[count_col] = ms[count_col].astype(int)
        ms = ms.sort_values(count_col, ascending=False)

        st.bar_chart(ms.set_index(specialty_col)[count_col])
    else:
        st.warning("No numeric count column found.")


# -------------------------------------------------
# Tutor Lookup
# -------------------------------------------------
st.header("Tutor Lookup")

if tutor_long.empty:
    st.warning("Tutor lookup disabled.")
else:
    search = st.sidebar.text_input("Search tutor name")

    flt = tutor_long.copy()
    if search:
        flt = flt[flt["name"].str.contains(search, case=False, na=False)]

    tutors_df = flt.groupby(["tutor_id", "name"], as_index=False).agg(
        subjects=("coverage_subject", lambda s: ", ".join(sorted(set(s)))),
        grades=("grade", lambda g: ", ".join(map(str, sorted(set(g))))),
        languages=("languages_str", "first"),
        certs=("cert_subjects", lambda c: ", ".join(sorted(set(sum(c, []))))),
    )

    st.dataframe(tutors_df)

    buffer = io.BytesIO()
    tutors_df.to_excel(buffer, index=False, engine="openpyxl")

    st.download_button(
        label="Download filtered tutors",
        data=buffer.getvalue(),
        file_name="filtered_inactive_tutors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
