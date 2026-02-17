import streamlit as st
import pandas as pd
import altair as alt
import json
import io
import re
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

def make_unique_columns(cols):
    """Return list of unique column names by suffixing duplicates."""
    seen = {}
    out = []
    for c in cols:
        name = str(c)
        if name not in seen:
            seen[name] = 0
            out.append(name)
        else:
            seen[name] += 1
            out.append(f"{name}.{seen[name]}")
    return out

def clean_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    # Drop typical Excel artifact columns (Unnamed: 0, etc.)
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed", case=False, regex=True)]
    # Ensure truly unique column labels (important for sort_values)
    df.columns = make_unique_columns(df.columns)
    return df

def pick_label_and_count_columns(df: pd.DataFrame):
    """Pick a label (category) column and a count column robustly."""
    df = df.copy()
    df.columns = make_unique_columns(df.columns)

    # Prefer first non-numeric column as label, but avoid 'index' if present
    label_col = None
    for c in df.columns:
        if str(c).lower().startswith("index"):
            continue
        if not pd.api.types.is_numeric_dtype(df[c]):
            label_col = c
            break
    if label_col is None:
        label_col = df.columns[0]

    # Prefer a known count name, otherwise last numeric column (avoid index columns)
    preferred = ["Inactive Tutor Count", "Tutor Count", "Count"]
    count_col = None
    cols_lower = {str(c).lower(): c for c in df.columns}
    for name in preferred:
        if name.lower() in cols_lower:
            count_col = cols_lower[name.lower()]
            break

    if count_col is None:
        numeric_cols = [c for c in df.select_dtypes(include="number").columns.tolist() if not str(c).lower().startswith("index")]
        if numeric_cols:
            count_col = numeric_cols[-1]
        else:
            # Try coercing last non-index column
            non_index = [c for c in df.columns if not str(c).lower().startswith("index")]
            count_col = non_index[-1] if non_index else df.columns[-1]
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

def mask_has_any_language(df: pd.DataFrame, langs: list[str]) -> pd.Series:
    """Boolean mask: row has ANY of the selected languages. Safe on empty frames."""
    if df.empty:
        return pd.Series([], dtype=bool, index=df.index)
    selected = [str(l).strip() for l in (langs or []) if str(l).strip()]
    if not selected:
        return pd.Series([True]*len(df), dtype=bool, index=df.index)
    s = df["languages"].apply(lambda ls: any(l in (ls or []) for l in selected))
    # Ensure boolean dtype and aligned index
    return s.fillna(False).astype(bool)

def grade_label_to_int(label) -> int:
    s = str(label).strip().upper()
    if s in {"K", "KG", "KINDER", "KINDERGARTEN"}:
        return 0
    m = re.search(r"\d+", s)
    if m:
        return int(m.group(0))
    return 999

def sort_grade_cols(cols):
    return sorted(cols, key=grade_label_to_int)

def grade_int_to_label(g: int) -> str:
    return "K" if int(g) == 0 else str(int(g))

def grade_token_to_grades(token: str):
    t = str(token).strip().upper()
    if t == "ES":
        return list(range(0, 6))  # K-5
    if t == "MS":
        return [6,7,8]
    if t == "HS":
        return [9,10,11,12]
    if t in {"K", "KG"}:
        return [0]
    m = re.search(r"\d+", t)
    if m:
        return [int(m.group(0))]
    return []

def display_grade_label(label) -> str:
    g = grade_label_to_int(label)
    if g == 0:
        return "K"
    if g == 999:
        return str(label)
    return str(g)

def make_grade_band_series(series_by_grade_label: pd.Series) -> pd.Series:
    items = []
    for lbl, val in series_by_grade_label.items():
        g = grade_label_to_int(lbl)
        items.append((g, int(val)))
    df = pd.DataFrame(items, columns=["grade", "value"])
    def band(g):
        if g <= 5: return "K-5"
        if 6 <= g <= 8: return "6-8"
        if 9 <= g <= 12: return "9-12"
        return "Other"
    df["band"] = df["grade"].apply(band)
    out = df.groupby("band")["value"].sum()
    order = ["K-5","6-8","9-12","Other"]
    out = out.reindex([b for b in order if b in out.index])
    return out

# -----------------------------
# Load sources
# -----------------------------
excel_exists = Path(EXCEL_FILE).exists()
json_exists  = Path(JSON_FILE).exists()

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
# Sidebar filters (lookup only)
# -----------------------------
with st.sidebar:
    st.subheader("Lookup filters")
    if tutor_long.empty:
        st.caption("Tutor-level filters available when parsed_tutor_data.json is present.")
        search = ""
        f_subjects = []
        f_grades = []
        f_specs = []
        f_langs = []
        require_ela = require_math = require_sped = require_ir = require_spanish = False
    else:
        search = st.text_input("Search tutor name", value="").strip()

        subjects = sorted(tutor_long["coverage_subject"].dropna().unique().tolist())
        f_subjects = st.multiselect("Coverage subject", subjects, default=subjects)

        grades = sorted([int(x) for x in tutor_long["grade"].dropna().unique().tolist()])
        # Grade filter supports ES/MS/HS + individual grades (K shown instead of 0)
        grade_tokens = ["ES", "MS", "HS"] + [grade_int_to_label(g) for g in grades]
        default_tokens = [grade_int_to_label(g) for g in grades]
        selected_tokens = st.multiselect("Grade", grade_tokens, default=default_tokens)
        # Expand tokens into grade integers
        expanded = set()
        for tok in selected_tokens:
            for g in grade_token_to_grades(tok):
                expanded.add(int(g))
        f_grades = sorted(expanded)


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

# -----------------------------
# Executive summary
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

    flags_raw = clean_excel_df(sheets["Special Certification Flags"])
    # Drop any real "index" columns that may exist in the sheet
    flags_raw = flags_raw.loc[:, ~flags_raw.columns.astype(str).str.match(r"(?i)^index(\.|$)")]

    # Try to robustly pick a label column + count column, then chart.
    try:
        label_col, count_col, flags = pick_label_and_count_columns(flags_raw)
        flags[count_col] = pd.to_numeric(flags[count_col], errors="coerce").fillna(0).astype(int)
        flags[label_col] = flags[label_col].astype(str).str.strip()
        flags = flags[flags[label_col].ne("")]

        agg = (
            flags[[label_col, count_col]]
            .groupby(label_col, as_index=False)[count_col].sum()
            .sort_values(count_col, ascending=False)
        )

        # Keep the pie readable: top 10 categories + "Other"
        top_n = min(10, len(agg))
        top = agg.head(top_n).copy()
        if len(agg) > top_n:
            other_val = int(agg.iloc[top_n:][count_col].sum())
            top = pd.concat([top, pd.DataFrame({label_col: ["Other"], count_col: [other_val]})], ignore_index=True)

        pie_df = top.rename(columns={label_col: "Flag", count_col: "Count"})

        pie = (
            alt.Chart(pie_df)
            .mark_arc(innerRadius=40)
            .encode(
                theta=alt.Theta("Count:Q", stack=True),
                color=alt.Color("Flag:N", legend=alt.Legend(orient="bottom", title=None)),
                tooltip=["Flag:N", alt.Tooltip("Count:Q", format=",")],
            )
            .properties(height=320)
        )

        labels = (
            alt.Chart(pie_df)
            .mark_text(radius=120, size=12)
            .encode(
                theta=alt.Theta("Count:Q", stack=True),
                text=alt.Text("Count:Q", format=","),
            )
        )

        st.altair_chart(pie + labels, use_container_width=True)
        st.caption("Pie shows top flags by count (top 10 + Other). Values are tutor counts.")
    except Exception as e:
        st.warning(f"Couldn't chart Special Certification Flags (showing table instead): {e}")
        st.dataframe(flags_raw, use_container_width=True, hide_index=True)

# -----------------------------
# Coverage by grade level
# -----------------------------
st.header("Coverage by Grade Level")

if sheets and "Coverage Matrix" in sheets:
    cov = clean_excel_df(sheets["Coverage Matrix"])
    cert = clean_excel_df(sheets["Certified Coverage Matrix"]) if (sheets and "Certified Coverage Matrix" in sheets) else None

    subject_col = cov.columns[0]
    grade_cols_raw = list(cov.columns[1:])
    grade_cols = sort_grade_cols(grade_cols_raw)

    subject = st.selectbox("Select subject area", sorted(cov[subject_col].dropna().unique().tolist()))
    view_mode = st.radio("View", ["Grades", "Grade bands (K-5 / 6-8 / 9-12)"], horizontal=True, index=0)

    total_row = cov.loc[cov[subject_col] == subject, grade_cols].iloc[0].fillna(0)
    total_row = pd.to_numeric(total_row, errors="coerce").fillna(0).astype(int)

    show_cert = st.checkbox("Show certified coverage overlay", value=True)

    if view_mode == "Grades":
        grade_order = ["K"] + [str(i) for i in range(1, 13)]
        df_plot = pd.DataFrame({
            "Grade": [display_grade_label(g) for g in grade_cols],
            "Total Coverage": total_row.values,
        })
        if show_cert and cert is not None and subject in cert[subject_col].values:
            cert_row = cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0)
            cert_row = pd.to_numeric(cert_row, errors="coerce").fillna(0).astype(int)
            df_plot["Certified Coverage"] = cert_row.values

        # Force correct grade ordering (K,1,2,...,12) in the x-axis
        df_plot["Grade"] = pd.Categorical(df_plot["Grade"], categories=grade_order, ordered=True)

        value_cols = [c for c in df_plot.columns if c != "Grade"]
        df_long = df_plot.melt(id_vars=["Grade"], value_vars=value_cols, var_name="Series", value_name="Count")

        base = (
            alt.Chart(df_long)
            .encode(
                x=alt.X("Grade:N", sort=grade_order, title="Grade"),
                y=alt.Y("Count:Q", title="Tutor count"),
                color=alt.Color("Series:N", legend=alt.Legend(orient="top")),
                xOffset="Series:N",
                tooltip=["Grade:N", "Series:N", "Count:Q"],
            )
        )

        bars = base.mark_bar()
        labels = base.mark_text(dy=-8).encode(text=alt.Text("Count:Q", format=","))

        chart = (bars + labels).properties(height=360)
        st.altair_chart(chart, use_container_width=True)
    else:
        total_band = make_grade_band_series(pd.Series(total_row.values, index=grade_cols))
        band_df = pd.DataFrame({"Total Coverage": total_band.values}, index=total_band.index)
        if show_cert and cert is not None and subject in cert[subject_col].values:
            cert_row = cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0)
            cert_row = pd.to_numeric(cert_row, errors="coerce").fillna(0).astype(int)
            cert_band = make_grade_band_series(pd.Series(cert_row.values, index=grade_cols))
            band_df["Certified Coverage"] = cert_band.values
        band_plot = band_df.reset_index().rename(columns={"index": "Band"})
        value_cols = [c for c in band_plot.columns if c != "Band"]
        band_long = band_plot.melt(id_vars=["Band"], value_vars=value_cols, var_name="Series", value_name="Count")

        band_base = (
            alt.Chart(band_long)
            .encode(
                x=alt.X("Band:N", title="Grade band"),
                y=alt.Y("Count:Q", title="Tutor count"),
                color=alt.Color("Series:N", legend=alt.Legend(orient="top", title=None)),
                xOffset="Series:N",
                tooltip=["Band:N", "Series:N", "Count:Q"],
            )
        )
        band_bars = band_base.mark_bar()
        band_labels = band_base.mark_text(dy=-8).encode(text=alt.Text("Count:Q", format=","))

        st.altair_chart((band_bars + band_labels).properties(height=320), use_container_width=True)

    st.caption("Counts are unique inactive tutors from the executive workbook matrices.")
else:
    st.info("Coverage Matrix sheet not found in the workbook.")

# -----------------------------
# Coverage by language
# -----------------------------
st.header("Coverage by Language")

if tutor_long.empty:
    st.info("Language coverage requires parsed_tutor_data.json.")
else:
    # Do NOT drop_duplicates before exploding (languages is a list -> unhashable).
    exploded = tutor_long[["tutor_id", "languages"]].explode("languages")
    exploded["languages"] = exploded["languages"].fillna("").astype(str).str.strip()
    exploded = exploded[exploded["languages"].ne("")]

    # Unique tutor-language pairs, then count unique tutors per language
    exploded = exploded.drop_duplicates(subset=["tutor_id", "languages"])
    lang_counts = exploded.groupby("languages")["tutor_id"].nunique().sort_values(ascending=False)

    top_n = st.slider(
        "Show top N languages",
        min_value=5,
        max_value=50,
        value=min(20, len(lang_counts)) if len(lang_counts) else 5,
        step=5,
    )
    lang_df = lang_counts.head(top_n).reset_index()
    lang_df.columns = ["Language", "Tutors"]

    lang_base = (
        alt.Chart(lang_df)
        .encode(
            x=alt.X("Language:N", sort="-y", title=None),
            y=alt.Y("Tutors:Q", title="Unique tutors"),
            tooltip=["Language:N", alt.Tooltip("Tutors:Q", format=",")],
        )
    )
    lang_bars = lang_base.mark_bar()
    lang_labels = lang_base.mark_text(dy=-8).encode(text=alt.Text("Tutors:Q", format=","))

    st.altair_chart((lang_bars + lang_labels).properties(height=360), use_container_width=True)
    st.caption("Y-axis = unique inactive tutors who report speaking the language.")

# -----------------------------
# Math Specialty Coverage + drill-down
# -----------------------------
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
        ms_base = (
            alt.Chart(plot_df)
            .encode(
                x=alt.X(f"{label_col}:N", sort="-y", title=None),
                y=alt.Y(f"{count_col}:Q", title="Unique tutors"),
                tooltip=[alt.Tooltip(f"{label_col}:N", title="Specialty"), alt.Tooltip(f"{count_col}:Q", format=",")],
            )
        )
        ms_bars = ms_base.mark_bar()
        ms_labels = ms_base.mark_text(dy=-8).encode(text=alt.Text(f"{count_col}:Q", format=","))

        st.altair_chart((ms_bars + ms_labels).properties(height=360), use_container_width=True)
        st.caption("X-axis: math specialty label. Y-axis: unique inactive tutor counts.")

        if not tutor_long.empty:
            st.subheader("Math Specialty Drill-down")
            specialty = st.selectbox("Select a math specialty to view tutors", plot_df[label_col].tolist())

            mflt = tutor_long[(tutor_long["coverage_subject"] == "Math") & (tutor_long["math_specialty"] == specialty)].copy()

            if f_grades:
                mflt = mflt[mflt["grade"].isin(f_grades)]
            if f_langs:
                mflt = mflt.loc[mask_has_any_language(mflt, f_langs)]
            if require_math:
                mflt = mflt[mflt["has_math_cert"] == True]
            if require_sped:
                mflt = mflt[mflt["has_sped_cert"] == True]
            if require_ela:
                mflt = mflt[mflt["has_ela_cert"] == True]
            if require_ir:
                mflt = mflt[mflt["has_ir_cert"] == True]
            if require_spanish:
                mflt = mflt[mflt["has_spanish_cert"] == True]
            if search:
                mflt = mflt[mflt["name"].fillna("").str.contains(search, case=False)]

            st.write(f"Matching tutor-grade rows: **{len(mflt):,}**")
            st.write(f"Unique tutors: **{mflt['tutor_id'].nunique():,}**")

            tutors_df = (
                mflt.groupby(["tutor_id", "name"], as_index=False)
                .agg(
                    grades=("grade", lambda g: ", ".join(map(str, sorted(set(map(int, g)))))),
                    languages=("languages_str", "first"),
                    certs=("cert_subjects", lambda c: ", ".join(sorted(set(sum(c, []))))),
                )
                .sort_values("name")
            )

            st.dataframe(tutors_df, use_container_width=True, hide_index=True)

            specialty_safe = re.sub(r"[^A-Za-z0-9]+", "_", str(specialty)).strip("_").lower() or "specialty"
            out = io.BytesIO()
            tutors_df.to_excel(out, index=False, engine="openpyxl")
            st.download_button(
                "Download this specialty list (Excel)",
                data=out.getvalue(),
                file_name=f"math_specialty_{specialty_safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Math Specialty Coverage sheet not found in the workbook.")

# -----------------------------
# Tutor lookup + export
# -----------------------------
st.header("Tutor Lookup")

if tutor_long.empty:
    st.info("Tutor lookup is disabled until parsed_tutor_data.json is present and readable.")
else:
    flt = tutor_long.copy()

    if f_subjects:
        flt = flt[flt["coverage_subject"].isin(f_subjects)]
    if f_grades:
        flt = flt[flt["grade"].isin(f_grades)]

    if f_specs:
        flt = flt[flt["math_specialty"].isin(f_specs)]
    if f_langs:
        flt = flt.loc[mask_has_any_language(flt, f_langs)]

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
