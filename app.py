import streamlit as st
import pandas as pd
import altair as alt
import json
import io
import re
from pathlib import Path

EXCEL_FILE = "Inactive_Tutor_Executive_Report_v7_FULL_FINAL.xlsx"
JSON_FILE  = "parsed_tutor_data.json"

st.set_page_config(page_title="Tutor Dashboard - Summaries & Lookup", layout="wide")
st.title("Tutor Dashboard - Summaries & Lookup (Inactive Pool)")

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
        # Email is stored at top-level in parsed_tutor_data.json (preferred).
        # If missing, try to pull it from raw_data fallbacks.
        email = (t.get("email")
                 or (t.get("raw_data", {}) or {}).get("personal_email")
                 or (t.get("raw_data", {}) or {}).get("email")
                 or (t.get("raw_data", {}) or {}).get("email_address"))

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
                    "email": email,
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


def _bar_with_value_labels(df: pd.DataFrame, x: str, y: str, *, color: str | None = None, sort=None, height: int = 320, title: str | None = None):
    """Altair bar chart with numeric labels above bars."""
    enc = {
        "x": alt.X(f"{x}:N", sort=sort, title=title or x),
        "y": alt.Y(f"{y}:Q", title=y),
        "tooltip": [alt.Tooltip(f"{x}:N"), alt.Tooltip(f"{y}:Q")],
    }
    if color:
        enc["color"] = alt.Color(f"{color}:N", legend=alt.Legend(orient="top"))
        enc["tooltip"].insert(1, alt.Tooltip(f"{color}:N"))

    base = alt.Chart(df)
    bars = base.mark_bar().encode(**enc)
    text = base.mark_text(dy=-6).encode(
        x=enc["x"],
        y=enc["y"],
        text=alt.Text(f"{y}:Q", format=",.0f"),
    )
    return (bars + text).properties(height=height)

def _grouped_bar_with_labels(df_long: pd.DataFrame, x: str, y: str, group: str, *, sort=None, height: int = 320, x_title: str | None = None, y_title: str | None = None):
    """Altair grouped bar chart (xOffset) with labels."""
    base = alt.Chart(df_long)
    bars = (
        base.mark_bar()
        .encode(
            x=alt.X(f"{x}:N", sort=sort, title=x_title or x),
            y=alt.Y(f"{y}:Q", title=y_title or y),
            color=alt.Color(f"{group}:N", legend=alt.Legend(orient="top")),
            xOffset=f"{group}:N",
            tooltip=[f"{x}:N", f"{group}:N", f"{y}:Q"],
        )
    )
    text = (
        base.mark_text(dy=-6)
        .encode(
            x=alt.X(f"{x}:N", sort=sort),
            y=alt.Y(f"{y}:Q"),
            xOffset=f"{group}:N",
            text=alt.Text(f"{y}:Q", format=",.0f"),
        )
    )
    return (bars + text).properties(height=height)

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
# UI polish (CSS)
# -----------------------------
st.markdown(
    """
<style>
/* reduce top padding a bit */
.block-container { padding-top: 2.75rem; padding-bottom: 3rem; }

/* sticky summary bar */
.sticky-summary {
  position: sticky;
  top: 3.25rem;
  z-index: 999;
  backdrop-filter: blur(8px);
  background: rgba(15, 17, 22, 0.75);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 18px;
  padding: 12px 14px;
  margin: 10px 0 14px 0;
}

/* filter card */
.filter-card {
  border: 1px solid rgba(255,255,255,0.08);
  background: rgba(255,255,255,0.03);
  border-radius: 18px;
  padding: 14px 14px 6px 14px;
  margin: 6px 0 12px 0;
}

/* chips */
.chips { display:flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
.chip {
  font-size: 0.85rem;
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid rgba(255,255,255,0.14);
  background: rgba(255,255,255,0.05);
  opacity: 0.95;
}
.muted { opacity: 0.75; font-size: 0.9rem; }
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Lookup filter helpers (filters live in the Tutor Lookup tab)
# -----------------------------
def apply_lookup_filters(
    df: pd.DataFrame,
    *,
    search: str = "",
    f_subjects: list[str] | None = None,
    f_grades: list[int] | None = None,
    f_specs: list[str] | None = None,
    f_langs: list[str] | None = None,
    require_ela: bool = False,
    require_math: bool = False,
    require_sped: bool = False,
    require_ir: bool = False,
    require_spanish: bool = False,
) -> pd.DataFrame:
    """Apply the Tutor Lookup filters to a tutor_long-style dataframe."""
    if df.empty:
        return df

    out = df.copy()

    if f_subjects:
        out = out[out["coverage_subject"].isin(f_subjects)]
    if f_grades:
        out = out[out["grade"].isin(f_grades)]
    if f_specs:
        out = out[out["math_specialty"].isin(f_specs)]
    if f_langs:
        out = out.loc[mask_has_any_language(out, f_langs)]

    if require_ela:
        out = out[out["has_ela_cert"] == True]
    if require_math:
        out = out[out["has_math_cert"] == True]
    if require_sped:
        out = out[out["has_sped_cert"] == True]
    if require_ir:
        out = out[out["has_ir_cert"] == True]
    if require_spanish:
        out = out[out["has_spanish_cert"] == True]

    if search:
        out = out[out["name"].fillna("").str.contains(str(search).strip(), case=False)]

    return out


def get_lookup_filter_options(tutor_long: pd.DataFrame):
    """Build option lists for the Tutor Lookup filters."""
    subjects_all = sorted(tutor_long["coverage_subject"].dropna().unique().tolist())
    grades_all = sorted([int(x) for x in tutor_long["grade"].dropna().unique().tolist()])
    grade_tokens_all = ["ES", "MS", "HS"] + [grade_int_to_label(g) for g in grades_all]
    default_tokens = [grade_int_to_label(g) for g in grades_all]

    specs_all = sorted(tutor_long["math_specialty"].dropna().unique().tolist())
    lang_set = set()
    for ls in tutor_long["languages"].tolist():
        for l in (ls or []):
            lang_set.add(str(l).strip())
    langs_all = sorted([l for l in lang_set if l])

    return subjects_all, grade_tokens_all, default_tokens, specs_all, langs_all


# -----------------------------
# Sticky summary (top)
# -----------------------------
if tutor_long.empty:
    total_unique = 0
    total_math = 0
    total_ela = 0
    total_sped = 0
else:
    total_unique = int(tutor_long["tutor_id"].nunique())
    total_math = int(tutor_long.loc[tutor_long["coverage_subject"] == "Math", "tutor_id"].nunique())
    total_ela = int(tutor_long.loc[tutor_long["coverage_subject"] == "ELA/Literacy", "tutor_id"].nunique())
    total_sped = int(tutor_long.loc[tutor_long["has_sped_cert"] == True, "tutor_id"].nunique())

st.markdown('<div class="sticky-summary">', unsafe_allow_html=True)
m1, m2, m3, m4 = st.columns([1, 1, 1, 1])
m1.metric("Total inactive tutors", f"{total_unique:,}" if total_unique else "—")
m2.metric("Math coverage", f"{total_math:,}" if total_unique else "—")
m3.metric("ELA/Lit coverage", f"{total_ela:,}" if total_unique else "—")
m4.metric("SPED certified", f"{total_sped:,}" if total_unique else "—")

# -----------------------------
# Tabs
# -----------------------------
tab_overview, tab_lookup = st.tabs(["Coverage Overview", "Tutor Filter & Lookup"])

with tab_overview:
    # -----------------------------
    # Coverage by grade level
    # -----------------------------

    st.subheader("Coverage by Grade Level")

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


            chart = _grouped_bar_with_labels(
                df_long,
                x="Grade",
                y="Count",
                group="Series",
                sort=grade_order,
                height=360,
                x_title="Grade",
                y_title="Tutor count",
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            total_band = make_grade_band_series(pd.Series(total_row.values, index=grade_cols))
            band_df = pd.DataFrame({"Total Coverage": total_band.values}, index=total_band.index)
            
            if show_cert and cert is not None and subject in cert[subject_col].values:
                cert_row = cert.loc[cert[subject_col] == subject, grade_cols].iloc[0].fillna(0)
                cert_row = pd.to_numeric(cert_row, errors="coerce").fillna(0).astype(int)
                cert_band = make_grade_band_series(pd.Series(cert_row.values, index=grade_cols))
                band_df["Certified Coverage"] = cert_band.values
            
            band_plot = band_df.reset_index()

# Ensure first column is named Band (avoids KeyError in melt when index isn't called "index")

if band_plot.columns[0] != "Band":
    band_plot = band_plot.rename(columns={band_plot.columns[0]: "Band"})

band_long = band_plot.melt(
    id_vars=["Band"],
    var_name="Series",
    value_name="Count"
)
chart_band = _grouped_bar_with_labels(
    band_long, 
    x="Band", 
    y="Count", 
    group="Series", 
    sort=["K-5","6-8","9-12","Other"], 
    height=320, 
    x_title="Grade band", 
    y_title="Tutor count"
)
st.altair_chart(chart_band, use_container_width=True)
st.caption("Individual grades show unique tutors. Grade bands are culumlative and will include overlap (e.g. a tutor that is certified in K-5 will be represented five times in that grade band).")
    
else:
    st.info("Coverage Matrix sheet not found in the workbook.")

    # -----------------------------
    # Math Specialty Coverage 
    # -----------------------------

    st.divider()
    st.subheader("Coverage by Language")

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
        chart_lang = _bar_with_value_labels(lang_df, x="Language", y="Tutors", sort=lang_df["Language"].tolist(), height=320, title="Language")
        st.altair_chart(chart_lang, use_container_width=True)
        st.caption("Y-axis = Unique tutors who report speaking the language.")

    # -----------------------------
    # Special Certification Flags
    st.divider()
    st.subheader("Math Specialty Coverage")

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

            chart_ms = _bar_with_value_labels(plot_df, x=label_col, y=count_col, sort=plot_df[label_col].tolist(), height=320, title="Math specialty")
            st.altair_chart(chart_ms, use_container_width=True)
            st.caption("X-axis: Math specialty. Y-axis: Unique tutor count")

    else:
        st.info("Math Specialty Coverage sheet not found in the workbook.")
    
    # -----------------------------
    # Coverage by language
    # -----------------------------

    # -----------------------------

    if sheets and "Special Certification Flags" in sheets:
        st.divider()
        st.subheader("Special Certification Flags")
        flags_raw = clean_excel_df(sheets["Special Certification Flags"])
        # Drop any real "index" columns that may exist in the sheet
        flags_raw = flags_raw.loc[:, ~flags_raw.columns.astype(str).str.match(r"(?i)^index(\.|$)")]
        if flags_raw.shape[1] < 2:
            st.dataframe(flags_raw, use_container_width=True, hide_index=True)
        else:
            label_col, count_col, flags = pick_label_and_count_columns(flags_raw)
            flags[label_col] = flags[label_col].astype(str).str.strip()
            flags[count_col] = pd.to_numeric(flags[count_col], errors="coerce").fillna(0).astype(int)
            plot_df = (
                flags[[label_col, count_col]]
                .groupby(label_col, as_index=False)[count_col].sum()
                .sort_values(count_col, ascending=False)
            )
            # Keep chart readable: top 10 + "Other"
            top_k = 10
            if len(plot_df) > top_k:
                top = plot_df.head(top_k).copy()
                other_sum = int(plot_df.iloc[top_k:][count_col].sum())
                other = pd.DataFrame([{label_col: "Other", count_col: other_sum}])
                plot_df = pd.concat([top, other], ignore_index=True)

            base = alt.Chart(plot_df).encode(
    theta=alt.Theta(f"{count_col}:Q"),
    color=alt.Color(f"{label_col}:N", legend=alt.Legend(orient="right")),
    tooltip=[
        alt.Tooltip(f"{label_col}:N", title="Flag"),
        alt.Tooltip(f"{count_col}:Q", title="Tutors"),
    ],
)

arcs = base.mark_arc()

labels = base.mark_text(radius=90).encode(
    text=alt.Text(f"{count_col}:Q", format=",.0f")
)

st.altair_chart(arcs + labels, use_container_width=True)
st.caption("Pie shows unique inactive tutor counts by flag (top 10 + Other).")
    else:
        st.divider()
        st.subheader("Special Certification Flags not found in the workbook")

with tab_lookup:
    # -----------------------------
    # Tutor lookup + export
    # -----------------------------
    st.subheader("Tutor Filters & Lookup")

    if tutor_long.empty:
        st.info("Tutor lookup is disabled until parsed_tutor_data.json is present and readable.")
    else:
        # Build filter options
        subjects_all, grade_tokens_all, default_tokens, specs_all, langs_all = get_lookup_filter_options(tutor_long)

        # Filters (apply only to this tab)
        st.markdown('<div class="filter-card">', unsafe_allow_html=True)
        st.subheader("Filters", anchor=False)

        export_slot = None


        cA, cB, cC, cD = st.columns([1.2, 1.2, 1.0, 0.6])
        with cD:
            if st.button("Reset", use_container_width=True):
                for k in [
                    "lookup_search", "lookup_f_subjects", "lookup_selected_tokens",
                    "lookup_f_langs", "lookup_f_specs",
                    "lookup_require_ela", "lookup_require_math", "lookup_require_sped",
                    "lookup_require_ir", "lookup_require_spanish",
                ]:
                    st.session_state.pop(k, None)
                st.rerun()

            export_slot = st.empty()


        with cA:
            search = st.text_input(
                "Search tutor name",
                value=st.session_state.get("lookup_search", ""),
            ).strip()
            st.session_state["lookup_search"] = search

            f_subjects = st.multiselect(
                "Coverage subject",
                subjects_all,
                default=st.session_state.get("lookup_f_subjects", subjects_all),
            )
            st.session_state["lookup_f_subjects"] = f_subjects

        with cB:
            selected_tokens = st.multiselect(
                "Grade (ES/MS/HS or individual)",
                grade_tokens_all,
                default=st.session_state.get("lookup_selected_tokens", default_tokens),
            )
            st.session_state["lookup_selected_tokens"] = selected_tokens

            expanded = set()
            for tok in selected_tokens:
                for g in grade_token_to_grades(tok):
                    expanded.add(int(g))
            f_grades = sorted(expanded)

            f_langs = st.multiselect(
                "Language spoken (optional)",
                langs_all,
                default=st.session_state.get("lookup_f_langs", []),
            )
            st.session_state["lookup_f_langs"] = f_langs

        with cC:
            f_specs = st.multiselect(
                "Math specialty (optional)",
                specs_all,
                default=st.session_state.get("lookup_f_specs", []),
            )
            st.session_state["lookup_f_specs"] = f_specs

            st.caption("Certification requirements")
            require_ela = st.checkbox("Require ELA cert", value=st.session_state.get("lookup_require_ela", False))
            require_math = st.checkbox("Require Math cert", value=st.session_state.get("lookup_require_math", False))
            require_sped = st.checkbox("Require SPED cert", value=st.session_state.get("lookup_require_sped", False))
            require_ir = st.checkbox("Require IR cert", value=st.session_state.get("lookup_require_ir", False))
            require_spanish = st.checkbox("Require Spanish cert", value=st.session_state.get("lookup_require_spanish", False))

            st.session_state["lookup_require_ela"] = require_ela
            st.session_state["lookup_require_math"] = require_math
            st.session_state["lookup_require_sped"] = require_sped
            st.session_state["lookup_require_ir"] = require_ir
            st.session_state["lookup_require_spanish"] = require_spanish

        st.markdown("</div>", unsafe_allow_html=True)

        # Apply filters
        flt = apply_lookup_filters(
            tutor_long,
            search=search,
            f_subjects=f_subjects,
            f_grades=f_grades,
            f_specs=f_specs,
            f_langs=f_langs,
            require_ela=require_ela,
            require_math=require_math,
            require_sped=require_sped,
            require_ir=require_ir,
            require_spanish=require_spanish,
        )

        # Summary for this tab
        unique_filtered = int(flt["tutor_id"].nunique()) if not flt.empty else 0
        chips = []
        if search:
            chips.append(f"Search: {search}")
        if f_subjects and len(f_subjects) != len(subjects_all):
            chips.append(f"Subjects: {len(f_subjects)}")
        if f_grades:
            chips.append(f"Grades: {len(f_grades)}")
        if f_langs:
            chips.append("Lang: " + ", ".join(f_langs[:3]) + ("…" if len(f_langs) > 3 else ""))
        if f_specs:
            chips.append("Spec: " + ", ".join(f_specs[:2]) + ("…" if len(f_specs) > 2 else ""))
        if require_ela: chips.append("Req: ELA")
        if require_math: chips.append("Req: Math")
        if require_sped: chips.append("Req: SPED")
        if require_ir: chips.append("Req: IR")
        if require_spanish: chips.append("Req: Spanish")

        st.caption(" | ".join(chips) if chips else "No filters applied.")

        # Build export table (one row per tutor). Keep tutor_id for export, but hide it in the on-page table.
        agg_map = {
            "subjects": ("coverage_subject", lambda s: ", ".join(sorted(set(s)))),
            "grades": ("grade", lambda g: ", ".join(map(str, sorted(set(map(int, g)))))),
            "specialties": ("math_specialty", lambda s: ", ".join(sorted(set([x for x in s if pd.notna(x)])))),
            "languages": ("languages_str", "first"),
            "certs": ("cert_subjects", lambda c: ", ".join(sorted(set(sum(c, []))))),
        }
        if "email" in flt.columns:
            agg_map["email"] = ("email", "first")

        tutors_df = (
            flt.groupby(["tutor_id", "name"], as_index=False)
            .agg(**agg_map)
            .sort_values("name")
        )

        st.subheader("Results")
        st.write(f"Matching tutor-grade rows: **{len(flt):,}**")
        st.write(f"Unique tutors: **{len(tutors_df):,}**")

        display_df = tutors_df.drop(columns=["tutor_id"], errors="ignore")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # Export button lives in the Filters panel (top-right) for visibility.
        out = io.BytesIO()
        tutors_df.to_excel(out, index=False, engine="openpyxl")
        if export_slot is not None:
            export_slot.download_button(
                "Download (Excel)",
                data=out.getvalue(),
                file_name="filtered_inactive_tutors.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.download_button(
                "Download filtered tutors (Excel)",
                data=out.getvalue(),
                file_name="filtered_inactive_tutors.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
