from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# 1. PAGE CONFIGURATION AND FILE PATHS
# ============================================================

st.set_page_config(
    page_title="The Invisible Workforce | Lebanon",
    page_icon="📊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

def find_input_file(expected_name, keywords):
    """Find an Excel workbook in the project folder or its data subfolder."""
    candidates = []
    for root in (DATA_DIR, BASE_DIR):
        if root.exists():
            candidates.extend(
                p for p in root.rglob("*")
                if p.is_file() and p.suffix.lower() in {".xlsx", ".xls"}
            )
    expected = expected_name.lower().replace(" ", "").replace("-", "").replace("_", "")
    for file in candidates:
        stem = file.stem.lower().replace(" ", "").replace("-", "").replace("_", "")
        if stem == expected:
            return file
    for file in candidates:
        stem = file.stem.lower()
        if all(word.lower() in stem for word in keywords):
            return file
    return None

# Keep these two Excel workbooks beside app.py or inside a folder named data.
LFS_FILE = find_input_file("Labour-force-FULFS-2022", ["labour", "2022"])
HISTORICAL_FILE = find_input_file("unemployment rate women", ["unemployment", "women"])


# ============================================================
# 2. COLORS AND CHART STYLE
# ============================================================

COLORS = {
    "Women": "#D96C5F",
    "Men": "#087E8B",
    "Women & Men": "#D7A84B",
}

STATUS_COLORS = {
    "Employed": "#087E8B",
    "Unemployed": "#D7A84B",
    "Outside labour force": "#D96C5F",
}

PLOT_LAYOUT = {
    "template": "plotly_white",
    "font": {"family": "Arial", "color": "#172A3A"},
    "paper_bgcolor": "rgba(0,0,0,0)",
    "plot_bgcolor": "rgba(0,0,0,0)",
    "margin": {"l": 25, "r": 25, "t": 65, "b": 45},
    "legend_title_text": "",
}


# ============================================================
# 3. PAGE STYLING
# ============================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #F7F8F6;
    }
    .block-container {
        max-width: 1250px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    .hero {
        background: linear-gradient(120deg, #122B3A, #1E5260);
        padding: 2.4rem;
        border-radius: 18px;
        color: white;
        margin-bottom: 1.4rem;
    }
    .hero h1 {
        color: white;
        font-size: 3rem;
        line-height: 1.1;
        margin-bottom: .6rem;
    }
    .hero p {
        color: #E6F0F0;
        font-size: 1.08rem;
    }
    .insight {
        background: white;
        border-left: 4px solid #087E8B;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin: .7rem 0 1rem 0;
    }
    .section-kicker { color:#247C78; text-transform:uppercase; letter-spacing:.12em; font-size:.75rem; font-weight:800; }
    .visual-card { background:#FFFFFF; border:1px solid #E6E3DC; border-radius:15px; padding:1rem 1.1rem; min-height:145px; box-shadow:0 3px 12px rgba(18,59,70,.05); }
    .visual-icon { font-size:1.7rem; line-height:1.3; margin-bottom:.35rem; }
    .visual-card h4 { color:#123B46; margin:.15rem 0 .35rem 0; }
    .visual-card p { color:#52636A; font-size:.92rem; margin:0; }
    .summary-panel { background:linear-gradient(120deg,#123B46,#247C78); color:white; padding:1.25rem 1.4rem; border-radius:16px; margin:.8rem 0 1.2rem 0; }
    .summary-panel h3 { color:white; margin:0 0 .4rem 0; }
    .summary-panel p { color:#F0F6F4; margin:0; }
    .roadmap-step { background:#F4EDE2; border-top:4px solid #247C78; padding:1rem; border-radius:12px; min-height:190px; }
    .roadmap-step .num { color:#D96C5F; font-size:1.55rem; font-weight:800; }
    .photo-credit { color:#66757A; font-size:.78rem; margin-top:-.45rem; }
    .gap-label { font-weight:700; color:#123B46; }
    div[data-testid="stMetric"] { background:#FFFFFF; border:1px solid #E6E3DC; padding: .8rem 1rem; border-radius:12px; }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ============================================================
# 4. HISTORICAL UNEMPLOYMENT DATA
#    Loaded from the user's "unemployment rate women" Excel file
# ============================================================

def load_historical_data():
    """Load World Bank/ILO modelled unemployment series from the supplied workbook."""
    if HISTORICAL_FILE is None or not HISTORICAL_FILE.exists():
        raise FileNotFoundError(
            "Historical workbook not found. Place 'unemployment rate women.xlsx' "
            "in the project folder or its data subfolder."
        )

    # The supplied workbook is a World Bank-style export. Its header is on
    # Excel row 9, with years in column E and total/female/male series in F:H.
    raw = pd.read_excel(HISTORICAL_FILE, sheet_name=0, header=None)
    header_row = None
    year_col = total_col = female_col = male_col = None

    for row_index in range(min(20, len(raw))):
        labels = [
            str(value).strip().lower() if pd.notna(value) else ""
            for value in raw.iloc[row_index].tolist()
        ]
        candidate_year = next(
            (i for i, value in enumerate(labels) if value in {"year", "years"}), None
        )
        candidate_total = next(
            (i for i, value in enumerate(labels) if "unemployment, total" in value), None
        )
        candidate_female = next(
            (i for i, value in enumerate(labels) if "unemployment, female" in value), None
        )
        candidate_male = next(
            (i for i, value in enumerate(labels)
             if "unemployment, male" in value and "female" not in value), None
        )
        if None not in (candidate_year, candidate_total, candidate_female, candidate_male):
            header_row = row_index
            year_col, total_col = candidate_year, candidate_total
            female_col, male_col = candidate_female, candidate_male
            break

    if header_row is None:
        raise ValueError(
            "The historical workbook was found, but its year, total, female and male "
            "unemployment columns could not be identified."
        )

    rows = raw.iloc[header_row + 1:].copy()
    result = pd.DataFrame({
        "Year": pd.to_numeric(rows.iloc[:, year_col], errors="coerce"),
        "Total unemployment": pd.to_numeric(rows.iloc[:, total_col], errors="coerce"),
        "Female unemployment": pd.to_numeric(rows.iloc[:, female_col], errors="coerce"),
        "Male unemployment": pd.to_numeric(rows.iloc[:, male_col], errors="coerce"),
    })
    result = result.dropna(subset=[
        "Year", "Total unemployment", "Female unemployment", "Male unemployment"
    ])
    result["Year"] = result["Year"].astype(int)
    result = result.sort_values("Year").reset_index(drop=True)

    if result.empty:
        raise ValueError("No historical unemployment data rows were found.")

    return result


# 5. SURVEY DATA HELPERS
# ============================================================

def read_section(sheet_name, start_row, end_row):
    """Read a verified four-column section from the survey workbook."""
    raw = pd.read_excel(
        LFS_FILE,
        sheet_name=sheet_name,
        header=None,
    )

    # EA39 and EA40 both store Category, Women, Men and Women & Men
    # contiguously in columns A:D in the supplied workbook.
    table = raw.iloc[start_row:end_row, 0:4].copy()

    table.columns = ["Category", "Women", "Men", "Women & Men"]

    table["Category"] = table["Category"].astype(str).str.strip()

    for column in ["Women", "Men", "Women & Men"]:
        table[column] = pd.to_numeric(
            table[column].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce",
        )

    table = table.dropna(subset=["Women", "Men"])
    table = table[
        ~table["Category"].str.lower().isin(["", "nan", "total"])
    ]

    return table.reset_index(drop=True)


def find_indicator(indicators, keyword):
    """Find an EA1 indicator by a distinctive part of its label."""
    for label, values in indicators.items():
        if keyword.lower() in label.lower():
            return values

    raise KeyError(
        f"Could not find the indicator containing '{keyword}' in EA1."
    )


@st.cache_data(show_spinner=False)
def load_survey():
    if LFS_FILE is None or not LFS_FILE.exists():
        raise FileNotFoundError(
            "2022 labour-force workbook not found. Place "
            "'Labour-force-FULFS-2022.xlsx' in the project folder or data folder."
        )

    available_sheets = pd.ExcelFile(LFS_FILE).sheet_names
    required_sheets = ["EA1", "EA3", "EA6", "EA8", "EA37", "EA39", "EA40"]
    missing_sheets = [
        sheet for sheet in required_sheets
        if sheet not in available_sheets
    ]

    if missing_sheets:
        raise ValueError(
            "The survey workbook is missing sheets: "
            + ", ".join(missing_sheets)
        )

    # EA1: Main labour-force indicators. Rows 6-11, columns A-D.
    headline_raw = pd.read_excel(
        LFS_FILE,
        sheet_name="EA1",
        header=None,
    )

    indicators = {}

    for row_index in range(6, 12):
        label = str(headline_raw.iloc[row_index, 0]).strip()

        if not label or label.lower() == "nan":
            continue

        indicators[label] = {
            "Women": pd.to_numeric(
                headline_raw.iloc[row_index, 1], errors="coerce"
            ),
            "Men": pd.to_numeric(
                headline_raw.iloc[row_index, 2], errors="coerce"
            ),
            "Women & Men": pd.to_numeric(
                headline_raw.iloc[row_index, 3], errors="coerce"
            ),
        }

    # EA40: Age rows 6-16; education rows 18-24.
    age_lfp = read_section("EA40", 6, 17)
    education_lfp = read_section("EA40", 18, 25)

    # EA39: Age rows 6-16; education rows 18-24.
    age_unemployment = read_section("EA39", 6, 17)
    education_unemployment = read_section("EA39", 18, 25)

    # EA6A: Employed population by completed education, rows 6-12.
    employed_education = read_section("EA6", 6, 13)

    # EA8A: Employment status among employed people, rows 7-11.
    employment_status = read_section("EA8", 7, 12)

    # EA37A: Job-search methods, rows 7-11.
    job_search = read_section("EA37", 7, 12)

    # EA3B: Economic status by sex, rows 16-17.
    status_raw = pd.read_excel(
        LFS_FILE,
        sheet_name="EA3",
        header=None,
    )

    status = status_raw.iloc[16:18, 0:4].copy()
    status.columns = [
        "Gender",
        "Employed",
        "Unemployed",
        "Outside labour force",
    ]

    status["Gender"] = status["Gender"].astype(str).str.strip()

    for column in ["Employed", "Unemployed", "Outside labour force"]:
        status[column] = pd.to_numeric(status[column], errors="coerce")

    status = status.dropna(
        subset=["Employed", "Unemployed", "Outside labour force"]
    )

    return {
        "indicators": indicators,
        "age_lfp": age_lfp,
        "education_lfp": education_lfp,
        "age_unemployment": age_unemployment,
        "education_unemployment": education_unemployment,
        "employed_education": employed_education,
        "employment_status": employment_status,
        "job_search": job_search,
        "status": status,
    }


# ============================================================
# 6. GENERAL CHART HELPERS
# ============================================================

def to_long(data, value_name, genders=None):
    """Reshape gender columns for Plotly, respecting the selected gender filter."""
    selected = ["Women", "Men"] if genders is None else genders
    selected = [gender for gender in selected if gender in data.columns]
    if not selected:
        return pd.DataFrame(columns=["Category", "Gender", value_name])
    return data.melt(
        id_vars="Category",
        value_vars=selected,
        var_name="Gender",
        value_name=value_name,
    )


def draw_gender_bar(data, value_name, title, y_title, genders=None):
    if data.empty:
        st.warning("No observations are available for this chart. Check the source worksheet and selected filters.")
        return
    long_data = to_long(data, value_name, genders)
    if long_data.empty:
        st.info("Select at least one gender to display this chart.")
        return
    fig = px.bar(
        long_data,
        x="Category",
        y=value_name,
        color="Gender",
        barmode="group",
        color_discrete_map=COLORS,
        title=title,
        labels={"Category": "Category", value_name: y_title},
    )
    fig.update_layout(
        **PLOT_LAYOUT,
        xaxis_title="",
        yaxis_title=y_title,
        xaxis_tickangle=-12,
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, key=f"gender_bar_{title}")


def draw_gender_line(data, value_name, title, y_title, genders=None):
    if data.empty:
        st.warning("No observations are available for this chart. Check the source worksheet and selected filters.")
        return
    long_data = to_long(data, value_name, genders)
    if long_data.empty:
        st.info("Select at least one gender to display this chart.")
        return
    fig = px.line(
        long_data,
        x="Category",
        y=value_name,
        color="Gender",
        markers=True,
        color_discrete_map=COLORS,
        title=title,
        labels={"Category": "Category", value_name: y_title},
    )
    fig.update_layout(
        **PLOT_LAYOUT,
        xaxis_title="",
        yaxis_title=y_title,
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, key=f"gender_line_{title}")


def top_category(data, column, highest=True):
    """Return the category and value at the maximum/minimum, or None."""
    if data.empty or column not in data.columns:
        return None
    usable = data.dropna(subset=[column])
    if usable.empty:
        return None
    row = usable.loc[usable[column].idxmax() if highest else usable[column].idxmin()]
    return str(row["Category"]), float(row[column])


# ============================================================
# 7. LOAD AND VALIDATE DATA
# ============================================================

try:
    historical = load_historical_data()
    survey = load_survey()

    indicators = survey["indicators"]
    age_lfp = survey["age_lfp"]
    education_lfp = survey["education_lfp"]
    age_unemployment = survey["age_unemployment"]
    education_unemployment = survey["education_unemployment"]
    employed_education = survey["employed_education"]
    employment_status = survey["employment_status"]
    job_search = survey["job_search"]
    status_by_gender = survey["status"]

    lfp = find_indicator(indicators, "Labour Force Participation Rate")
    unemployment = find_indicator(indicators, "LU1: Unemployment rate")

except Exception as error:
    st.error(
        "The dashboard could not load the data. "
        "Check the workbook filename and worksheet structure."
    )
    st.exception(error)
    st.stop()


# ============================================================
# 8. HEADLINE CALCULATIONS
# ============================================================

participation_gap = lfp["Men"] - lfp["Women"]
unemployment_gap = unemployment["Women"] - unemployment["Men"]


# ============================================================
# 9. DASHBOARD HEADER AND HEADLINE METRICS
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div style="
            letter-spacing: .14em;
            text-transform: uppercase;
            font-size: .75rem;
            color: #9DD4D3;
            font-weight: 700;
        ">
            Lebanon · Labour market · 2022
        </div>
        <h1>The Invisible Workforce</h1>
        <p>
            <b>Understanding gender, education and paid employment in Lebanon.</b>
            <br>
            Explore how labour-force participation and unemployment
            differ between women and men, and how these patterns
            vary by age and education.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write(
    "This interactive research blog examines gender differences "
    "in Lebanon's labour market using the 2022 Labour Force and "
    "Household Living Conditions Survey and the historical "
    "unemployment table supplied with the project."
)

st.markdown(
    f"""
    <div class="summary-panel">
      <div class="section-kicker" style="color:#B9DFD8">Executive summary · Lebanon 2022</div>
      <h3>What is this research about?</h3>
      <p>This research examines how gender, age and education relate to labour-force participation and unemployment in Lebanon. The clearest headline pattern is a {participation_gap:.1f}-percentage-point participation gap between men and women. The analysis also highlights what the available data cannot yet reveal: unpaid care work and differences in individual earnings.</p>
    </div>
    """, unsafe_allow_html=True
)

st.subheader("The headline indicators")

metric1, metric2, metric3, metric4 = st.columns(4)

metric1.metric(
    "Women's labour-force participation",
    f"{lfp['Women']:.1f}%",
)
metric2.metric(
    "Men's labour-force participation",
    f"{lfp['Men']:.1f}%",
)
metric3.metric(
    "Participation gap",
    f"{participation_gap:.1f} pp",
)
metric4.metric(
    "Women's unemployment rate",
    f"{unemployment['Women']:.1f}%",
)

st.markdown(
    f"""
    <div class="insight">
        <b>Key observation:</b>
        Labour-force participation was {lfp['Women']:.1f}% among women
        and {lfp['Men']:.1f}% among men in 2022.
        The difference is {participation_gap:.1f} percentage points.
        Unemployment alone does not capture the full picture because
        people outside the labour force are not counted as unemployed.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 10. NAVIGATION TABS
# ============================================================

story_tab, explore_tab, findings_tab, action_tab, methods_tab = st.tabs(
    [
        "The story",
        "Explore the data",
        "Key findings",
        "From evidence to action",
        "Methodology & sources",
    ]
)


# ============================================================
# 11. THE STORY TAB
# ============================================================

with story_tab:
    st.header("Overview of the problem")
    st.write(
        "Educational attainment does not always translate into labour-force "
        "participation or paid employment in Lebanon, particularly for women. "
        "Despite their qualifications and capabilities, some women remain "
        "outside paid employment or face barriers between education and work. "
        "This raises concerns about the underutilization of women's human "
        "capital and the hidden barriers that may limit their economic participation."
    )

    photo_col, story_col = st.columns([1.05, 1.35], gap="large")
    with photo_col:
        st.image(
            "https://images.pexels.com/photos/5965685/pexels-photo-5965685.jpeg?auto=compress&cs=tinysrgb&w=1200",
            use_container_width=True,
        )
        st.markdown('<div class="photo-credit">Real photograph: Charlotte May / Pexels · students collaborating on a project</div>', unsafe_allow_html=True)
    with story_col:
        st.markdown("<div class='section-kicker'>The context</div>", unsafe_allow_html=True)
        st.markdown("### Education, opportunity and the transition to work")
        st.write("The story is about the distance that can exist between having skills and being able to use them in paid employment. The photograph illustrates the education-to-work journey; it is contextual, not evidence from the Lebanese survey.")

    st.subheader("The problem being addressed")
    st.write(
        "This project investigates how gender, educational attainment and age "
        "are associated with participation in paid employment in Lebanon, and "
        "how these relationships may intersect with marital status and household "
        "composition. It also introduces the gender pay gap as a related dimension "
        "of labour-market inequality. The supplied tables allow us to study "
        "participation and unemployment by gender, age and education; they do not "
        "currently provide the wage, marital-status or household-composition data "
        "needed to measure those dimensions directly."
    )

    problem_a, problem_b, problem_c = st.columns(3)
    with problem_a:
        st.markdown("<div class='visual-card'><div class='visual-icon'>📉</div><h4>Participation gap</h4><p>Women and men do not participate in the labour force at the same rate.</p></div>", unsafe_allow_html=True)
    with problem_b:
        st.markdown("<div class='visual-card'><div class='visual-icon'>🧩</div><h4>Skills not fully used</h4><p>Education alone does not guarantee access to paid employment.</p></div>", unsafe_allow_html=True)
    with problem_c:
        st.markdown("<div class='visual-card'><div class='visual-icon'>🏠</div><h4>Invisible responsibilities</h4><p>Unpaid care may shape opportunities, but is not quantified in these tables.</p></div>", unsafe_allow_html=True)

    st.subheader("Research questions")
    rq1, rq2, rq3 = st.columns(3)
    with rq1:
        st.markdown("<div class='visual-card'><div class='visual-icon'>⚖️</div><h4>Gender</h4><p>How does participation in paid work differ between women and men?</p></div>", unsafe_allow_html=True)
    with rq2:
        st.markdown("<div class='visual-card'><div class='visual-icon'>🎓</div><h4>Education</h4><p>Does educational attainment translate into labour-market participation?</p></div>", unsafe_allow_html=True)
    with rq3:
        st.markdown("<div class='visual-card'><div class='visual-icon'>👥</div><h4>Age and household context</h4><p>How do experiences vary across age groups, and what further data are needed?</p></div>", unsafe_allow_html=True)
    st.caption("Marital status and household composition are part of the broader research question, but are not measured in the supplied extracts.")

    st.markdown("---")
    st.subheader("The invisible workforce: educated yet outside paid employment")
    st.write(
        "The central issue is not simply whether people have completed their "
        "education. It is whether their skills and qualifications are translated "
        "into access to paid work. A person outside the labour force is not counted "
        "as unemployed, so unemployment alone can understate the extent of labour "
        "market exclusion. Unpaid care and household responsibilities may be part "
        "of the wider context, but they are not measured in the supplied tables."
    )

    m1, m2, m3 = st.columns(3)
    m1.metric("Women's labour-force participation", f"{lfp['Women']:.1f}%")
    m2.metric("Men's labour-force participation", f"{lfp['Men']:.1f}%")
    m3.metric("Participation gap", f"{participation_gap:.1f} percentage points")

    st.subheader("Who is this dashboard for?")
    audience1, audience2, audience3 = st.columns(3)
    audience1.markdown("**Young adults and youth**\\n\\nTo understand the transition from education to employment and the barriers facing new labour-market entrants.")
    audience2.markdown("**Women and university students**\\n\\nTo examine how education, age and gender relate to access to paid work and career opportunities.")
    audience3.markdown("**Policymakers and organizations**\\n\\nTo identify evidence gaps and consider policies that connect women's skills with decent paid employment.")

    st.info(
        "Explore the data to investigate the patterns, review the key findings "
        "to understand their implications, and move to From Evidence to Action "
        "for proposed responses."
    )


# ============================================================
# 12. EXPLORE THE DATA TAB
# ============================================================

with explore_tab:
    st.header("Explore the data")
    st.write(
        "Use the controls to investigate the 2022 survey patterns. The charts "
        "are descriptive: they show differences and associations, not causes."
    )

    control1, control2 = st.columns([1, 2])
    with control1:
        selected_genders = st.multiselect(
            "Compare gender groups",
            options=["Women", "Men"],
            default=["Women", "Men"],
            key="explore_gender_filter",
        )
    with control2:
        st.caption("Select one or both groups. Hover over a chart to see exact values. Use the expanders to inspect and download the underlying tables.")

    st.subheader("1. Historical unemployment")
    if not historical.empty:
        year_min = int(historical["Year"].min())
        year_max = int(historical["Year"].max())
        year_range = st.slider(
            "Choose the historical period",
            min_value=year_min,
            max_value=year_max,
            value=(year_min, year_max),
            key="historical_year_range",
        )
        historical_view = historical[historical["Year"].between(year_range[0], year_range[1])]
        available_historical = [
            c for c in ["Total unemployment", "Female unemployment", "Male unemployment"]
            if c in historical_view.columns and historical_view[c].notna().any()
        ]
        historical_long = historical_view.melt(
            id_vars="Year", value_vars=available_historical,
            var_name="Indicator", value_name="Rate",
        )
        historical_colors = {
            "Total unemployment": "#D7A84B",
            "Female unemployment": "#D96C5F",
            "Male unemployment": "#087E8B",
        }
        fig = px.line(
            historical_long, x="Year", y="Rate", color="Indicator",
            markers=True, color_discrete_map=historical_colors,
            title="Unemployment estimates for Lebanon",
            labels={"Rate": "Unemployment rate (%)", "Year": "Year"},
        )
        fig.update_layout(**PLOT_LAYOUT, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True, key="explore_historical_unemployment")
        st.caption("Historical values are World Bank modelled ILO estimates for total, female and male unemployment (1991–2023) from the supplied workbook. The 2022 survey uses a separate national-survey source and methodology, so compare the series cautiously.")
    else:
        st.warning("The historical unemployment table is empty.")

    st.subheader("2. Unemployment by gender in 2022")
    gender_unemployment = pd.DataFrame({
        "Gender": ["Women", "Men"],
        "Rate": [unemployment["Women"], unemployment["Men"]],
    })
    gender_unemployment = gender_unemployment[gender_unemployment["Gender"].isin(selected_genders)]
    if not gender_unemployment.empty:
        fig = px.bar(
            gender_unemployment, x="Gender", y="Rate", color="Gender",
            text_auto=".1f", color_discrete_map=COLORS,
            title="Unemployment among each gender's labour force",
        )
        fig.update_layout(**PLOT_LAYOUT, showlegend=False, xaxis_title="", yaxis_title="Unemployment rate (%)")
        st.plotly_chart(fig, use_container_width=True, key="explore_unemployment_gender_2022")
    else:
        st.info("Select at least one gender above to display this chart.")

    st.subheader("3. Unemployment by age")
    age_options = age_lfp["Category"].tolist()
    selected_ages = st.multiselect(
        "Select age groups", options=age_options, default=age_options,
        key="explore_age_filter",
    )
    filtered_age_unemployment = age_unemployment[age_unemployment["Category"].isin(selected_ages)]
    draw_gender_line(
        filtered_age_unemployment, "Unemployment",
        "Unemployment rates across age groups", "Unemployment rate (%)",
        genders=selected_genders,
    )

    st.subheader("4. Labour-force participation by age")
    filtered_age_lfp = age_lfp[age_lfp["Category"].isin(selected_ages)]
    draw_gender_line(
        filtered_age_lfp, "Participation",
        "Labour-force participation across age groups", "Participation rate (%)",
        genders=selected_genders,
    )

    st.subheader("5. Labour-force participation by education")
    education_options = education_lfp["Category"].tolist()
    selected_education = st.multiselect(
        "Select education levels", options=education_options,
        default=education_options, key="explore_education_filter",
    )
    filtered_education = education_lfp[education_lfp["Category"].isin(selected_education)]
    draw_gender_bar(
        filtered_education, "Participation",
        "Participation by completed education level", "Participation rate (%)",
        genders=selected_genders,
    )

    st.subheader("6. Education composition of employed people")
    draw_gender_bar(
        employed_education, "Share",
        "Educational composition within employed women and men",
        "Share of employed group (%)", genders=selected_genders,
    )
    st.caption("This chart describes the education distribution among employed people. It is not the same as the participation rate among all people with that education level.")

    st.subheader("7. Employment status among employed people")
    draw_gender_bar(
        employment_status, "Share",
        "Employment status within employed women and men",
        "Share within employed group (%)", genders=selected_genders,
    )

    st.subheader("8. Job-search methods")
    draw_gender_bar(
        job_search, "Share",
        "Main job-search methods among unemployed people",
        "Share within unemployed group (%)", genders=selected_genders,
    )

    st.subheader("9. Unemployment by education")
    filtered_education_unemployment = education_unemployment[
        education_unemployment["Category"].isin(selected_education)
    ]
    draw_gender_bar(
        filtered_education_unemployment, "Unemployment",
        "Unemployment rates by completed education level",
        "Unemployment rate (%)", genders=selected_genders,
    )

    with st.expander("Inspect and download the underlying survey tables"):
        tables = {
            "Participation by age": age_lfp,
            "Unemployment by age": age_unemployment,
            "Participation by education": education_lfp,
            "Unemployment by education": education_unemployment,
            "Education composition among employed people": employed_education,
            "Employment status": employment_status,
            "Job-search methods": job_search,
            "Economic status by gender": status_by_gender,
        }
        selected_table_name = st.selectbox("Choose a table", list(tables.keys()), key="source_table_selector")
        selected_table = tables[selected_table_name]
        st.dataframe(selected_table, use_container_width=True, hide_index=True)
        st.download_button(
            "Download selected table as CSV",
            data=selected_table.to_csv(index=False).encode("utf-8-sig"),
            file_name=selected_table_name.lower().replace(" ", "_") + "_2022.csv",
            mime="text/csv",
            key="download_selected_survey_table",
        )


# ============================================================
# 13. KEY FINDINGS TAB
# ============================================================

with findings_tab:
    st.header("Key findings: what does the evidence tell us?")
    st.write("This section interprets the evidence rather than repeating the exploration charts.")

    st.subheader("1. A substantial gender gap in labour-force participation")
    f1, f2, f3 = st.columns(3)
    f1.metric("Women", f"{lfp['Women']:.1f}%")
    f2.metric("Men", f"{lfp['Men']:.1f}%")
    f3.metric("Gender gap", f"{participation_gap:.1f} pp")
    gender_compare = pd.DataFrame({"Group": ["Women", "Men"], "Participation": [lfp["Women"], lfp["Men"]]})
    gap_fig = px.bar(
        gender_compare, x="Participation", y="Group", orientation="h", text="Participation",
        color="Group", color_discrete_map={"Women":"#D96C5F", "Men":"#087E8B"},
        range_x=[0, 100], title="Labour-force participation by gender"
    )
    gap_fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
    gap_fig.update_layout(**{**PLOT_LAYOUT, "height":270, "showlegend":False, "xaxis_title":"Share of each group (%)", "yaxis_title":"", "margin":{"l":10,"r":55,"t":55,"b":35}})
    st.plotly_chart(gap_fig, use_container_width=True, key="finding_gender_gap_visual")
    st.caption(f"The difference is {participation_gap:.1f} percentage points. The scale starts at zero to make the size of the gap visually comparable.")
    st.markdown(
        "The difference indicates that women are substantially less represented "
        "in the labour force than men. This is a participation gap: it includes "
        "people who are outside the labour force, not only people who are unemployed."
    )

    st.subheader("2. Unemployment affects both genders, but the rate is higher among women")
    st.write(
        f"In the 2022 indicators, women's unemployment rate is "
        f"{unemployment['Women']:.1f}%, compared with {unemployment['Men']:.1f}% "
        f"among men — a difference of {unemployment_gap:.1f} percentage points. "
        "The unemployment rate is calculated only among people in the labour force."
    )

    st.subheader("3. Education does not automatically translate into paid employment")
    if not education_lfp.empty:
        female_high = top_category(education_lfp, "Women", highest=True)
        female_low = top_category(education_lfp, "Women", highest=False)
        gaps = education_lfp.copy()
        gaps["Gender gap"] = gaps["Men"] - gaps["Women"]
        largest_gap = top_category(gaps.rename(columns={"Gender gap": "Gap"}), "Gap", highest=True)
        if female_high and female_low:
            st.write(
                f"Among the education categories in the supplied table, women's "
                f"highest participation is recorded for {female_high[0]} "
                f"({female_high[1]:.1f}%), while the lowest is {female_low[0]} "
                f"({female_low[1]:.1f}%). This demonstrates variation across "
                "education categories; it does not mean that a degree guarantees a job."
            )
        if largest_gap:
            st.write(
                f"The largest male–female participation difference across the "
                f"listed education categories is observed for {largest_gap[0]} "
                f"({largest_gap[1]:.1f} percentage points)."
            )

    st.subheader("4. Age is associated with different labour-market experiences")
    age_icons = ["🌱", "🎓", "💼", "📈", "🧭"]
    age_cols = st.columns(5)
    for idx, (_, age_row) in enumerate(age_lfp.iterrows()):
        with age_cols[idx % 5]:
            st.markdown(f"<div class='visual-card'><div class='visual-icon'>{age_icons[idx % len(age_icons)]}</div><h4>{age_row['Category']}</h4><p>Age-group participation profile</p></div>", unsafe_allow_html=True)
    for gender in ["Women", "Men"]:
        high_participation = top_category(age_lfp, gender, highest=True)
        high_unemployment = top_category(age_unemployment, "Women" if gender == "Women" else "Men", highest=True)
        if high_participation:
            st.write(f"**{gender}:** the highest participation rate in the age table is for {high_participation[0]} ({high_participation[1]:.1f}%).")
        if high_unemployment:
            st.write(f"The highest unemployment rate for {gender.lower()} is recorded in {high_unemployment[0]} ({high_unemployment[1]:.1f}%).")

    st.subheader("5. The invisible dimension: unpaid work and the gender pay gap")
    inv1, inv2 = st.columns(2)
    with inv1:
        st.markdown("<div class='visual-card'><div class='visual-icon'>🏡</div><h4>Unpaid care and domestic work</h4><p>Cooking · cleaning · childcare · eldercare. Important work that may constrain time for paid employment, but is not measured in the supplied extracts.</p></div>", unsafe_allow_html=True)
    with inv2:
        st.markdown("<div class='visual-card'><div class='visual-icon'>💶</div><h4>The gender pay gap</h4><p>Comparing average earnings requires wage data and context such as hours, occupation, education and experience.</p></div>", unsafe_allow_html=True)
    st.write(
        "The findings raise questions about unpaid care responsibilities and "
        "whether women's qualifications are fully used in paid employment. "
        "However, the current survey extracts do not measure unpaid household "
        "work or individual wages. The gender pay gap — the difference in average "
        "earnings between women and men — therefore cannot be calculated from "
        "these charts. It requires earnings data and appropriate controls, such "
        "as hours worked, occupation, education and experience."
    )

    st.subheader("What do these findings mean?")
    st.markdown(
        "- **For young women and students:** qualifications matter, but the "
        "transition from education into paid work also depends on access to "
        "opportunities and supportive conditions.\n"
        "- **For employers:** recruitment, pay, promotion and retention should "
        "be reviewed for gender differences.\n"
        "- **For policymakers:** participation, unemployment and people outside "
        "the labour force should be monitored separately.\n"
        "- **For researchers:** wage, time-use, marital-status and household "
        "microdata are needed to test the wider explanations."
    )

    st.warning(
        "The available cross-tabulations do not include marital status, household "
        "composition, unpaid-care hours or individual earnings. These are important "
        "research gaps, not findings that can be inferred from the current charts."
    )


# ============================================================
# 14. FROM EVIDENCE TO ACTION TAB
# ============================================================

with action_tab:
    st.header("From evidence to action")
    st.write(
        "The following are proposed responses to the barriers raised by the "
        "research. They are recommendations for a programme to pilot and evaluate, "
        "not outcomes already demonstrated by this dataset."
    )

    st.subheader("A proposed women's employment and career-support programme")
    st.write(
        "A six-month pathway could connect education with paid employment through "
        "skills development, mentoring, employer engagement and placement support."
    )

    roadmap = [
        ("01", "Reach and assess", "Month 1", "Recruit participants; map skills, interests and barriers."),
        ("02", "Build skills", "Months 2–3", "Digital, technical and professional training."),
        ("03", "Connect and mentor", "Month 4", "Mentors, CV support, interview practice and employer sessions."),
        ("04", "Access paid work", "Month 5", "Internships, apprenticeships and paid vacancies."),
        ("05", "Track progress", "Month 6+", "Measure placements, retention and participant feedback."),
    ]
    road_cols = st.columns(5, gap="small")
    for col, (num, title, timing, desc) in zip(road_cols, roadmap):
        with col:
            st.markdown(f"<div class='roadmap-step'><div class='num'>{num}</div><h4>{title}</h4><div class='section-kicker'>{timing}</div><p>{desc}</p></div>", unsafe_allow_html=True)
    st.caption("Proposed sequence · the roadmap is a programme design recommendation, not a result measured by the survey.")

    with st.expander("Pillar 1 — Career readiness and skills development", expanded=True):
        st.markdown(
            "- Digital literacy, data skills and workplace technology.\n"
            "- CV writing, interview preparation and job-search support.\n"
            "- Career counselling that connects university majors with labour-market demand.\n"
            "- Short professional training and certification opportunities."
        )
    with st.expander("Pillar 2 — Mentorship and employer connections"):
        st.markdown(
            "- Match students and jobseekers with women professionals and mentors.\n"
            "- Build internship, apprenticeship and entry-level job partnerships.\n"
            "- Organize employer networking events and career days.\n"
            "- Encourage transparent recruitment and promotion criteria."
        )
    with st.expander("Pillar 3 — Removing barriers to participation"):
        st.markdown(
            "- Assess childcare, transport, safety and flexible-work needs.\n"
            "- Offer return-to-work and skills-refresh pathways.\n"
            "- Promote family-friendly and gender-inclusive workplace practices.\n"
            "- Connect participants with relevant support services."
        )

    st.subheader("Illustrative six-month implementation plan")
    programme = pd.DataFrame([
        {"Phase": "1. Outreach and assessment", "Timing": "Month 1", "Activities": "Recruit participants; assess skills, interests and employment barriers."},
        {"Phase": "2. Skills development", "Timing": "Months 2–3", "Activities": "Deliver digital, technical and professional training."},
        {"Phase": "3. Mentoring and career preparation", "Timing": "Month 4", "Activities": "Provide mentoring, CV support, interviews and employer sessions."},
        {"Phase": "4. Placement and transition", "Timing": "Month 5", "Activities": "Connect participants to internships, apprenticeships and paid vacancies."},
        {"Phase": "5. Follow-up and evaluation", "Timing": "Month 6 and follow-up", "Activities": "Track placements, employment status and participant feedback."},
    ])
    st.dataframe(programme, use_container_width=True, hide_index=True)

    st.subheader("Potential partners")
    st.write(
        "A programme could be developed with women's rights and gender-equality "
        "organizations, youth-employment NGOs, university career centres, public "
        "employment institutions, professional associations and private-sector "
        "employers. Potential institutional stakeholders to consult include the "
        "National Commission for Lebanese Women, UN Women and the ILO. Their "
        "current programmes and willingness to partner should be verified before "
        "approaching or naming them as implementing partners."
    )

    st.subheader("How would we measure impact?")
    impact_options = [
        "Number of women enrolled and completing training",
        "Share receiving mentoring or career counselling",
        "Internship and apprenticeship placement rate",
        "Share entering paid employment after the programme",
        "Employment retention after three and six months",
        "Participant-reported confidence and job-readiness",
        "Pay, hours and job quality among placed participants",
    ]
    selected_indicators = st.multiselect(
        "Select indicators for a programme evaluation",
        options=impact_options,
        default=impact_options[:4],
        key="programme_impact_indicators",
    )
    if selected_indicators:
        st.write("**Selected evaluation indicators:**")
        for indicator in selected_indicators:
            st.write(f"- {indicator}")
    st.caption("These are proposed monitoring indicators. They are not measured by the 2022 descriptive analysis.")


# ============================================================
# 15. METHODOLOGY AND SOURCES TAB
# ============================================================

with methods_tab:
    st.header("Methodology and sources")
    st.info("Visual assets are contextual photographs from Pexels, not survey evidence. No AI-generated images are used. The research charts and roadmap are generated from the project data and programme recommendations.")
    st.write(
        "This section documents where the data came from, how the tables were "
        "prepared, what the indicators mean and which questions the available "
        "data cannot answer."
    )

    st.subheader("1. Data sources")
    st.markdown(
        "**2022 national survey:** Lebanon Labour Force and Household Living "
        "Conditions Survey 2022, using the supplied workbook tables EA1, EA3, "
        "EA6, EA8, EA37, EA39 and EA40. These tables provide gender comparisons "
        "for participation, unemployment, education, employment status and job-search methods."
    )
    st.markdown(
        "**Historical unemployment:** the supplied historical Excel workbook "
        f"({HISTORICAL_FILE.name if HISTORICAL_FILE else 'workbook not found'}). "
        "It contains World Bank modelled ILO estimates for total, female and male "
        "unemployment from 1991 to 2023. The 2022 survey estimates are a separate "
        "source and should not be treated as methodologically identical."
    )
    st.markdown(
        "**International reference sources:** [World Bank — total unemployment](https://data.worldbank.org/indicator/SL.UEM.TOTL.ZS?locations=LB), "
        "[female unemployment](https://data.worldbank.org/indicator/SL.UEM.TOTL.FE.ZS?locations=LB), "
        "[male unemployment](https://data.worldbank.org/indicator/SL.UEM.TOTL.MA.ZS?locations=LB), "
        "and [ILOSTAT — ILO labour statistics](https://ilostat.ilo.org/data/). "
        "The historical workbook currently loaded by this app is used for unemployment. "
        "The 2022 survey supplies the gender, age and education participation comparisons. "
        "A separate historical labour-force participation series must be added before "
        "the dashboard can present historical participation trends."
    )

    st.subheader("2. How the data were prepared")
    st.markdown(
        "1. **Workbook validation:** checked that the expected survey worksheets were present.\n"
        "2. **Table extraction:** selected the specified rows and columns from each EA worksheet.\n"
        "3. **Column alignment:** assigned consistent names to categories and gender values; "
        "EA39 and EA40 use the adjacent category, women, men and combined columns (A:D).\n"
        "4. **Data cleaning:** converted indicator values to numeric format, removed blank "
        "or total rows, and excluded records without the required gender values.\n"
        "5. **Reshaping:** converted the women and men columns into a long format for "
        "interactive visualizations.\n"
        "6. **Exploratory analysis:** compared rates and shares across gender, age and "
        "education categories and calculated percentage-point differences."
    )

    st.subheader("3. Indicator definitions")
    with st.expander("Labour-force participation rate"):
        st.write("The percentage of the working-age population that is in the labour force — that is, people who are employed or unemployed.")
        st.latex(r"\text{LFPR} = \frac{\text{Employed} + \text{Unemployed}}{\text{Working-age population}} \times 100")
    with st.expander("Unemployment rate"):
        st.write("The percentage of the labour force that is unemployed. Under the standard international definition, unemployed people are without work, available for work and seeking work, subject to the survey's operational rules.")
        st.latex(r"\text{Unemployment rate} = \frac{\text{Unemployed}}{\text{Employed} + \text{Unemployed}} \times 100")
    with st.expander("Employment and paid employment"):
        st.write("Employment is broader than employee work. The supplied employment-status table distinguishes employers, own-account workers, employees, contributing family workers and trainees/apprentices. Contributing family work may be unpaid, so the overall employment and labour-force indicators should not be interpreted as counts of paid employees only. The survey's exact employment definition should be checked against its metadata.")
    with st.expander("Outside the labour force — and why this is not unemployment"):
        st.write("People outside the labour force are neither employed nor classified as unemployed. They are not included in the labour-force denominator. This group may include people studying, providing unpaid care, retired, unable to work, or not currently seeking/available for work. The dataset does not identify the reason for every person being outside the labour force.")
    with st.expander("Educational attainment"):
        st.write("The highest level of education an individual has successfully completed. It is different from years of schooling, current enrolment or qualifications still in progress.")
    with st.expander("Gender pay gap"):
        st.write("The difference between women's and men's average earnings, often expressed relative to men's average earnings. It cannot be calculated from the current participation and unemployment tables; individual or suitably aggregated earnings data are required.")
    with st.expander("Percentage points versus percent change"):
        st.write("A percentage-point gap is the subtraction of two rates (for example, 66% minus 22% = 44 percentage points). A percent change measures a relative change from a starting value. The dashboard uses percentage points when comparing rates across groups.")

    st.subheader("4. Limitations and interpretation")
    st.markdown(
        "- The 2022 survey is a cross-sectional snapshot; associations do not establish causality.\n"
        "- The supplied cross-tabulations do not provide marital-status or household-composition breakdowns.\n"
        "- The supplied tables do not measure unpaid domestic and care work, so its scale or distribution cannot be quantified here.\n"
        "- No individual earnings or wage data were supplied; therefore, the gender pay gap cannot be estimated.\n"
        "- Education composition among employed people is not the same measure as labour-force participation by education.\n"
        "- Historical modelled estimates and 2022 survey estimates may use different methods, definitions and reference populations; comparisons should be made cautiously.\n"
        "- A historical labour-force participation series is not currently loaded in this version of the dashboard.\n"
        "- Findings describe the categories and years available in the supplied workbooks and should not be generalized beyond their coverage."
    )

    st.subheader("5. Data availability check")
    quality_summary = pd.DataFrame([
        {"Dataset section": "Historical unemployment", "Rows available": len(historical), "Status": "Loaded" if not historical.empty else "No data"},
        {"Dataset section": "Unemployment by age", "Rows available": len(age_unemployment), "Status": "Loaded" if not age_unemployment.empty else "No data — inspect EA39 extraction"},
        {"Dataset section": "Participation by age", "Rows available": len(age_lfp), "Status": "Loaded" if not age_lfp.empty else "No data — inspect EA40 extraction"},
        {"Dataset section": "Participation by education", "Rows available": len(education_lfp), "Status": "Loaded" if not education_lfp.empty else "No data — inspect EA40 extraction"},
        {"Dataset section": "Unemployment by education", "Rows available": len(education_unemployment), "Status": "Loaded" if not education_unemployment.empty else "No data — inspect EA39 extraction"},
    ])
    st.dataframe(quality_summary, use_container_width=True, hide_index=True)


# ============================================================
# 16. FOOTER
# ============================================================

st.markdown("---")
st.caption("The Invisible Workforce · Lebanon · Exploratory and explanatory data storytelling")
