"""
WellCompletionDashboard.py

Group 10 Well Completion Report Web Dashboard

Purpose
-------
This Streamlit dashboard replaces the previous Excel workbook deliverable.
It loads Well Completion Report data, filters for agriculture/animal-use wells,
cleans dates and numeric columns, builds summary tables, creates interactive
charts, an interactive map, and provides plain-English analysis for presentation
preparation.

Beginner note
-------------
This is a PYTHON script. Save this file as:
    WellCompletionDashboard.py

Run it from the RStudio Terminal, Windows Command Prompt, or Anaconda Prompt with:
    streamlit run WellCompletionDashboard.py

If using RStudio with reticulate, install/require packages in the R Console with:
    library(reticulate)
    py_require(c("pandas", "numpy", "streamlit", "plotly"))

Then run the dashboard from the RStudio Terminal, not the R Console:
    streamlit run WellCompletionDashboard.py

NEW in this version
-------------------
- Section 10: Interactive Map — plots every well with valid coordinates on a
  Plotly scatter-mapbox. Points are coloured by county and sized by well yield.
  A sidebar control lets you pick which numeric field to use for point size.
  The map uses the free OpenStreetMap tile layer so no API key is needed.

- Section 15: Live News & Policy Updates — calls the Anthropic API (claude-sonnet-4-20250514)
  with the built-in web_search tool to fetch current California news across three topics:
    1. Well legislation and groundwater policy
    2. Water infrastructure and conservation projects
    3. Wildfire and drought risk
  Results are cached for 30 minutes. A Refresh button lets users re-fetch without
  restarting Streamlit. Requires the ANTHROPIC_API_KEY environment variable to be set
  (see setup instructions below).

Setup for Section 15
--------------------
1. Get a free Anthropic API key at https://console.anthropic.com
2. Set the environment variable before launching Streamlit.

   Windows Command Prompt:
       set ANTHROPIC_API_KEY=sk-ant-...
       streamlit run WellCompletionDashboard.py

   Windows PowerShell:
       $env:ANTHROPIC_API_KEY="sk-ant-..."
       streamlit run WellCompletionDashboard.py

   Mac / Linux Terminal:
       export ANTHROPIC_API_KEY=sk-ant-...
       streamlit run WellCompletionDashboard.py

   Alternatively, create a file called .env in the same folder as this script
   and add the line:
       ANTHROPIC_API_KEY=sk-ant-...
   then pip install python-dotenv and uncomment the dotenv lines near the top
   of this script.
"""

from pathlib import Path
import re
import sys
import os
import json
import time
import datetime
from typing import Iterable, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# ---------------------------------------------------------------------------
# Manual .env loader — no third-party package required.
# Create a file called .env in the same folder as this script and add:
#     ANTHROPIC_API_KEY=sk-ant-...
# The block below reads it automatically every time the dashboard starts.
# ---------------------------------------------------------------------------
def _load_dotenv_manual() -> None:
    """Read KEY=VALUE lines from a .env file and push them into os.environ."""
    env_path = Path(__file__).resolve().parent / ".env"
    if not env_path.exists():
        return
    with open(env_path, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if not _line or _line.startswith("#") or "=" not in _line:
                continue
            _key, _, _val = _line.partition("=")
            _key = _key.strip()
            _val = _val.strip().strip('"').strip("'")
            if _key and _key not in os.environ:
                os.environ[_key] = _val

_load_dotenv_manual()


# =============================================================================
# 1. FILE PATH SETTINGS
# =============================================================================
# The CSV is downloaded automatically from the California open data portal
# the first time the dashboard runs, then cached locally as wellcompletionreports.csv.
# No manual download or file path changes needed.

CSV_URL  = "https://data.cnra.ca.gov/dataset/647afc02-8954-426d-aabd-eff418d2652c/resource/8da7b93b-4e69-495d-9caa-335691a1896b/download/wellcompletionreports.csv"
CSV_FILE = Path("wellcompletionreports.csv")   # cached local copy


def download_csv_if_needed(url: str, dest: Path) -> None:
    """
    Download the CSV from the California open data portal if it is not
    already cached locally. Shows a Streamlit progress bar during download.
    The file is large (~500 MB) so this may take a few minutes on first run.
    """
    if dest.exists():
        return   # Already downloaded — skip.

    st.info(
        f"Downloading the Well Completion Reports CSV from the California open data "
        f"portal for the first time. This file is large and may take a few minutes. "
        f"It will be cached locally as `{dest.name}` so subsequent launches are instant."
    )

    try:
        with requests.get(url, stream=True, timeout=300) as response:
            response.raise_for_status()
            total = int(response.headers.get("content-length", 0))
            downloaded = 0
            progress_bar = st.progress(0, text="Downloading…")

            with open(dest, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024 * 1024):  # 1 MB chunks
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total:
                            pct = min(downloaded / total, 1.0)
                            mb_done = downloaded / 1_048_576
                            mb_total = total / 1_048_576
                            progress_bar.progress(
                                pct,
                                text=f"Downloading… {mb_done:.0f} MB / {mb_total:.0f} MB"
                            )

            progress_bar.progress(1.0, text="Download complete!")
            st.success(f"CSV saved as `{dest.name}`. Reloading dashboard…")
            st.rerun()

    except Exception as exc:
        st.error(
            f"Could not download the CSV automatically.\n\n"
            f"Error: {exc}\n\n"
            f"**Manual fix:** Download the file from:\n{url}\n\n"
            f"Save it as `{dest.name}` in the same folder as this script, "
            f"then relaunch Streamlit."
        )
        st.stop()


# Trigger download before anything else loads.
download_csv_if_needed(CSV_URL, CSV_FILE)

ENCODINGS_TO_TRY = ["utf-8-sig", "utf-8", "cp1252", "latin1"]

# Columns from the project that need to be treated as dates.
DATE_COLUMNS = ["DATEWORKENDED", "RECEIVEDDATE", "PERMITDATE"]

# Columns from the project that need to be treated as numbers.
NUMERIC_COLUMNS = [
    "TOTALDRILLDEPTH",
    "TOTALCOMPLETEDDEPTH",
    "STATICWATERLEVEL",
    "WELLYIELD",
    "GROUNDSURFACEELEVATION",
    "TOPOFPERFORATEDINTERVAL",
    "BOTTOMOFPERFORATEDINTERVAL",
    "CASINGDIAMETER",
    "TOTALDRAWDOWN",
    "PUMPTESTLENGTH",
    "DECIMALLATITUDE",
    "DECIMALLONGITUDE",
]

# Required column for the agriculture/animal-use filter.
USE_COLUMN = "PLANNEDUSEFORMERUSE"
COUNTY_COLUMN = "COUNTYNAME"

# Latitude / longitude column names (after numeric cleaning).
LAT_COL = "DECIMALLATITUDE_NUM"
LON_COL = "DECIMALLONGITUDE_NUM"

# California bounding box — used to drop obviously wrong coordinates.
CA_LAT_MIN, CA_LAT_MAX = 32.5, 42.0
CA_LON_MIN, CA_LON_MAX = -124.5, -114.0

# Map point-size options shown in the sidebar.
MAP_SIZE_OPTIONS = {
    "Well Yield": "WELLYIELD_NUM",
    "Total Drill Depth": "TOTALDRILLDEPTH_NUM",
    "Total Completed Depth": "TOTALCOMPLETEDDEPTH_NUM",
    "Static Water Level": "STATICWATERLEVEL_NUM",
    "Uniform (no sizing)": None,
}

# Maximum number of points plotted on the map (keeps the browser responsive).
MAP_MAX_POINTS = 5_000


# =============================================================================
# 2. PAGE SETUP
# =============================================================================
st.set_page_config(
    page_title="Well Completion Dashboard",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# 3. HELPER FUNCTIONS
# =============================================================================
def find_existing_file(path: Path) -> Path:
    """Return the CSV path — by this point it should already exist from the download step."""
    if path.exists():
        return path
    # Fallback: check same folder as the script.
    script_folder_path = Path(__file__).resolve().parent / path.name
    if script_folder_path.exists():
        return script_folder_path
    st.error(
        f"CSV file `{path.name}` not found. "
        "This should have been downloaded automatically on startup. "
        "Try restarting the dashboard."
    )
    st.stop()


def read_csv_with_encoding_fallback(path: Path) -> Tuple[pd.DataFrame, str]:
    """Read a CSV file by trying several common encodings."""
    last_error: Optional[Exception] = None

    for encoding in ENCODINGS_TO_TRY:
        try:
            df = pd.read_csv(path, encoding=encoding, low_memory=False)
            return df, encoding
        except UnicodeDecodeError as error:
            last_error = error
        except Exception as error:
            last_error = error
            break

    st.error("The CSV file could not be read.")
    st.write("Last error:")
    st.code(str(last_error))
    st.stop()


def clean_text_series(series: pd.Series) -> pd.Series:
    """Convert a column to clean text while preserving missing values as blanks."""
    return series.fillna("").astype(str).str.strip()


def make_numeric(series: pd.Series) -> pd.Series:
    """
    Convert a messy text/numeric column to numeric.

    This removes commas and common unit text before conversion. Invalid values
    become NaN instead of crashing the dashboard.
    """
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def add_clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add clean text, date, year, and numeric helper columns."""
    df = df.copy()

    if COUNTY_COLUMN in df.columns:
        df["COUNTYNAME_CLEAN"] = clean_text_series(df[COUNTY_COLUMN]).replace("", "Unknown")
    else:
        df["COUNTYNAME_CLEAN"] = "Unknown"

    if USE_COLUMN in df.columns:
        df["PLANNEDUSEFORMERUSE_CLEAN"] = clean_text_series(df[USE_COLUMN]).replace("", "Unknown")
    else:
        df["PLANNEDUSEFORMERUSE_CLEAN"] = "Unknown"

    for date_col in DATE_COLUMNS:
        if date_col in df.columns:
            parsed_col = f"{date_col}_PARSED"
            year_col = "RECEIVEDYEAR" if date_col == "RECEIVEDDATE" else f"{date_col}_YEAR"
            df[parsed_col] = pd.to_datetime(df[date_col], errors="coerce")
            df[year_col] = df[parsed_col].dt.year.astype("Int64")

    for numeric_col in NUMERIC_COLUMNS:
        if numeric_col in df.columns:
            df[f"{numeric_col}_NUM"] = make_numeric(df[numeric_col])

    if {"TOTALDRILLDEPTH_NUM", "TOTALCOMPLETEDDEPTH_NUM"}.issubset(df.columns):
        df["DEPTH_DIFFERENCE"] = df["TOTALDRILLDEPTH_NUM"] - df["TOTALCOMPLETEDDEPTH_NUM"]

    return df


def filter_agriculture_animal(df: pd.DataFrame) -> pd.DataFrame:
    """Return only rows where planned/former use contains agriculture or animal."""
    if USE_COLUMN not in df.columns:
        st.warning(
            f"The column `{USE_COLUMN}` was not found, so the dashboard cannot apply "
            "the agriculture/animal filter. It will use all rows in the CSV."
        )
        return df.copy()

    mask = (
        df[USE_COLUMN]
        .fillna("")
        .astype(str)
        .str.contains("agriculture|animal", case=False, regex=True, na=False)
    )
    return df.loc[mask].copy()


def numeric_summary_table(df: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    """Create a summary table for numeric helper columns."""
    rows = []
    label_map = {
        "TOTALDRILLDEPTH_NUM": "Total Drill Depth",
        "TOTALCOMPLETEDDEPTH_NUM": "Total Completed Depth",
        "STATICWATERLEVEL_NUM": "Static Water Level",
        "WELLYIELD_NUM": "Well Yield",
        "GROUNDSURFACEELEVATION_NUM": "Ground Surface Elevation",
        "TOPOFPERFORATEDINTERVAL_NUM": "Top of Perforated Interval",
        "BOTTOMOFPERFORATEDINTERVAL_NUM": "Bottom of Perforated Interval",
        "CASINGDIAMETER_NUM": "Casing Diameter",
        "TOTALDRAWDOWN_NUM": "Total Drawdown",
        "PUMPTESTLENGTH_NUM": "Pump Test Length",
    }

    for col in columns:
        if col not in df.columns:
            continue
        valid = df[col].dropna()
        rows.append(
            {
                "Variable": label_map.get(col, col),
                "Valid Count": int(valid.count()),
                "Missing/Invalid": int(df[col].isna().sum()),
                "Mean": round(valid.mean(), 2) if len(valid) else np.nan,
                "Median": round(valid.median(), 2) if len(valid) else np.nan,
                "Minimum": round(valid.min(), 2) if len(valid) else np.nan,
                "Maximum": round(valid.max(), 2) if len(valid) else np.nan,
            }
        )

    return pd.DataFrame(rows)


def safe_top_value(table: pd.DataFrame, name_col: str, count_col: str) -> Tuple[str, int]:
    """Return the top name and count from a table."""
    if table.empty:
        return "No data", 0
    first_row = table.iloc[0]
    return str(first_row[name_col]), int(first_row[count_col])


def prepare_boxplot_data(df: pd.DataFrame, max_counties: int = 10) -> pd.DataFrame:
    """Keep only top counties for a readable box plot."""
    if "COUNTYNAME_CLEAN" not in df.columns or "TOTALCOMPLETEDDEPTH_NUM" not in df.columns:
        return pd.DataFrame()

    top_counties = df["COUNTYNAME_CLEAN"].value_counts().head(max_counties).index
    box_df = df[df["COUNTYNAME_CLEAN"].isin(top_counties)].copy()
    box_df = box_df.dropna(subset=["TOTALCOMPLETEDDEPTH_NUM"])
    return box_df


def cap_outliers_for_display(
    series: pd.Series, lower_q: float = 0.01, upper_q: float = 0.99
) -> pd.Series:
    """Clip a numeric series for chart display only, not for summary calculations."""
    valid = series.dropna()
    if valid.empty:
        return series
    lower = valid.quantile(lower_q)
    upper = valid.quantile(upper_q)
    return series.clip(lower=lower, upper=upper)


def format_number(value) -> str:
    """Format numbers for metric cards."""
    if pd.isna(value):
        return "N/A"
    try:
        return f"{int(value):,}"
    except Exception:
        return str(value)


def prepare_map_data(df: pd.DataFrame, size_col: Optional[str]) -> pd.DataFrame:
    """
    Return a clean subset of the dataframe suitable for map plotting.

    Steps
    -----
    1. Require valid latitude and longitude within California's bounding box.
    2. Drop rows where the chosen size column is NaN (if a size column is used).
    3. Cap the size column at the 99th percentile so one extreme well doesn't
       make every other point invisible.
    4. If more than MAP_MAX_POINTS rows remain, take a random sample so the
       browser stays responsive.
    """
    required_cols = {LAT_COL, LON_COL}
    if not required_cols.issubset(df.columns):
        return pd.DataFrame()

    map_df = df.dropna(subset=[LAT_COL, LON_COL]).copy()

    # Keep only coordinates that fall inside California.
    map_df = map_df[
        map_df[LAT_COL].between(CA_LAT_MIN, CA_LAT_MAX)
        & map_df[LON_COL].between(CA_LON_MIN, CA_LON_MAX)
    ]

    if size_col and size_col in map_df.columns:
        map_df = map_df.dropna(subset=[size_col])
        cap = map_df[size_col].quantile(0.99)
        map_df["_size_display"] = map_df[size_col].clip(upper=cap)
    else:
        map_df["_size_display"] = 6  # Uniform dot size when no size column chosen.

    if len(map_df) > MAP_MAX_POINTS:
        map_df = map_df.sample(MAP_MAX_POINTS, random_state=42)

    return map_df


# =============================================================================
# 4. LOAD AND PREPARE DATA
# =============================================================================
@st.cache_data(show_spinner=True)
def load_and_prepare_data(csv_path_string: str) -> Tuple[pd.DataFrame, pd.DataFrame, str]:
    """Load the CSV and return original data, filtered data, and encoding used."""
    csv_path = find_existing_file(Path(csv_path_string))
    original_df, encoding_used = read_csv_with_encoding_fallback(csv_path)
    original_df = add_clean_columns(original_df)
    filtered_df = filter_agriculture_animal(original_df)
    filtered_df = add_clean_columns(filtered_df)
    return original_df, filtered_df, encoding_used


# =============================================================================
# 5. DASHBOARD HEADER
# =============================================================================
st.title("Agriculture and Animal Well Completion Report Dashboard")
st.markdown(
    """
This dashboard summarizes California Well Completion Report records where the planned or former use contains
**agriculture** or **animal**. It replaces the previous Excel workbook with an interactive web dashboard.

Use the filters in the sidebar to explore counties, years, and well-use categories.
"""
)

with st.sidebar:
    st.header("Dashboard Controls")
    st.caption(
        "Beginner note: edit `CSV_FILE` near the top of the Python script if your file path changes."
    )
    st.write("Current CSV path:")
    st.code(str(CSV_FILE))

original_data, data, encoding = load_and_prepare_data(str(CSV_FILE))


# =============================================================================
# 6. SIDEBAR FILTERS
# =============================================================================
filtered_data = data.copy()

with st.sidebar:
    st.subheader("Filters")

    county_options = sorted(filtered_data["COUNTYNAME_CLEAN"].dropna().unique().tolist())
    selected_counties = st.multiselect(
        "County",
        options=county_options,
        default=county_options,
    )

    if selected_counties:
        filtered_data = filtered_data[filtered_data["COUNTYNAME_CLEAN"].isin(selected_counties)]

    if "RECEIVEDYEAR" in filtered_data.columns and filtered_data["RECEIVEDYEAR"].notna().any():
        year_values = filtered_data["RECEIVEDYEAR"].dropna().astype(int)
        min_year, max_year = int(year_values.min()), int(year_values.max())
        selected_year_range = st.slider(
            "Received Year Range",
            min_value=min_year,
            max_value=max_year,
            value=(min_year, max_year),
        )
        filtered_data = filtered_data[
            filtered_data["RECEIVEDYEAR"].between(
                selected_year_range[0], selected_year_range[1]
            )
            | filtered_data["RECEIVEDYEAR"].isna()
        ]

    use_options = sorted(filtered_data["PLANNEDUSEFORMERUSE_CLEAN"].dropna().unique().tolist())
    if len(use_options) <= 100:
        selected_uses = st.multiselect(
            "Planned/Former Use",
            options=use_options,
            default=use_options,
        )
        if selected_uses:
            filtered_data = filtered_data[
                filtered_data["PLANNEDUSEFORMERUSE_CLEAN"].isin(selected_uses)
            ]
    else:
        st.caption("Planned/former use filter hidden because there are many unique values.")

    st.subheader("Chart Options")
    top_n_counties = st.slider("Number of top counties to show", 5, 25, 10)
    show_trendline = st.checkbox("Show trend line on scatter chart", value=True)
    show_raw_data = st.checkbox("Show filtered raw data preview", value=False)

    # ------------------------------------------------------------------
    # MAP OPTIONS  (new)
    # ------------------------------------------------------------------
    st.subheader("Map Options")
    map_size_label = st.selectbox(
        "Point size represents",
        options=list(MAP_SIZE_OPTIONS.keys()),
        index=0,
    )
    map_size_col = MAP_SIZE_OPTIONS[map_size_label]

    map_color_by_county = st.checkbox("Colour points by county", value=True)


# =============================================================================
# 7. SUMMARY METRIC CARDS
# =============================================================================
st.header("Summary Metrics")

records_by_county = (
    filtered_data.groupby("COUNTYNAME_CLEAN", dropna=False)
    .size()
    .reset_index(name="Record Count")
    .sort_values("Record Count", ascending=False)
)

top_county, top_county_count = safe_top_value(records_by_county, "COUNTYNAME_CLEAN", "Record Count")

if "RECEIVEDYEAR" in filtered_data.columns:
    records_by_year = (
        filtered_data.dropna(subset=["RECEIVEDYEAR"])
        .assign(RECEIVEDYEAR=lambda d: d["RECEIVEDYEAR"].astype(int))
        .groupby("RECEIVEDYEAR")
        .size()
        .reset_index(name="Record Count")
        .sort_values("RECEIVEDYEAR")
    )
    top_year_table = records_by_year.sort_values("Record Count", ascending=False)
    top_year, top_year_count = safe_top_value(top_year_table, "RECEIVEDYEAR", "Record Count")
    missing_received_date = int(filtered_data["RECEIVEDYEAR"].isna().sum())
else:
    records_by_year = pd.DataFrame()
    top_year, top_year_count = "No data", 0
    missing_received_date = 0

metric_1, metric_2, metric_3, metric_4, metric_5 = st.columns(5)
metric_1.metric("Original CSV Rows", format_number(len(original_data)))
metric_2.metric("Filtered Dashboard Rows", format_number(len(filtered_data)))
metric_3.metric("Unique Counties", format_number(filtered_data["COUNTYNAME_CLEAN"].nunique()))
metric_4.metric("Top County", top_county, format_number(top_county_count))
metric_5.metric("Top Received Year", str(top_year), format_number(top_year_count))

st.caption(
    f"CSV encoding used: `{encoding}`. "
    f"Missing/invalid received dates in current filter: {missing_received_date:,}."
)


# =============================================================================
# 8. KEY FINDINGS
# =============================================================================
st.header("Key Findings")

numeric_cols_for_summary = [
    "TOTALDRILLDEPTH_NUM",
    "TOTALCOMPLETEDDEPTH_NUM",
    "STATICWATERLEVEL_NUM",
    "WELLYIELD_NUM",
]
num_summary = numeric_summary_table(filtered_data, numeric_cols_for_summary)

well_yield_median = np.nan
if "WELLYIELD_NUM" in filtered_data.columns:
    well_yield_median = filtered_data["WELLYIELD_NUM"].median()

completed_depth_median = np.nan
if "TOTALCOMPLETEDDEPTH_NUM" in filtered_data.columns:
    completed_depth_median = filtered_data["TOTALCOMPLETEDDEPTH_NUM"].median()

st.markdown(
    f"""
- The current filter contains **{len(filtered_data):,}** agriculture/animal-use well records.
- The county with the most records is **{top_county}** with **{top_county_count:,}** records.
- The received year with the most records is **{top_year}** with **{top_year_count:,}** records.
- The median total completed depth is **{completed_depth_median:,.2f}** when valid depth data is available.
- The median well yield is **{well_yield_median:,.2f}** when valid yield data is available.
"""
)

st.info(
    "Presentation note: use this section to start your future PowerPoint. "
    "It identifies the main county, year, depth, and yield findings from the dashboard."
)


# =============================================================================
# 9. SUMMARY TABLES
# =============================================================================
st.header("Summary Tables")

table_tab_1, table_tab_2, table_tab_3, table_tab_4 = st.tabs(
    ["Records by County", "Records by Received Year", "Records by Use", "Numeric Summary"]
)

with table_tab_1:
    st.subheader("Records by County")
    st.dataframe(records_by_county, use_container_width=True, hide_index=True)
    st.markdown(
        "This table shows where agriculture/animal-use well records are concentrated. "
        "Higher counts suggest counties with more reporting activity or more "
        "agriculture/animal-use wells in this dataset."
    )

with table_tab_2:
    st.subheader("Records by Received Year")
    if records_by_year.empty:
        st.warning("No valid received-year data is available.")
    else:
        st.dataframe(records_by_year, use_container_width=True, hide_index=True)
        st.markdown(
            "This table summarizes records by the year they were received. "
            "It helps identify reporting spikes or periods with fewer records."
        )

with table_tab_3:
    st.subheader("Records by Planned/Former Use")
    records_by_use = (
        filtered_data.groupby("PLANNEDUSEFORMERUSE_CLEAN", dropna=False)
        .size()
        .reset_index(name="Record Count")
        .sort_values("Record Count", ascending=False)
    )
    st.dataframe(records_by_use, use_container_width=True, hide_index=True)
    st.markdown(
        "This table shows the specific planned or former use categories that matched "
        "the agriculture/animal filter. It is useful for checking whether the filtered "
        "data matches the project definition."
    )

with table_tab_4:
    st.subheader("Numeric Summary")
    st.dataframe(num_summary, use_container_width=True, hide_index=True)
    st.markdown(
        "This table checks whether important numeric fields are usable for analysis. "
        "The missing/invalid count is important because large missing values can affect "
        "conclusions from charts and averages."
    )


# =============================================================================
# 10. INTERACTIVE MAP  (new section)
# =============================================================================
st.header("Interactive Well Location Map")

st.markdown(
    """
This map plots each well with a valid California latitude and longitude from the filtered dataset.
- **Point colour** represents the county (toggle in the sidebar).
- **Point size** represents the numeric field chosen in the sidebar (default: Well Yield).
- Hover over any point to see county, use type, depth, and yield details.
- Use the map controls to zoom, pan, and click individual wells.

> **Note:** Up to {:,} points are shown. If the filtered dataset is larger, a random
> sample is displayed so the map stays responsive.
""".format(MAP_MAX_POINTS)
)

map_df = prepare_map_data(filtered_data, map_size_col)

if map_df.empty:
    st.warning(
        "No wells with valid California coordinates are available in the current filter. "
        "Try widening the county or year filters, or check that the CSV contains "
        "`DECIMALLATITUDE` and `DECIMALLONGITUDE` columns."
    )
else:
    # Build hover data dict — include a field only if the column exists.
    hover_fields = {
        "COUNTYNAME_CLEAN": True,
        "PLANNEDUSEFORMERUSE_CLEAN": True,
        "TOTALDRILLDEPTH_NUM": ":.1f",
        "TOTALCOMPLETEDDEPTH_NUM": ":.1f",
        "WELLYIELD_NUM": ":.1f",
        "STATICWATERLEVEL_NUM": ":.1f",
        LAT_COL: ":.4f",
        LON_COL: ":.4f",
    }
    hover_data = {k: v for k, v in hover_fields.items() if k in map_df.columns}

    # Centre the map on the median well location.
    map_center_lat = float(map_df[LAT_COL].median())
    map_center_lon = float(map_df[LON_COL].median())

    color_col = "COUNTYNAME_CLEAN" if map_color_by_county else None

    fig_map = px.scatter_map(
        map_df,
        lat=LAT_COL,
        lon=LON_COL,
        color=color_col,
        size="_size_display",
        size_max=18,
        hover_data=hover_data,
        zoom=5,
        center={"lat": map_center_lat, "lon": map_center_lon},
        title=f"Agriculture/Animal Well Locations — {len(map_df):,} points shown",
        labels={
            "COUNTYNAME_CLEAN": "County",
            "PLANNEDUSEFORMERUSE_CLEAN": "Planned Use",
            "TOTALDRILLDEPTH_NUM": "Total Drill Depth",
            "TOTALCOMPLETEDDEPTH_NUM": "Completed Depth",
            "WELLYIELD_NUM": "Well Yield",
            "STATICWATERLEVEL_NUM": "Static Water Level",
            LAT_COL: "Latitude",
            LON_COL: "Longitude",
            "_size_display": map_size_label,
        },
        map_style="open-street-map",
    )

    fig_map.update_layout(height=650, margin={"r": 0, "t": 40, "l": 0, "b": 0})
    st.plotly_chart(fig_map, use_container_width=True)

    # Summary line below the map.
    mapped_pct = 100 * len(map_df) / max(len(filtered_data), 1)
    st.caption(
        f"{len(map_df):,} of {len(filtered_data):,} filtered wells "
        f"({mapped_pct:.1f}%) have valid California coordinates and are shown on the map. "
        f"Point size represents: **{map_size_label}**."
    )

    st.markdown(
        "The map uses the free OpenStreetMap tile layer — no API key is required. "
        "Geographic clusters on the map often correspond to high-agricultural regions "
        "such as the Central Valley (Fresno, Tulare, Kings counties) and the Sacramento Valley."
    )


# =============================================================================
# 11. DASHBOARD CHARTS
# =============================================================================
st.header("Interactive Charts")

chart_tab_1, chart_tab_2, chart_tab_3, chart_tab_4, chart_tab_5 = st.tabs(
    [
        "County Bar Chart",
        "Received Year Line Chart",
        "Depth vs Yield Scatter",
        "Completed Depth Box Plot",
        "Static Water Level Histogram",
    ]
)

with chart_tab_1:
    st.subheader("Chart 1: Filtered Records by County")
    top_county_chart_data = records_by_county.head(top_n_counties).sort_values("Record Count")
    fig_county = px.bar(
        top_county_chart_data,
        x="Record Count",
        y="COUNTYNAME_CLEAN",
        orientation="h",
        title=f"Top {top_n_counties} Counties by Agriculture/Animal Well Records",
        labels={"COUNTYNAME_CLEAN": "County", "Record Count": "Number of Records"},
    )
    fig_county.update_layout(height=550)
    st.plotly_chart(fig_county, use_container_width=True)
    st.markdown(
        "This bar chart satisfies the column/bar chart requirement. "
        "It shows which counties have the largest number of agriculture/animal-use "
        "well completion records."
    )

with chart_tab_2:
    st.subheader("Chart 2: Filtered Records by Received Year")
    if records_by_year.empty:
        st.warning("No valid received-year data is available for this chart.")
    else:
        fig_year = px.line(
            records_by_year,
            x="RECEIVEDYEAR",
            y="Record Count",
            markers=True,
            title="Agriculture/Animal Well Records by Received Year",
            labels={"RECEIVEDYEAR": "Received Year", "Record Count": "Number of Records"},
        )
        fig_year.update_layout(height=500)
        st.plotly_chart(fig_year, use_container_width=True)
        st.markdown(
            "This line chart satisfies the line/scatter chart requirement. "
            "It shows how record counts changed over time based on the received date."
        )

with chart_tab_3:
    st.subheader("Chart 3: Total Drill Depth Compared with Well Yield")
    required_scatter_cols = {"TOTALDRILLDEPTH_NUM", "WELLYIELD_NUM"}
    if not required_scatter_cols.issubset(filtered_data.columns):
        st.warning("Total drill depth or well yield numeric columns are not available.")
    else:
        scatter_df = filtered_data.dropna(
            subset=["TOTALDRILLDEPTH_NUM", "WELLYIELD_NUM"]
        ).copy()
        if scatter_df.empty:
            st.warning("No valid total drill depth and well yield pairs are available.")
        else:
            x_cap = scatter_df["TOTALDRILLDEPTH_NUM"].quantile(0.99)
            y_cap = scatter_df["WELLYIELD_NUM"].quantile(0.99)
            scatter_display = scatter_df[
                (scatter_df["TOTALDRILLDEPTH_NUM"] <= x_cap)
                & (scatter_df["WELLYIELD_NUM"] <= y_cap)
            ].copy()

            use_color = filtered_data["COUNTYNAME_CLEAN"].nunique() <= 15

            # Build scatter without plotly's built-in trendline (requires statsmodels).
            fig_scatter = px.scatter(
                scatter_display,
                x="TOTALDRILLDEPTH_NUM",
                y="WELLYIELD_NUM",
                color="COUNTYNAME_CLEAN" if use_color else None,
                hover_data=["COUNTYNAME_CLEAN", "PLANNEDUSEFORMERUSE_CLEAN"],
                title="Total Drill Depth vs Well Yield"
                + (" (with OLS Trend Line)" if show_trendline else ""),
                labels={
                    "TOTALDRILLDEPTH_NUM": "Total Drill Depth (ft)",
                    "WELLYIELD_NUM": "Well Yield (gpm)",
                    "COUNTYNAME_CLEAN": "County",
                },
            )

            # Add trend line manually using numpy — no statsmodels needed.
            if show_trendline and len(scatter_display) >= 2:
                x_vals = scatter_display["TOTALDRILLDEPTH_NUM"].values
                y_vals = scatter_display["WELLYIELD_NUM"].values
                x_mean = x_vals.mean()
                y_mean = y_vals.mean()
                slope = np.sum((x_vals - x_mean) * (y_vals - y_mean)) / np.sum(
                    (x_vals - x_mean) ** 2
                )
                intercept = y_mean - slope * x_mean
                x_line = np.array([x_vals.min(), x_vals.max()])
                y_line = slope * x_line + intercept
                fig_scatter.add_trace(
                    go.Scatter(
                        x=x_line,
                        y=y_line,
                        mode="lines",
                        name="Trend Line",
                        line=dict(color="#e63946", width=2.5, dash="solid"),
                    )
                )
            fig_scatter.update_layout(height=550)
            st.plotly_chart(fig_scatter, use_container_width=True)

            # Show the regression equation and R² if trendline is on.
            if show_trendline and len(scatter_display) >= 2:
                x_vals = scatter_display["TOTALDRILLDEPTH_NUM"].values
                y_vals = scatter_display["WELLYIELD_NUM"].values
                # Simple OLS by hand — no extra package needed.
                x_mean, y_mean = x_vals.mean(), y_vals.mean()
                slope = np.sum((x_vals - x_mean) * (y_vals - y_mean)) / np.sum(
                    (x_vals - x_mean) ** 2
                )
                intercept = y_mean - slope * x_mean
                y_pred = slope * x_vals + intercept
                ss_res = np.sum((y_vals - y_pred) ** 2)
                ss_tot = np.sum((y_vals - y_mean) ** 2)
                r_squared = 1 - ss_res / ss_tot if ss_tot != 0 else float("nan")
                direction = "increases" if slope > 0 else "decreases"
                st.info(
                    f"**Trend line equation:** Well Yield = {slope:.4f} × Drill Depth "
                    f"+ {intercept:.2f}  |  **R² = {r_squared:.4f}**\n\n"
                    f"Interpretation: For every additional foot of drill depth, well yield "
                    f"**{direction}** by {abs(slope):.4f} gpm on average. "
                    f"R² of {r_squared:.4f} means drill depth explains "
                    f"{r_squared * 100:.1f}% of the variation in well yield."
                )

            st.markdown(
                "This scatter chart helps check whether deeper wells appear to have higher "
                "or lower yield. Extreme values (top 1%) are excluded from the display for "
                "readability. Toggle the trend line in the sidebar Chart Options."
            )

with chart_tab_4:
    st.subheader("Chart 4: Total Completed Depth Distribution by County")
    box_df = prepare_boxplot_data(filtered_data, max_counties=top_n_counties)
    if box_df.empty:
        st.warning("No valid total completed depth data is available for the box plot.")
    else:
        fig_box = px.box(
            box_df,
            x="COUNTYNAME_CLEAN",
            y="TOTALCOMPLETEDDEPTH_NUM",
            title=f"Total Completed Depth Distribution for Top {top_n_counties} Counties",
            labels={
                "COUNTYNAME_CLEAN": "County",
                "TOTALCOMPLETEDDEPTH_NUM": "Total Completed Depth",
            },
            points=False,
        )
        fig_box.update_layout(height=550, xaxis_tickangle=-45)
        st.plotly_chart(fig_box, use_container_width=True)
        st.markdown(
            "This box-and-whisker chart satisfies the box plot requirement. "
            "It compares the median and spread of completed well depth across the top counties."
        )

with chart_tab_5:
    st.subheader("Optional Chart 5: Distribution of Static Water Level")
    if "STATICWATERLEVEL_NUM" not in filtered_data.columns:
        st.warning("Static water level numeric column is not available.")
    else:
        hist_df = filtered_data.dropna(subset=["STATICWATERLEVEL_NUM"]).copy()
        if hist_df.empty:
            st.warning("No valid static water level data is available.")
        else:
            hist_df["STATICWATERLEVEL_DISPLAY"] = cap_outliers_for_display(
                hist_df["STATICWATERLEVEL_NUM"]
            )
            fig_hist = px.histogram(
                hist_df,
                x="STATICWATERLEVEL_DISPLAY",
                nbins=40,
                title="Distribution of Static Water Level",
                labels={
                    "STATICWATERLEVEL_DISPLAY": "Static Water Level, capped for display"
                },
            )
            fig_hist.update_layout(height=500)
            st.plotly_chart(fig_hist, use_container_width=True)
            st.markdown(
                "This histogram shows the distribution of static water level. "
                "The chart display caps extreme outliers so the main distribution "
                "is easier to see."
            )




# =============================================================================
# 12. COUNTY COMPARISON TOOL  (new section)
# =============================================================================
st.header("County Comparison Tool")

st.markdown(
    "Select two counties to compare them side by side across key well metrics."
)

county_list = sorted(filtered_data["COUNTYNAME_CLEAN"].dropna().unique().tolist())

if len(county_list) < 2:
    st.warning("At least two counties are needed for comparison. Widen your county filter.")
else:
    comp_col1, comp_col2 = st.columns(2)
    with comp_col1:
        county_a = st.selectbox("County A", options=county_list, index=0, key="county_a")
    with comp_col2:
        default_b = 1 if len(county_list) > 1 else 0
        county_b = st.selectbox("County B", options=county_list, index=default_b, key="county_b")

    if county_a == county_b:
        st.warning("Please select two different counties.")
    else:
        df_a = filtered_data[filtered_data["COUNTYNAME_CLEAN"] == county_a]
        df_b = filtered_data[filtered_data["COUNTYNAME_CLEAN"] == county_b]

        # ── Metric cards ────────────────────────────────────────────────────
        COMPARE_METRICS = [
            ("Total Records",          None,                       "count"),
            ("Median Drill Depth",     "TOTALDRILLDEPTH_NUM",      "median"),
            ("Median Completed Depth", "TOTALCOMPLETEDDEPTH_NUM",  "median"),
            ("Median Well Yield",      "WELLYIELD_NUM",            "median"),
            ("Median Static Water",    "STATICWATERLEVEL_NUM",     "median"),
        ]

        def get_metric(df, col, agg):
            if agg == "count":
                return len(df)
            if col not in df.columns:
                return np.nan
            valid = df[col].dropna()
            if valid.empty:
                return np.nan
            return round(valid.median() if agg == "median" else valid.mean(), 2)

        st.subheader("Key Metric Comparison")
        header_cols = st.columns([3, 2, 2])
        header_cols[0].markdown("**Metric**")
        header_cols[1].markdown(f"**{county_a}**")
        header_cols[2].markdown(f"**{county_b}**")

        for label, col, agg in COMPARE_METRICS:
            val_a = get_metric(df_a, col, agg)
            val_b = get_metric(df_b, col, agg)
            row = st.columns([3, 2, 2])
            row[0].write(label)

            # Highlight the higher value in green.
            def fmt(v):
                if pd.isna(v):
                    return "N/A"
                return f"{v:,.2f}" if isinstance(v, float) else f"{v:,}"

            if not pd.isna(val_a) and not pd.isna(val_b) and val_a != val_b:
                winner_a = val_a > val_b
                row[1].markdown(
                    f"{'🟢 ' if winner_a else ''}{fmt(val_a)}"
                )
                row[2].markdown(
                    f"{'🟢 ' if not winner_a else ''}{fmt(val_b)}"
                )
            else:
                row[1].write(fmt(val_a))
                row[2].write(fmt(val_b))

        st.caption("🟢 indicates the higher value for each metric.")

        # ── Side-by-side bar charts ─────────────────────────────────────────
        st.subheader("Depth & Yield Distribution")

        CHART_COMPARE_COLS = [
            ("TOTALDRILLDEPTH_NUM",     "Total Drill Depth (ft)"),
            ("TOTALCOMPLETEDDEPTH_NUM", "Completed Depth (ft)"),
            ("WELLYIELD_NUM",           "Well Yield (gpm)"),
            ("STATICWATERLEVEL_NUM",    "Static Water Level (ft)"),
        ]

        available_cols = [
            (c, lbl) for c, lbl in CHART_COMPARE_COLS
            if c in filtered_data.columns
        ]

        if available_cols:
            # Build a long-form dataframe for grouped bar chart.
            bar_rows = []
            for col, label in available_cols:
                for county, df_c in [(county_a, df_a), (county_b, df_b)]:
                    valid = df_c[col].dropna()
                    bar_rows.append({
                        "Metric": label,
                        "County": county,
                        "Median Value": round(valid.median(), 2) if not valid.empty else 0,
                    })

            bar_compare_df = pd.DataFrame(bar_rows)

            fig_compare = px.bar(
                bar_compare_df,
                x="Metric",
                y="Median Value",
                color="County",
                barmode="group",
                title=f"Median Values — {county_a} vs {county_b}",
                color_discrete_sequence=["#1f77b4", "#ff7f0e"],
                labels={"Median Value": "Median"},
            )
            fig_compare.update_layout(height=450, xaxis_tickangle=-20)
            st.plotly_chart(fig_compare, use_container_width=True)

        # ── Records by year comparison ──────────────────────────────────────
        if "RECEIVEDYEAR" in filtered_data.columns:
            st.subheader("Records by Year")

            def year_series(df, county):
                return (
                    df.dropna(subset=["RECEIVEDYEAR"])
                    .assign(RECEIVEDYEAR=lambda d: d["RECEIVEDYEAR"].astype(int))
                    .groupby("RECEIVEDYEAR")
                    .size()
                    .reset_index(name="Records")
                    .assign(County=county)
                )

            year_df = pd.concat([year_series(df_a, county_a), year_series(df_b, county_b)])

            fig_year_comp = px.line(
                year_df,
                x="RECEIVEDYEAR",
                y="Records",
                color="County",
                markers=True,
                title=f"Records by Year — {county_a} vs {county_b}",
                color_discrete_sequence=["#1f77b4", "#ff7f0e"],
                labels={"RECEIVEDYEAR": "Year"},
            )
            fig_year_comp.update_layout(height=400)
            st.plotly_chart(fig_year_comp, use_container_width=True)

        st.markdown(
            "Use this tool to directly compare two counties across well depth, yield, "
            "static water level, and reporting history. Useful for identifying regional "
            "differences in agricultural well characteristics."
        )


# =============================================================================
# 13. DATA QUALITY CHECKS
# =============================================================================
st.header("Data Quality Checks")

quality_rows = []
for date_col in DATE_COLUMNS:
    parsed_col = f"{date_col}_PARSED"
    if parsed_col in filtered_data.columns:
        quality_rows.append(
            {
                "Field": date_col,
                "Type": "Date",
                "Valid Count": int(filtered_data[parsed_col].notna().sum()),
                "Missing/Invalid": int(filtered_data[parsed_col].isna().sum()),
            }
        )

for numeric_col in NUMERIC_COLUMNS:
    clean_col = f"{numeric_col}_NUM"
    if clean_col in filtered_data.columns:
        quality_rows.append(
            {
                "Field": numeric_col,
                "Type": "Numeric",
                "Valid Count": int(filtered_data[clean_col].notna().sum()),
                "Missing/Invalid": int(filtered_data[clean_col].isna().sum()),
            }
        )

quality_table = pd.DataFrame(quality_rows)
st.dataframe(quality_table, use_container_width=True, hide_index=True)
st.markdown(
    "This table documents which date and numeric fields are usable. "
    "It is important for explaining why some charts may have fewer records than "
    "the full filtered dataset."
)


# =============================================================================
# 13. RAW DATA PREVIEW
# =============================================================================
if show_raw_data:
    st.header("Filtered Raw Data Preview")
    st.dataframe(filtered_data.head(1000), use_container_width=True)
    st.caption("Showing the first 1,000 filtered rows only so the dashboard stays responsive.")


# =============================================================================
# 14. DOWNLOAD OPTIONS
# =============================================================================
st.header("Download Prepared Data")

csv_download = filtered_data.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="Download current filtered dashboard data as CSV",
    data=csv_download,
    file_name="Prepared_Agriculture_Animal_Well_Data.csv",
    mime="text/csv",
)

# =============================================================================
# 15. LIVE NEWS & POLICY UPDATES
# =============================================================================
st.header("Live News & Policy Updates")

st.markdown(
    "Click the button below to fetch a short AI-generated summary of recent California "
    "news covering well legislation, water projects, and wildfire/drought risk. "
    "Results are cached for the session so you only need to fetch once."
)

# --- Gemini API configuration -----------------------------------------------
GEMINI_MODEL    = "gemini-2.0-flash-lite"   # Lighter model = higher free rate limit
NEWS_CACHE_TTL  = 60 * 60                   # Cache for 1 hour
NEWS_MAX_TOKENS = 256                        # Very short response

COMBINED_PROMPT = (
    "You are a California water and agriculture expert. "
    "Give me exactly 3 bullet points total — one for each topic below. "
    "Each bullet must be a single sentence. Bold the topic label at the start.\n\n"
    "• **Well Legislation**: One key 2024-2025 California groundwater policy update.\n"
    "• **Water Projects**: One key 2024-2025 California water infrastructure development.\n"
    "• **Wildfire/Drought**: One key 2024-2025 California fire or drought update.\n\n"
    "Be factual. No invented URLs needed."
)


def fetch_news_from_gemini(api_key: str) -> str:
    """Call the Gemini API with a single minimal combined prompt."""
    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={api_key}"
    )
    payload = {
        "contents": [{"parts": [{"text": COMBINED_PROMPT}]}],
        "generationConfig": {
            "maxOutputTokens": NEWS_MAX_TOKENS,
            "temperature": 0.3,
        },
    }
    try:
        response = requests.post(url, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()
        candidates = data.get("candidates", [])
        if not candidates:
            return "⚠️ No response returned. Try again in a moment."
        parts = candidates[0].get("content", {}).get("parts", [])
        text_parts = [p["text"] for p in parts if "text" in p and p["text"].strip()]
        return "\n\n".join(text_parts) if text_parts else "⚠️ Empty response."
    except requests.exceptions.Timeout:
        return "⚠️ Request timed out. Try again."
    except requests.exceptions.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        if status == 429:
            return "⚠️ Rate limit hit. Wait 60 seconds and try again."
        if status in (400, 403):
            return f"⚠️ API key error ({status}). Check GEMINI_API_KEY in your .env file."
        return f"⚠️ HTTP error {status}: {exc}"
    except Exception as exc:
        return f"⚠️ Unexpected error: {exc}"


# --- API key check & render --------------------------------------------------
gemini_api_key = os.environ.get("GEMINI_API_KEY", "").strip()

if not gemini_api_key:
    st.warning(
        "**GEMINI_API_KEY not found.** Add `GEMINI_API_KEY=AIza...` to your `.env` file "
        "and relaunch Streamlit."
    )
else:
    # Only fetch when the user explicitly clicks — never on page load.
    cache     = st.session_state.get("news_cache_combined", {})
    is_fresh  = "text" in cache and (time.time() - cache.get("fetched_at", 0)) < NEWS_CACHE_TTL

    if is_fresh:
        st.markdown(cache["text"])
        age = int(time.time() - cache["fetched_at"])
        st.caption(f"Fetched {age // 60}m {age % 60}s ago — cached for 1 hour.")
        if st.button("🔄 Refresh"):
            st.session_state["news_cache_combined"] = {}
            st.rerun()
    else:
        st.info("Press the button below to load the latest California water & wildfire news summary.")
        if st.button("📰 Load News Summary"):
            with st.spinner("Asking Gemini for a summary…"):
                result = fetch_news_from_gemini(gemini_api_key)
            st.session_state["news_cache_combined"] = {
                "text": result,
                "fetched_at": time.time(),
            }
            st.rerun()

    st.caption(
        "Summary generated by Google Gemini. Verify details at "
        "water.ca.gov · fire.ca.gov · droughtmonitor.unl.edu"
    )


# =============================================================================
# 16. RAW DATA PREVIEW
# =============================================================================
if show_raw_data:
    st.header("Filtered Raw Data Preview")
    st.dataframe(filtered_data.head(1000), use_container_width=True)
    st.caption("Showing the first 1,000 filtered rows only so the dashboard stays responsive.")


st.caption("End of dashboard.")
