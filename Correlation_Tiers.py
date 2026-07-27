import streamlit as st
import os
import pandas as pd
import numpy as np
from scipy import stats
import altair as alt

st.set_page_config(page_title="Team Stats Dashboard", layout="wide")

# ---------------------------------------------------------------
# Default source file/sheet
# ---------------------------------------------------------------
DEFAULT_FILE_PATH = r"C:\Users\abrams\OneDrive - FC Cincinnati\Documents\FCC CSVS\mls_team_data.xlsx"
DEFAULT_SHEET_NAME = "final_plz"

POINTS_CANDIDATES = ["P", "Points", "points", "PTS"]

# ---------------------------------------------------------------
# Columns that are identifiers/metadata, not performance metrics.
# Everything else numeric is treated as a metric.
# ---------------------------------------------------------------
NON_METRIC_COLS = {
    "team_name",
    "season_name",
    "team_stats.Column1",
    "team_stats.account_id",
    "team_stats.team_name",
    "team_stats.team_id",
    "team_stats.competition_id",
    "team_stats.competition_name",
    "team_stats.season_id",
    "team_stats.season_name",
    "team_stats.team_female",
}

# ---------------------------------------------------------------
# Manual labels for important columns.
# Anything not listed here gets auto-cleaned by pretty_col().
# ---------------------------------------------------------------
COLUMN_LABELS = {
    "team_name": "Team",
    "season_name": "Season",
    "GP": "Games Played",
    "W": "Wins",
    "D": "Draws",
    "L": "Losses",
    "F": "Goals For",
    "A": "Goals Against",
    "GD": "Goal Difference",
    "P": "Points",

    "team_stats.team_season_matches": "Matches",
    "team_stats.team_season_minutes": "Minutes",
    "team_stats.team_season_gd": "Goal Difference",
    "team_stats.team_season_xgd": "Expected Goal Difference",

    "team_stats.team_season_np_shots_pg": "Non-Penalty Shots / Game",
    "team_stats.team_season_op_shots_pg": "Open Play Shots / Game",
    "team_stats.team_season_op_shots_outside_box_pg": "Open Play Shots Outside Box / Game",
    "team_stats.team_season_sp_shots_pg": "Set Piece Shots / Game",
    "team_stats.team_season_np_xg_pg": "Non-Penalty Expected Goals / Game",
    "team_stats.team_season_op_xg_pg": "Open Play Expected Goals / Game",
    "team_stats.team_season_sp_xg_pg": "Set Piece Expected Goals / Game",
    "team_stats.team_season_np_xg_per_shot": "Non-Penalty xG / Shot",
    "team_stats.team_season_np_shot_distance": "Non-Penalty Shot Distance",
    "team_stats.team_season_op_shot_distance": "Open Play Shot Distance",
    "team_stats.team_season_sp_shot_distance": "Set Piece Shot Distance",
    "team_stats.team_season_goals_pg": "Goals / Game",
    "team_stats.team_season_penalty_goals_pg": "Penalty Goals / Game",
    "team_stats.team_season_own_goals_pg": "Own Goals / Game",

    "team_stats.team_season_np_shots_conceded_pg": "Non-Penalty Shots Allowed / Game",
    "team_stats.team_season_op_shots_conceded_pg": "Open Play Shots Allowed / Game",
    "team_stats.team_season_op_shots_conceded_outside_box_pg": "Open Play Shots Allowed Outside Box / Game",
    "team_stats.team_season_sp_shots_conceded_pg": "Set Piece Shots Allowed / Game",
    "team_stats.team_season_np_xg_conceded_pg": "Non-Penalty Expected Goals Allowed / Game",
    "team_stats.team_season_op_xg_conceded_pg": "Open Play Expected Goals Allowed / Game",
    "team_stats.team_season_sp_xg_conceded_pg": "Set Piece Expected Goals Allowed / Game",
    "team_stats.team_season_np_xg_per_shot_conceded": "Non-Penalty xG / Shot Allowed",
    "team_stats.team_season_np_shot_distance_conceded": "Non-Penalty Shot Distance Allowed",
    "team_stats.team_season_op_shot_distance_conceded": "Open Play Shot Distance Allowed",
    "team_stats.team_season_sp_shot_distance_conceded": "Set Piece Shot Distance Allowed",
    "team_stats.team_season_goals_conceded_pg": "Goals Allowed / Game",
    "team_stats.team_season_penalty_goals_conceded_pg": "Penalty Goals Allowed / Game",

    "team_stats.team_season_possessions": "Possessions",
    "team_stats.team_season_possession": "Possession %",
    "team_stats.team_season_directness": "Directness",
    "team_stats.team_season_pace_towards_goal": "Pace Towards Goal",
    "team_stats.team_season_passing_ratio": "Passing %",
    "team_stats.team_season_ball_in_play_time": "Ball in Play Time",
    "team_stats.team_season_gk_pass_distance": "Goalkeeper Pass Distance",
    "team_stats.team_season_gk_long_pass_ratio": "Goalkeeper Long Pass %",
    "team_stats.team_season_box_cross_ratio": "Box Cross %",
    "team_stats.team_season_crosses_into_box_pg": "Crosses into Box / Game",
    "team_stats.team_season_successful_crosses_into_box_pg": "Successful Crosses into Box / Game",
    "team_stats.team_season_successful_box_cross_ratio": "Successful Box Cross %",

    "team_stats.team_season_passes_inside_box_pg": "Passes Inside Box / Game",
    "team_stats.team_season_passes_inside_box_conceded_pg": "Passes Inside Box Allowed / Game",
    "team_stats.team_season_deep_completions_pg": "Deep Completions / Game",
    "team_stats.team_season_deep_completions_conceded_pg": "Deep Completions Allowed / Game",
    "team_stats.team_season_deep_progressions_pg": "Deep Progressions / Game",
    "team_stats.team_season_deep_progressions_conceded_pg": "Deep Progressions Allowed / Game",

    "team_stats.team_season_defensive_distance": "Defensive Distance",
    "team_stats.team_season_ppda": "PPDA",
    "team_stats.team_season_defensive_distance_ppda": "Defensive Distance PPDA",
    "team_stats.team_season_opp_passing_ratio": "Opponent Passing %",
    "team_stats.team_season_opp_final_third_pass_ratio": "Opponent Final Third Passing %",
    "team_stats.team_season_pressures_pg": "Pressures / Game",
    "team_stats.team_season_counterpressures_pg": "Counterpressures / Game",
    "team_stats.team_season_pressure_regains_pg": "Pressure Regains / Game",
    "team_stats.team_season_counterpressure_regains_pg": "Counterpressure Regains / Game",
    "team_stats.team_season_defensive_action_regains_pg": "Defensive Action Regains / Game",
    "team_stats.team_season_fhalf_pressures_pg": "First-Half Pressures / Game",
    "team_stats.team_season_fhalf_counterpressures_pg": "First-Half Counterpressures / Game",
    "team_stats.team_season_fhalf_pressures_ratio": "First-Half Pressure %",
    "team_stats.team_season_fhalf_counterpressures_ratio": "First-Half Counterpressure %",

    "team_stats.team_season_counter_attacking_shots_pg": "Counter-Attacking Shots / Game",
    "team_stats.team_season_high_press_shots_pg": "High Press Shots / Game",
    "team_stats.team_season_shots_in_clear_pg": "Clear Shots / Game",
    "team_stats.team_season_counter_attacking_shots_conceded_pg": "Counter-Attacking Shots Allowed / Game",
    "team_stats.team_season_high_press_shots_conceded_pg": "High Press Shots Allowed / Game",
    "team_stats.team_season_shots_in_clear_conceded_pg": "Clear Shots Allowed / Game",

    "team_stats.team_season_corners_pg": "Corners / Game",
    "team_stats.team_season_corner_xg_pg": "Corner xG / Game",
    "team_stats.team_season_xg_per_corner": "xG / Corner",
    "team_stats.team_season_free_kicks_pg": "Free Kicks / Game",
    "team_stats.team_season_free_kick_xg_pg": "Free Kick xG / Game",
    "team_stats.team_season_xg_per_free_kick": "xG / Free Kick",
    "team_stats.team_season_direct_free_kicks_pg": "Direct Free Kicks / Game",
    "team_stats.team_season_direct_free_kick_xg_pg": "Direct Free Kick xG / Game",
    "team_stats.team_season_xg_per_direct_free_kick": "xG / Direct Free Kick",
    "team_stats.team_season_throw_ins_pg": "Throw-Ins / Game",
    "team_stats.team_season_throw_in_xg_pg": "Throw-In xG / Game",
    "team_stats.team_season_xg_per_throw_in": "xG / Throw-In",
    "team_stats.team_season_sp_pg": "Set Pieces / Game",
    "team_stats.team_season_xg_per_sp": "xG / Set Piece",
    "team_stats.team_season_sp_shot_ratio": "Set Piece Shot %",
    "team_stats.team_season_sp_goals_pg": "Set Piece Goals / Game",
    "team_stats.team_season_sp_goal_ratio": "Set Piece Goal %",

    "team_stats.team_season_corners_conceded_pg": "Corners Allowed / Game",
    "team_stats.team_season_corner_xg_conceded_pg": "Corner xG Allowed / Game",
    "team_stats.team_season_free_kicks_conceded_pg": "Free Kicks Allowed / Game",
    "team_stats.team_season_free_kick_xg_conceded_pg": "Free Kick xG Allowed / Game",
    "team_stats.team_season_throw_ins_conceded_pg": "Throw-Ins Allowed / Game",
    "team_stats.team_season_throw_in_xg_conceded_pg": "Throw-In xG Allowed / Game",
    "team_stats.team_season_sp_pg_conceded": "Set Pieces Allowed / Game",
    "team_stats.team_season_xg_per_sp_conceded": "xG / Set Piece Allowed",
    "team_stats.team_season_sp_shot_ratio_conceded": "Set Piece Shot % Allowed",
    "team_stats.team_season_sp_goals_pg_conceded": "Set Piece Goals Allowed / Game",
    "team_stats.team_season_sp_goal_ratio_conceded": "Set Piece Goal % Allowed",

    "team_stats.team_season_penalties_won_pg": "Penalties Won / Game",
    "team_stats.team_season_penalties_conceded_pg": "Penalties Conceded / Game",
    "team_stats.team_season_completed_dribbles_pg": "Completed Dribbles / Game",
    "team_stats.team_season_failed_dribbles_pg": "Failed Dribbles / Game",
    "team_stats.team_season_total_dribbles_pg": "Total Dribbles / Game",
    "team_stats.team_season_dribble_ratio": "Dribble Success %",
    "team_stats.team_season_completed_dribbles_conceded_pg": "Completed Dribbles Allowed / Game",
    "team_stats.team_season_failed_dribbles_conceded_pg": "Failed Dribbles Against / Game",
    "team_stats.team_season_total_dribbles_conceded_pg": "Total Dribbles Against / Game",
    "team_stats.team_season_opposition_dribble_ratio": "Opponent Dribble Success %",
    "team_stats.team_season_yellow_cards_pg": "Yellow Cards / Game",
    "team_stats.team_season_second_yellow_cards_pg": "Second Yellow Cards / Game",
    "team_stats.team_season_red_cards_pg": "Red Cards / Game",

    "team_stats.team_season_gd_pg": "Goal Difference / Game",
    "team_stats.team_season_np_gd_pg": "Non-Penalty Goal Difference / Game",
    "team_stats.team_season_xgd_pg": "Expected Goal Difference / Game",
    "team_stats.team_season_np_xgd_pg": "Non-Penalty Expected Goal Difference / Game",
}


# ---------------------------------------------------------------
# Pretty name helpers
# ---------------------------------------------------------------
def auto_pretty_col(col: str) -> str:
    cleaned = col
    cleaned = cleaned.replace("team_stats.", "")
    cleaned = cleaned.replace("team_season_", "")

    replacements = {
        "np": "non penalty",
        "op": "open play",
        "sp": "set piece",
        "xg": "expected goals",
        "xgd": "expected goal difference",
        "gd": "goal difference",
        "pg": "per game",
        "ppda": "PPDA",
        "gk": "goalkeeper",
        "opp": "opponent",
        "fhalf": "first half",
    }

    parts = cleaned.split("_")
    pretty_parts = []

    for part in parts:
        pretty_parts.append(replacements.get(part, part))

    pretty = " ".join(pretty_parts)
    pretty = pretty.replace("conceded", "allowed")
    pretty = pretty.replace("per game", "/ Game")
    pretty = pretty.title()

    pretty = pretty.replace("Ppda", "PPDA")
    pretty = pretty.replace("Xg", "xG")
    pretty = pretty.replace("Xgd", "xGD")
    pretty = pretty.replace("Non Penalty", "Non-Penalty")
    pretty = pretty.replace("Open Play", "Open Play")
    pretty = pretty.replace("Set Piece", "Set Piece")
    pretty = pretty.replace("Expected Goals", "Expected Goals")
    pretty = pretty.replace("Expected Goal Difference", "Expected Goal Difference")

    return pretty


def pretty_col(col: str) -> str:
    return COLUMN_LABELS.get(col, auto_pretty_col(col))


def pretty_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: pretty_col(c) for c in df.columns})


def significance_label(p_value):
    if pd.isna(p_value):
        return "Not available"
    elif p_value < 0.001:
        return "Highly significant"
    elif p_value < 0.01:
        return "Very significant"
    elif p_value < 0.05:
        return "Significant"
    else:
        return "Not significant"


def significance_sentence(p_value):
    if pd.isna(p_value):
        return "There is not enough information to judge statistical significance."
    elif p_value < 0.05:
        return "This relationship is statistically significant at p < 0.05."
    else:
        return "This relationship is not statistically significant at p < 0.05."


# ---------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------
@st.cache_data
def load_data(file, sheet_name=None) -> pd.DataFrame:
    if hasattr(file, "name") and str(file.name).lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, sheet_name=sheet_name or 0)
    elif isinstance(file, str) and file.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, sheet_name=sheet_name or 0)
    else:
        df = pd.read_csv(file)

    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_metric_columns(df: pd.DataFrame) -> list:
    metrics = []

    for c in df.columns:
        if c in NON_METRIC_COLS:
            continue

        if c in POINTS_CANDIDATES:
            continue

        coerced = pd.to_numeric(df[c], errors="coerce")

        enough_numeric_values = coerced.notna().sum() >= max(3, int(0.5 * len(df)))
        not_constant = coerced.nunique(dropna=True) > 1

        if enough_numeric_values and not_constant:
            metrics.append(c)

    return metrics


def bucket_label(pct: float) -> str:
    if pct >= 75:
        return "Top 25%"
    elif pct >= 25:
        return "Middle 50%"
    else:
        return "Bottom 25%"


BUCKET_COLORS = {
    "Top 25%": "#1a9850",
    "Middle 50%": "#fdae61",
    "Bottom 25%": "#d73027",
}


def style_bucket(val):
    color = BUCKET_COLORS.get(val, "")
    return f"background-color: {color}; color: white;" if color else ""


# ---------------------------------------------------------------
# App title and file loading
# ---------------------------------------------------------------
st.title("Team Season Stats Dashboard")

df = None
default_exists = os.path.exists(DEFAULT_FILE_PATH)

if default_exists:
    df = load_data(DEFAULT_FILE_PATH, sheet_name=DEFAULT_SHEET_NAME)
    st.caption(f"Loaded `{os.path.basename(DEFAULT_FILE_PATH)}` → sheet `{DEFAULT_SHEET_NAME}`")

with st.expander("Use a different file instead"):
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])

    sheet_override = None
    if uploaded is not None and str(uploaded.name).lower().endswith((".xlsx", ".xls")):
        sheet_override = st.text_input("Sheet name", value=DEFAULT_SHEET_NAME)

    if uploaded is not None:
        df = load_data(uploaded, sheet_name=sheet_override)

if df is None:
    st.info(
        f"Couldn't find `{DEFAULT_FILE_PATH}` on this machine, and no file was uploaded. "
        "Upload a CSV or Excel file above to get started."
    )
    st.stop()

# ---------------------------------------------------------------
# Identify key columns
# ---------------------------------------------------------------
name_col = "team_name" if "team_name" in df.columns else df.columns[0]
season_col = "season_name" if "season_name" in df.columns else None

points_col = next((c for c in POINTS_CANDIDATES if c in df.columns), None)

if points_col is None:
    st.error("Couldn't find a points column. Expected one of: P, Points, points, PTS.")

# Season filter
if season_col and df[season_col].nunique() > 1:
    seasons = ["All seasons"] + sorted(df[season_col].dropna().unique().tolist())
    chosen_season = st.selectbox("Season", seasons)

    if chosen_season != "All seasons":
        df = df[df[season_col] == chosen_season].reset_index(drop=True)

metric_cols = get_metric_columns(df)

st.caption(f"{len(df)} rows · {len(metric_cols)} numeric metrics detected")

tab1, tab2, tab3 = st.tabs(
    ["Percentile Explorer", "Correlation to Points", "Regression Model"]
)

# ---------------------------------------------------------------
# TAB 1: Percentile Explorer
# ---------------------------------------------------------------
with tab1:
    st.subheader("Where each team ranks on a selected metric")

    if not metric_cols:
        st.warning("No numeric metric columns found.")
    else:
        metric = st.selectbox(
            "Choose a metric",
            metric_cols,
            format_func=pretty_col,
            key="metric_select",
        )

        work = df[[name_col] + ([season_col] if season_col else []) + [metric]].copy()
        work[metric] = pd.to_numeric(work[metric], errors="coerce")
        work = work.dropna(subset=[metric])

        if work.empty:
            st.warning("No valid data for this metric.")
        else:
            work["Percentile"] = work[metric].rank(pct=True) * 100
            work["Bucket"] = work["Percentile"].apply(bucket_label)
            work = work.sort_values(metric, ascending=False).reset_index(drop=True)
            work["Rank"] = np.arange(1, len(work) + 1)

            col_order = ["Rank", name_col] + ([season_col] if season_col else []) + [
                metric,
                "Percentile",
                "Bucket",
            ]

            display_df = work[col_order].round({metric: 3, "Percentile": 1})
            display_df = pretty_df_columns(display_df)

            counts = work["Bucket"].value_counts()

            c1, c2, c3 = st.columns(3)
            c1.metric("Top 25%", int(counts.get("Top 25%", 0)))
            c2.metric("Middle 50%", int(counts.get("Middle 50%", 0)))
            c3.metric("Bottom 25%", int(counts.get("Bottom 25%", 0)))

            st.dataframe(
                display_df.style.map(style_bucket, subset=["Bucket"]),
                use_container_width=True,
                hide_index=True,
                height=min(600, 40 + 35 * len(display_df)),
            )

            chart_df = work[[name_col, metric]].copy()
            chart_df = chart_df.rename(
                columns={
                    name_col: pretty_col(name_col),
                    metric: pretty_col(metric),
                }
            )

            st.bar_chart(
                chart_df.set_index(pretty_col(name_col))[pretty_col(metric)]
            )

            st.divider()

            st.subheader("Look up a specific team across every metric")

            team_pick = st.selectbox(
                "Team",
                sorted(df[name_col].dropna().unique().tolist()),
            )

            rows = []

            for m in metric_cols:
                vals = pd.to_numeric(df[m], errors="coerce")

                tmp = df[[name_col]].copy()
                tmp["val"] = vals
                tmp = tmp.dropna(subset=["val"])

                if team_pick not in tmp[name_col].values:
                    continue

                percentile_values = (
                    tmp["val"]
                    .rank(pct=True)
                    .loc[tmp[name_col] == team_pick]
                    .values
                )

                if len(percentile_values) == 0:
                    continue

                pct_val = percentile_values[0] * 100
                team_val = tmp.loc[tmp[name_col] == team_pick, "val"].values[0]

                rows.append(
                    {
                        "Metric": pretty_col(m),
                        "Raw Metric": m,
                        "Value": round(team_val, 3),
                        "Percentile": round(pct_val, 1),
                        "Bucket": bucket_label(pct_val),
                    }
                )

            team_df = pd.DataFrame(rows)

            if team_df.empty:
                st.warning("No metric data found for this team.")
            else:
                team_df = team_df.sort_values("Percentile", ascending=False)
                team_df_display = team_df[["Metric", "Value", "Percentile", "Bucket"]]

                st.dataframe(
                    team_df_display.style.map(style_bucket, subset=["Bucket"]),
                    use_container_width=True,
                    hide_index=True,
                    height=500,
                )

# ---------------------------------------------------------------
# TAB 2: Correlation to Points
# ---------------------------------------------------------------
with tab2:
    if points_col is None:
        st.warning("No points column found in this file.")
    else:
        st.subheader(f"Correlation of each metric to {pretty_col(points_col)}")

        pts = pd.to_numeric(df[points_col], errors="coerce")

        corr_rows = []

        for m in metric_cols:
            vals = pd.to_numeric(df[m], errors="coerce")
            paired = pd.concat([vals, pts], axis=1).dropna()

            if len(paired) < 5:
                continue

            if paired.iloc[:, 0].nunique() <= 1 or paired.iloc[:, 1].nunique() <= 1:
                continue

            r, p = stats.pearsonr(paired.iloc[:, 0], paired.iloc[:, 1])

            corr_rows.append(
                {
                    "Metric": pretty_col(m),
                    "Raw Metric": m,
                    "Correlation (r)": r,
                    "R-squared": r ** 2,
                    "p-value": p,
                    "Significance": significance_label(p),
                    "N": len(paired),
                }
            )

        corr_df = pd.DataFrame(corr_rows)

        if corr_df.empty:
            st.warning("No valid numeric metrics found for correlation.")
        else:
            corr_df["Abs r"] = corr_df["Correlation (r)"].abs()
            corr_df = corr_df.sort_values("Abs r", ascending=False)

            min_n = st.slider(
                "Minimum sample size (N)",
                5,
                int(corr_df["N"].max()),
                5,
                key="corr_min_n",
            )

            corr_df_f = corr_df[corr_df["N"] >= min_n].reset_index(drop=True)

            display_corr = corr_df_f.drop(columns=["Raw Metric", "Abs r"])

            st.dataframe(
                display_corr.round(
                    {
                        "Correlation (r)": 3,
                        "R-squared": 3,
                        "p-value": 4,
                    }
                ),
                use_container_width=True,
                hide_index=True,
                height=500,
            )

            st.subheader("Top 15 strongest correlations either direction")

            top15 = corr_df_f.sort_values("Abs r", ascending=False).head(15)

            st.bar_chart(top15.set_index("Metric")["Correlation (r)"])

            st.divider()

            st.subheader("Scatter Plot with Progression Line")

            scatter_raw_metric = st.selectbox(
                "Metric to plot against points",
                corr_df_f["Raw Metric"].tolist(),
                format_func=pretty_col,
                key="scatter_metric",
            )

            plot_df = df[[name_col, scatter_raw_metric, points_col]].copy()

            plot_df[scatter_raw_metric] = pd.to_numeric(
                plot_df[scatter_raw_metric], errors="coerce"
            )

            plot_df[points_col] = pd.to_numeric(
                plot_df[points_col], errors="coerce"
            )

            plot_df = plot_df.dropna()

            if len(plot_df) < 5:
                st.warning("Not enough data points to create a reliable scatter plot.")
            else:
                x = plot_df[scatter_raw_metric].values.astype(float)
                y = plot_df[points_col].values.astype(float)

                if np.std(x) == 0 or np.std(y) == 0:
                    st.warning(
                        "This metric or points column is constant, so a trendline cannot be calculated."
                    )
                else:
                    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
                    r_squared = r_value ** 2

                    x_label = pretty_col(scatter_raw_metric)
                    y_label = pretty_col(points_col)
                    team_label = pretty_col(name_col)

                    plot_df_display = plot_df.rename(
                        columns={
                            name_col: team_label,
                            scatter_raw_metric: x_label,
                            points_col: y_label,
                        }
                    )

                    line_x = np.linspace(x.min(), x.max(), 100)
                    line_y = intercept + slope * line_x

                    line_df = pd.DataFrame(
                        {
                            x_label: line_x,
                            y_label: line_y,
                        }
                    )

                    c1, c2, c3, c4 = st.columns(4)

                    c1.metric("Correlation r", round(r_value, 3))
                    c2.metric("R-squared", round(r_squared, 3))
                    c3.metric("p-value", round(p_value, 4))
                    c4.metric("Significance", significance_label(p_value))

                    st.caption(
                        f"Trendline equation: {y_label} = "
                        f"{round(slope, 3)} × {x_label} + {round(intercept, 3)}"
                    )

                    scatter = (
                        alt.Chart(plot_df_display)
                        .mark_circle(size=75, opacity=0.75)
                        .encode(
                            x=alt.X(
                                x_label,
                                title=x_label,
                                scale=alt.Scale(zero=False),
                            ),
                            y=alt.Y(
                                y_label,
                                title=y_label,
                                scale=alt.Scale(zero=False),
                            ),
                            tooltip=[
                                alt.Tooltip(team_label, title="Team"),
                                alt.Tooltip(x_label, title=x_label, format=".3f"),
                                alt.Tooltip(y_label, title=y_label, format=".3f"),
                            ],
                        )
                    )

                    trendline = (
                        alt.Chart(line_df)
                        .mark_line(color="red", size=3)
                        .encode(
                            x=alt.X(x_label, title=x_label),
                            y=alt.Y(y_label, title=y_label),
                        )
                    )

                    chart = (
                        (scatter + trendline)
                        .properties(
                            height=550,
                            title=f"{x_label} vs. {y_label}",
                        )
                        .interactive()
                    )

                    st.altair_chart(chart, use_container_width=True)

                    if p_value < 0.05:
                        st.success(significance_sentence(p_value))
                    else:
                        st.warning(significance_sentence(p_value))

# ---------------------------------------------------------------
# TAB 3: Regression Model
# ---------------------------------------------------------------
with tab3:
    st.subheader("Regression Model: Predict Points from Team Metrics")

    if points_col is None:
        st.warning("No points column found in this file.")
    else:
        st.write(
            "Pick a handful of metrics, and this tab builds a basic multiple linear regression model "
            "to estimate which stats help explain points."
        )

        st.info(
            "Tip: Start with 5–8 meaningful metrics. Anything more than that can get messy fast."
        )

        default_metrics = [
            "team_stats.team_season_np_xg_pg",
            "team_stats.team_season_np_xg_conceded_pg",
            "team_stats.team_season_possession",
            "team_stats.team_season_ppda",
            "team_stats.team_season_deep_progressions_pg",
            "team_stats.team_season_deep_progressions_conceded_pg",
            "team_stats.team_season_pressures_pg",
            "team_stats.team_season_pressure_regains_pg",
        ]

        default_metrics = [m for m in default_metrics if m in metric_cols]

        selected_metrics = st.multiselect(
            "Choose predictor metrics",
            metric_cols,
            default=default_metrics,
            format_func=pretty_col,
            key="regression_metrics",
        )

        if len(selected_metrics) == 0:
            st.warning("Select at least one metric to run the regression.")
        else:
            reg_df = df[[name_col, points_col] + selected_metrics].copy()

            reg_df[points_col] = pd.to_numeric(reg_df[points_col], errors="coerce")

            for m in selected_metrics:
                reg_df[m] = pd.to_numeric(reg_df[m], errors="coerce")

            reg_df = reg_df.dropna()

            if len(reg_df) < len(selected_metrics) + 2:
                st.warning(
                    "Not enough complete rows to run this regression. "
                    "Try selecting fewer metrics."
                )
            else:
                y = reg_df[points_col].values.astype(float)
                X = reg_df[selected_metrics].values.astype(float)

                X_mean = X.mean(axis=0)
                X_std = X.std(axis=0)

                valid_cols = X_std > 0

                if valid_cols.sum() == 0:
                    st.warning("All selected metrics are constant. Regression cannot run.")
                else:
                    X = X[:, valid_cols]
                    used_metrics = [
                        m for m, keep in zip(selected_metrics, valid_cols) if keep
                    ]

                    X_mean = X_mean[valid_cols]
                    X_std = X_std[valid_cols]

                    X_scaled = (X - X_mean) / X_std

                    X_design = np.column_stack([np.ones(len(X_scaled)), X_scaled])

                    beta, residuals, rank, s = np.linalg.lstsq(
                        X_design, y, rcond=None
                    )

                    y_pred = X_design @ beta

                    ss_res = np.sum((y - y_pred) ** 2)
                    ss_tot = np.sum((y - y.mean()) ** 2)

                    r_squared = 1 - (ss_res / ss_tot) if ss_tot != 0 else np.nan

                    n = len(y)
                    k = len(used_metrics)

                    if n > k + 1:
                        adj_r_squared = 1 - ((1 - r_squared) * (n - 1) / (n - k - 1))
                    else:
                        adj_r_squared = np.nan

                    rmse = np.sqrt(np.mean((y - y_pred) ** 2))

                    c1, c2, c3, c4, c5 = st.columns(5)
                    c1.metric("Rows Used", n)
                    c2.metric("Predictors", k)
                    c3.metric("R-squared", round(r_squared, 3))
                    c4.metric("Adjusted R-squared", round(adj_r_squared, 3))
                    c5.metric("RMSE", round(rmse, 3))

                    st.caption(
                        "R-squared shows how much variation in points is explained by the chosen metrics. "
                        "Adjusted R-squared penalizes overly large models. RMSE is the average prediction error in points."
                    )

                    coef_rows = []

                    for metric_name, coef in zip(used_metrics, beta[1:]):
                        coef_rows.append(
                            {
                                "Metric": pretty_col(metric_name),
                                "Raw Metric": metric_name,
                                "Standardized Coefficient": coef,
                                "Direction": "Positive" if coef > 0 else "Negative",
                            }
                        )

                    coef_df = pd.DataFrame(coef_rows)
                    coef_df["Abs Coefficient"] = coef_df[
                        "Standardized Coefficient"
                    ].abs()
                    coef_df = coef_df.sort_values(
                        "Abs Coefficient", ascending=False
                    )

                    st.subheader("Most Important Predictors in This Model")

                    st.dataframe(
                        coef_df[
                            ["Metric", "Standardized Coefficient", "Direction"]
                        ].round({"Standardized Coefficient": 3}),
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.bar_chart(
                        coef_df.set_index("Metric")["Standardized Coefficient"]
                    )

                    st.divider()

                    st.subheader("Actual vs. Predicted Points")

                    pred_df = pd.DataFrame(
                        {
                            pretty_col(name_col): reg_df[name_col].values,
                            "Actual Points": y,
                            "Predicted Points": y_pred,
                            "Residual": y - y_pred,
                        }
                    )

                    st.dataframe(
                        pred_df.round(
                            {
                                "Actual Points": 2,
                                "Predicted Points": 2,
                                "Residual": 2,
                            }
                        ),
                        use_container_width=True,
                        hide_index=True,
                        height=400,
                    )

                    if len(pred_df) >= 5:
                        actual = pred_df["Actual Points"].values.astype(float)
                        predicted = pred_df["Predicted Points"].values.astype(float)

                        slope, intercept, r_value, p_value, std_err = stats.linregress(
                            actual, predicted
                        )

                        line_x = np.linspace(actual.min(), actual.max(), 100)
                        line_y = intercept + slope * line_x

                        reg_line_df = pd.DataFrame(
                            {
                                "Actual Points": line_x,
                                "Predicted Points": line_y,
                            }
                        )

                        actual_pred_scatter = (
                            alt.Chart(pred_df)
                            .mark_circle(size=75, opacity=0.75)
                            .encode(
                                x=alt.X(
                                    "Actual Points",
                                    scale=alt.Scale(zero=False),
                                ),
                                y=alt.Y(
                                    "Predicted Points",
                                    scale=alt.Scale(zero=False),
                                ),
                                tooltip=[
                                    pretty_col(name_col),
                                    "Actual Points",
                                    "Predicted Points",
                                    "Residual",
                                ],
                            )
                        )

                        actual_pred_line = (
                            alt.Chart(reg_line_df)
                            .mark_line(color="red", size=3)
                            .encode(
                                x="Actual Points",
                                y="Predicted Points",
                            )
                        )

                        st.altair_chart(
                            (actual_pred_scatter + actual_pred_line)
                            .properties(
                                height=500,
                                title="Actual Points vs. Predicted Points",
                            )
                            .interactive(),
                            use_container_width=True,
                        )

                    st.caption(
                        "Positive coefficients mean higher values of that stat are associated with more points, "
                        "after controlling for the other selected stats. Negative coefficients mean higher values "
                        "are associated with fewer points."
                    )
