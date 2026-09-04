
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Indian Army Defence Conversation Intelligence",
    page_icon="🎖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========================= DESIGN =========================
st.markdown("""
<style>
* {font-family: "Times New Roman", Times, serif !important;}
html, body, [class*="css"] {color: #f8f4e8 !important;}
.stApp {
    background:
      radial-gradient(circle at 8% 5%, rgba(88,115,63,.28), transparent 28%),
      radial-gradient(circle at 92% 4%, rgba(191,151,55,.16), transparent 24%),
      linear-gradient(135deg,#020302 0%,#0a0e0a 48%,#030403 100%);
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#070a07 0%,#11180f 100%);
    border-right: 1px solid rgba(220,185,95,.42);
}
[data-testid="stSidebar"] * {color:#f7f2e5 !important;}

.header-wrap {
    padding: 18px 26px;
    border-radius: 20px;
    border: 1px solid rgba(220,185,95,.62);
    background: linear-gradient(90deg,rgba(5,8,5,.98),rgba(38,55,31,.96),rgba(7,9,6,.98));
    box-shadow: 0 16px 38px rgba(0,0,0,.48);
    margin-bottom: 18px;
}
.header-title {
    color:#e6c66f !important;
    font-size:40px;
    font-weight:700;
    letter-spacing:.8px;
    margin-bottom:6px;
}
.header-sub {
    color:#f3eee1 !important;
    font-size:17px;
    line-height:1.45;
}
.filter-panel {
    background:rgba(12,17,12,.94);
    border:1px solid rgba(220,185,95,.42);
    border-radius:16px;
    padding:15px 18px 8px 18px;
    margin-bottom:18px;
}
.section-title {
    color:#e6c66f !important;
    font-size:28px;
    font-weight:700;
    border-left:6px solid #d9b45d;
    padding-left:12px;
    margin:22px 0 12px 0;
}
.kpi-card {
    background:linear-gradient(145deg,rgba(32,41,30,.98),rgba(8,11,8,.98));
    border:1px solid rgba(220,185,95,.38);
    border-radius:16px;
    padding:16px;
    min-height:104px;
    text-align:center;
    box-shadow:0 8px 22px rgba(0,0,0,.25);
}
.kpi-label {color:#d7dbd1 !important;font-size:15px;font-weight:600;}
.kpi-value {color:#f0cf78 !important;font-size:29px;font-weight:700;margin-top:7px;}
.driver-card {
    background:rgba(17,23,17,.97);
    border:1px solid rgba(255,255,255,.13);
    border-left:7px solid #d9b45d;
    border-radius:12px;
    padding:17px 19px;
    margin-bottom:12px;
    color:#ffffff !important;
    box-shadow:0 7px 20px rgba(0,0,0,.22);
}
.driver-card * {color:#ffffff !important;}
.driver-state {font-size:21px !important;font-weight:700 !important;}
.note-box {
    background:rgba(65,52,18,.28);
    border:1px solid rgba(224,190,90,.40);
    border-radius:12px;
    padding:12px 16px;
    color:#f6efd9 !important;
}
.stButton button, .stDownloadButton button {
    background:#4b5f38 !important;
    color:#fff !important;
    border:1px solid #d9b45d !important;
}
div[data-testid="stMetric"] {background:rgba(13,17,13,.85);border-radius:10px;padding:8px;}
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] {color:#fff !important;}
label, p, span, li {color:#f5f1e6 !important;}
</style>
""", unsafe_allow_html=True)

STATE_COORDINATES = {
"Andhra Pradesh":(15.9129,79.7400),"Arunachal Pradesh":(28.2180,94.7278),"Assam":(26.2006,92.9376),
"Bihar":(25.0961,85.3131),"Chhattisgarh":(21.2787,81.8661),"Goa":(15.2993,74.1240),
"Gujarat":(22.2587,71.1924),"Haryana":(29.0588,76.0856),"Himachal Pradesh":(31.1048,77.1734),
"Jharkhand":(23.6102,85.2799),"Karnataka":(15.3173,75.7139),"Kerala":(10.8505,76.2711),
"Madhya Pradesh":(22.9734,78.6569),"Maharashtra":(19.7515,75.7139),"Manipur":(24.6637,93.9063),
"Meghalaya":(25.4670,91.3662),"Mizoram":(23.1645,92.9376),"Nagaland":(26.1584,94.5624),
"Odisha":(20.9517,85.0985),"Punjab":(31.1471,75.3412),"Rajasthan":(27.0238,74.2179),
"Sikkim":(27.5330,88.5122),"Tamil Nadu":(11.1271,78.6569),"Telangana":(18.1124,79.0193),
"Tripura":(23.9408,91.9882),"Uttar Pradesh":(26.8467,80.9462),"Uttarakhand":(30.0668,79.0193),
"West Bengal":(22.9868,87.8550),"Delhi":(28.7041,77.1025),"Jammu and Kashmir":(33.7782,76.5762),
"Ladakh":(34.1526,77.5771),"Puducherry":(11.9416,79.8083),"Andaman and Nicobar Islands":(11.7401,92.6586),
"Chandigarh":(30.7333,76.7794),"Dadra and Nagar Haveli and Daman and Diu":(20.3974,72.8328),
"Lakshadweep":(10.5667,72.6417)
}

COLORS = {
    "HOTSPOT":"#ff2b2b",
    "HIGH ACTIVITY":"#ff8c00",
    "ELEVATED":"#ffd21f",
    "NORMAL":"#32c76a"
}

@st.cache_data
def load_data():
    xlsx = Path("Indian_Army_Defence_Conversation_July_August_2026.xlsx")
    if xlsx.exists():
        return pd.read_excel(xlsx, sheet_name="Conversation_Data")
    csv = Path("Indian_Army_Defence_Conversation_July_August_2026.csv")
    if csv.exists():
        return pd.read_csv(csv)
    return None

df = load_data()
if df is None:
    st.error("Dataset not found. Keep the Excel/CSV file in the same folder as app.py.")
    st.stop()

df["Date"] = pd.to_datetime(df["Date"])
df["Conversation_Count"] = pd.to_numeric(df["Conversation_Count"], errors="coerce").fillna(0)
df["State"] = df["State"].astype(str).str.strip()

# ========================= HEADER WITH USER LOGO =========================
logo = Path("indian_army_logo.png")
left, right = st.columns([1, 7])
with left:
    if logo.exists():
        st.image(str(logo), width=105)
    else:
        st.markdown("## ")
with right:
    st.markdown("""
    <div class="header-wrap">
        <div class="header-title">INDIAN ARMY DEFENCE CONVERSATION INTELLIGENCE</div>
        <div class="header-sub">
        Geographic Heat Mapping of Defence Conversation • Date-wise Classification •
        Hotspot Detection • Trending News/Content Driver Analysis
        </div>
    </div>
    """, unsafe_allow_html=True)

# st.markdown("""
# <div class="note-box">
# <b>Dataset note:</b> This application uses an academic project dataset for July–August 2026.
# The records are synthetic/modelled for demonstration and analysis and are not official platform analytics.
# </div>
# """, unsafe_allow_html=True)

# ========================= FILTERS =========================
st.markdown('<div class="section-title">📅 Date Classification & Analysis Filters</div>', unsafe_allow_html=True)
st.markdown('<div class="filter-panel">', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns([1.25, 1.35, 1.5, 1.4])

min_date = df["Date"].min().date()
max_date = df["Date"].max().date()

with f1:
    analysis_level = st.selectbox(
        "Classify / Analyse by",
        ["Daily", "Weekly", "Monthly", "Custom Date Range"]
    )

with f2:
    if analysis_level == "Daily":
        selected_date = st.date_input("Select particular date", value=max_date,
                                      min_value=min_date, max_value=max_date)
        start_date = end_date = selected_date
        date_label = pd.Timestamp(selected_date).strftime("%d %B %Y")
    elif analysis_level == "Weekly":
        weeks = sorted(df["Week_Number"].unique())
        week = st.selectbox("Select ISO week", weeks, index=len(weeks)-1)
        temp = df[df["Week_Number"] == week]
        start_date = temp["Date"].min().date()
        end_date = temp["Date"].max().date()
        date_label = f"Week {week}: {start_date.strftime('%d %b')} – {end_date.strftime('%d %b %Y')}"
    elif analysis_level == "Monthly":
        months = ["July 2026", "August 2026"]
        month = st.selectbox("Select month", months, index=1)
        temp = df[df["Analysis_Period"] == month]
        start_date = temp["Date"].min().date()
        end_date = temp["Date"].max().date()
        date_label = month
    else:
        selected_range = st.date_input(
            "Select date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
        if isinstance(selected_range, tuple) and len(selected_range) == 2:
            start_date, end_date = selected_range
        else:
            start_date = end_date = selected_range
        date_label = f"{pd.Timestamp(start_date).strftime('%d %b %Y')} – {pd.Timestamp(end_date).strftime('%d %b %Y')}"

with f3:
    all_platforms = sorted(df["Platform"].unique())
    selected_platforms = st.multiselect("Social media platform(s)", all_platforms, default=all_platforms)

with f4:
    all_states = sorted(df["State"].unique())
    selected_states = st.multiselect("State & Unions", all_states, default=all_states)

st.markdown('</div>', unsafe_allow_html=True)

if not selected_states:
    st.warning("Please select at least one state.")
    st.stop()
if not selected_platforms:
    st.warning("Please select at least one social media platform.")
    st.stop()

filtered = df[
    (df["Date"].dt.date >= start_date) &
    (df["Date"].dt.date <= end_date) &
    (df["State"].isin(selected_states)) &
    (df["Platform"].isin(selected_platforms))
].copy()

if filtered.empty:
    st.warning("No data available for the selected filters.")
    st.stop()

# ========================= BASELINE =========================
period_days = max(1, (pd.Timestamp(end_date) - pd.Timestamp(start_date)).days + 1)
baseline_end = pd.Timestamp(start_date) - pd.Timedelta(days=1)
baseline_start = baseline_end - pd.Timedelta(days=period_days - 1)

baseline = df[
    (df["Date"] >= baseline_start) &
    (df["Date"] <= baseline_end) &
    (df["State"].isin(selected_states)) &
    (df["Platform"].isin(selected_platforms))
].copy()

summary = filtered.groupby("State", as_index=False).agg(
    Current_Activity=("Conversation_Count", "sum"),
    Current_Engagement=("Likes_Modelled", "sum")
)

if baseline.empty:
    baseline_summary = pd.DataFrame({"State": selected_states, "Baseline_Activity": 0})
else:
    baseline_summary = baseline.groupby("State", as_index=False).agg(
        Baseline_Activity=("Conversation_Count", "sum")
    )

summary = summary.merge(baseline_summary, on="State", how="left").fillna({"Baseline_Activity":0})
summary["Rise"] = summary["Current_Activity"] - summary["Baseline_Activity"]
summary["Rise_Percent"] = np.where(
    summary["Baseline_Activity"] > 0,
    summary["Rise"] / summary["Baseline_Activity"] * 100,
    np.where(summary["Current_Activity"] > 0, 100, 0)
)

# Hybrid hotspot score: current intensity + change from baseline
mean_activity = max(summary["Current_Activity"].mean(), 1)
summary["Activity_Intensity"] = summary["Current_Activity"] / mean_activity
summary["Hotspot_Score"] = (summary["Activity_Intensity"] * 0.55) + (
    np.clip(summary["Rise_Percent"], -100, 150) / 100 * 0.45
)

def classify(row):
    if row["Rise_Percent"] >= 45 or row["Activity_Intensity"] >= 1.85:
        return "HOTSPOT"
    if row["Rise_Percent"] >= 22 or row["Activity_Intensity"] >= 1.40:
        return "HIGH ACTIVITY"
    if row["Rise_Percent"] >= 7 or row["Activity_Intensity"] >= 1.08:
        return "ELEVATED"
    return "NORMAL"

summary["Classification"] = summary.apply(classify, axis=1)
summary["Color"] = summary["Classification"].map(COLORS)
summary["Latitude"] = summary["State"].map(lambda s: STATE_COORDINATES.get(s, (np.nan,np.nan))[0])
summary["Longitude"] = summary["State"].map(lambda s: STATE_COORDINATES.get(s, (np.nan,np.nan))[1])

# ========================= NEWS / CONTENT DRIVER =========================
driver_rows = []
for state in summary["State"]:
    cur = filtered[filtered["State"] == state].groupby(
        ["Topic", "News_or_Content_Driver"], as_index=False
    )["Conversation_Count"].sum()

    old = baseline[baseline["State"] == state].groupby(
        ["Topic", "News_or_Content_Driver"], as_index=False
    )["Conversation_Count"].sum() if not baseline.empty else pd.DataFrame(
        columns=["Topic","News_or_Content_Driver","Conversation_Count"]
    )

    cur = cur.rename(columns={"Conversation_Count":"Current_Topic_Activity"})
    old = old.rename(columns={"Conversation_Count":"Baseline_Topic_Activity"})

    combined = cur.merge(old, on=["Topic","News_or_Content_Driver"], how="outer").fillna(0)
    combined["Topic_Increase"] = combined["Current_Topic_Activity"] - combined["Baseline_Topic_Activity"]

    if combined.empty:
        driver_rows.append([state,"No topic data","No content driver",0,0,0])
    else:
        best = combined.sort_values(
            ["Topic_Increase","Current_Topic_Activity"],
            ascending=False
        ).iloc[0]
        driver_rows.append([
            state,
            best["Topic"],
            best["News_or_Content_Driver"],
            int(best["Current_Topic_Activity"]),
            int(best["Baseline_Topic_Activity"]),
            int(best["Topic_Increase"])
        ])

drivers = pd.DataFrame(driver_rows, columns=[
    "State","Trending_Topic","News_or_Content_Driver",
    "Current_Topic_Activity","Baseline_Topic_Activity","Topic_Increase"
])

summary = summary.merge(drivers, on="State", how="left")

# ========================= KPI =========================
st.markdown(f"### 📍 Current classification: **{date_label}**")
k1,k2,k3,k4,k5 = st.columns(5)

hotspots = int((summary["Classification"] == "HOTSPOT").sum())
high = int((summary["Classification"] == "HIGH ACTIVITY").sum())
elevated = int((summary["Classification"] == "ELEVATED").sum())

kpi_values = [
    ("TOTAL CONVERSATION ACTIVITY", f"{int(summary['Current_Activity'].sum()):,}"),
    ("HOTSPOTS", str(hotspots)),
    ("HIGH ACTIVITY", str(high)),
    ("ELEVATED", str(elevated)),
    ("STATES & UNIONS str(len(summary)))
]

for col, (label, value) in zip([k1,k2,k3,k4,k5], kpi_values):
    col.markdown(
        f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div></div>',
        unsafe_allow_html=True
    )

st.caption(
    f"Comparison baseline: previous {period_days} day(s), from "
    f"{baseline_start.strftime('%d %b %Y')} to {baseline_end.strftime('%d %b %Y')}."
)

# ========================= BLACK HEAT MAP =========================
st.markdown('<div class="section-title">🗺️ India Geographic Heat Map</div>', unsafe_allow_html=True)
st.markdown(
    "🔴 **HOTSPOT** &nbsp;&nbsp; 🟠 **HIGH ACTIVITY** &nbsp;&nbsp; "
    "🟡 **ELEVATED** &nbsp;&nbsp; 🟢 **NORMAL**"
)

mapdf = summary.dropna(subset=["Latitude","Longitude"]).copy()

fig = go.Figure()

for category in ["NORMAL","ELEVATED","HIGH ACTIVITY","HOTSPOT"]:
    d = mapdf[mapdf["Classification"] == category].copy()
    if d.empty:
        continue

    marker_size = np.clip(
        18 + d["Activity_Intensity"] * 26 + np.maximum(d["Rise_Percent"], 0) * 0.25,
        18, 88
    )

    hover_text = []
    for _, r in d.iterrows():
        hover_text.append(
            f"<b>{r['State']}</b><br>"
            f"<b>Classification:</b> {r['Classification']}<br>"
            f"<b>Conversation activity:</b> {int(r['Current_Activity']):,}<br>"
            f"<b>Rise/Fall:</b> {int(r['Rise']):+,}<br>"
            f"<b>Rise %:</b> {r['Rise_Percent']:.1f}%<br><br>"
            f"<b>🔥 What is causing the rise?</b><br>"
            f"{r['Trending_Topic']}<br>"
            f"<i>{r['News_or_Content_Driver']}</i>"
        )

    fig.add_trace(go.Scattergeo(
        lon=d["Longitude"],
        lat=d["Latitude"],
        text=hover_text,
        hovertemplate="%{text}<extra></extra>",
        mode="markers",
        name=category,
        marker=dict(
            size=marker_size,
            color=COLORS[category],
            opacity=0.62,
            line=dict(color="rgba(255,255,255,.88)", width=1.2)
        )
    ))

fig.update_geos(
    projection_type="mercator",
    scope="asia",
    showland=True,
    landcolor="#090b09",
    showocean=True,
    oceancolor="#000000",
    showcountries=True,
    countrycolor="#8c8c8c",
    showcoastlines=True,
    coastlinecolor="#555555",
    lataxis_range=[5, 38],
    lonaxis_range=[67, 99],
    bgcolor="#000000"
)

fig.update_layout(
    height=720,
    paper_bgcolor="#000000",
    plot_bgcolor="#000000",
    margin=dict(l=0,r=0,t=55,b=0),
    font=dict(family="Times New Roman", color="white", size=15),
    # title=dict(
    #     # text="BLACK INDIA HEAT MAP — HOVER OVER A CIRCLE TO SEE THE NEWS/CONTENT DRIVER",
    #     x=0.5,
    #     font=dict(size=19, color="#e6c66f")
    # ),
    legend=dict(
        bgcolor="rgba(15,15,15,.88)",
        bordercolor="rgba(230,198,111,.5)",
        borderwidth=1,
        font=dict(color="white", family="Times New Roman")
    )
)

st.plotly_chart(fig, use_container_width=True)

# ========================= WHAT IS CAUSING THE RISE =========================
st.markdown('<div class="section-title">📰 What News / Content Is Causing the Rise?</div>', unsafe_allow_html=True)

attention = summary[
    (summary["Classification"] != "NORMAL") | (summary["Rise"] > 0)
].sort_values(["Rise","Current_Activity"], ascending=False)

if attention.empty:
    st.info("No positive engagement rise was detected for the selected period.")
else:
    for _, r in attention.iterrows():
        color = COLORS[r["Classification"]]
        st.markdown(
            f"""
            <div class="driver-card" style="border-left-color:{color}">
                <div class="driver-state" style="color:{color} !important">
                    {r['State']} — {r['Classification']}
                </div>
                <br>
                <b>🔥 Trending topic:</b> {r['Trending_Topic']}<br>
                <b>📰 News / content driver:</b> {r['News_or_Content_Driver']}<br>
                <b>📈 Current topic activity:</b> {int(r['Current_Topic_Activity']):,}<br>
                <b>📊 Previous-period topic activity:</b> {int(r['Baseline_Topic_Activity']):,}<br>
                <b>⬆️ Increase linked to this topic:</b> {int(r['Topic_Increase']):+,}
            </div>
            """,
            unsafe_allow_html=True
        )

# ========================= TABLE =========================
st.markdown('<div class="section-title">📊 State-wise Hotspot Classification</div>', unsafe_allow_html=True)

table_cols = [
    "State","Classification","Current_Activity","Baseline_Activity",
    "Rise","Rise_Percent","Trending_Topic","News_or_Content_Driver"
]

display_table = summary[table_cols].copy()
display_table["Rise_Percent"] = display_table["Rise_Percent"].round(1)

st.dataframe(
    display_table.sort_values(["Current_Activity","Rise"], ascending=False),
    use_container_width=True,
    hide_index=True
)

# ========================= STATE DEEP DIVE =========================
st.markdown('<div class="section-title">🔍 State Deep Dive</div>', unsafe_allow_html=True)

chosen_state = st.selectbox("Select a state for detailed analysis", sorted(summary["State"]))
state_summary = summary[summary["State"] == chosen_state].iloc[0]
state_data = filtered[filtered["State"] == chosen_state].copy()

m1,m2,m3,m4 = st.columns(4)
m1.metric("Current Activity", f"{int(state_summary['Current_Activity']):,}")
m2.metric("Baseline Activity", f"{int(state_summary['Baseline_Activity']):,}")
m3.metric("Rise / Fall", f"{int(state_summary['Rise']):+,}")
m4.metric("Classification", state_summary["Classification"])

st.markdown(
    f"""
    <div class="driver-card" style="border-left-color:{COLORS[state_summary['Classification']]}">
        <b>🔥 Main trending topic:</b> {state_summary['Trending_Topic']}<br>
        <b>📰 Content/news driver:</b> {state_summary['News_or_Content_Driver']}<br>
        <b>📈 Topic increase:</b> {int(state_summary['Topic_Increase']):+,}
    </div>
    """,
    unsafe_allow_html=True
)

topic_data = state_data.groupby("Topic", as_index=False)["Conversation_Count"].sum()
topic_data = topic_data.sort_values("Conversation_Count", ascending=False)

if not topic_data.empty:
    chart = px.bar(
        topic_data.head(10).sort_values("Conversation_Count"),
        x="Conversation_Count",
        y="Topic",
        orientation="h",
        text="Conversation_Count"
    )
    chart.update_traces(marker_color="#d9b45d", textposition="outside")
    chart.update_layout(
        title="Topic-wise Conversation Activity",
        paper_bgcolor="#090b09",
        plot_bgcolor="#090b09",
        font=dict(family="Times New Roman", color="white"),
        xaxis=dict(gridcolor="#343434"),
        yaxis=dict(gridcolor="#343434")
    )
    st.plotly_chart(chart, use_container_width=True)

# ========================= PLATFORM COMPARISON =========================
st.markdown('<div class="section-title">📱 Platform-wise Conversation Analysis</div>', unsafe_allow_html=True)

platform_data = filtered.groupby("Platform", as_index=False)["Conversation_Count"].sum()
platform_chart = px.bar(
    platform_data,
    x="Platform",
    y="Conversation_Count",
    text="Conversation_Count"
)
platform_chart.update_traces(marker_color="#6e8c52", textposition="outside")
platform_chart.update_layout(
    paper_bgcolor="#090b09",
    plot_bgcolor="#090b09",
    font=dict(family="Times New Roman", color="white"),
    yaxis=dict(gridcolor="#343434")
)
st.plotly_chart(platform_chart, use_container_width=True)

with st.expander(""):
    st.markdown("""
    **Hotspot classification:** combines current conversation intensity with the percentage change
    from the immediately preceding period of equal length.

    **News/content driver:** for every state, the application compares topic-level activity in the
    selected period against the previous equivalent period. The topic/content with the largest positive
    increase is displayed as the main driver of the rise.

    # **Important:** this is an academic synthetic/modelled dataset for project demonstration.
    # It should not be interpreted as verified official analytics from X, Instagram or Facebook.
    # """)

st.markdown("""
<div style="text-align:center;padding:28px 10px;color:#ddd;">
<b style="color:#e6c66f;">INDIAN ARMY DEFENCE CONVERSATION INTELLIGENCE SYSTEM</b><br>
Academic Project • Geographic Heat Mapping of Defence Conversation<br><br>
Prepared by <b>Nitika Nema</b><br>
Bharati Vidyapeeth (Deemed to be University), College of Engineering, Pune
</div>
""", unsafe_allow_html=True)
