import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta, date
import io, random

# ---------------------------------------------
# App Title
# ---------------------------------------------
st.set_page_config(page_title="SMKL Scheduling Assistant", layout="centered")
st.markdown(
    """
    <h1 style='text-align:center; color:#A3BE8C;'>
        SMKL Scheduling Assistant — Designed by Timmy Nguyen 😎
    </h1>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------
# Constants & Colors
# ---------------------------------------------
DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
COLOR_FILL = {
    "DOT": PatternFill(start_color="C5E1A5", end_color="C5E1A5", fill_type="solid"),
    "DOT-Commingled": PatternFill(start_color="CE93D8", end_color="CE93D8", fill_type="solid"),
    "DOT-HelperRoute": PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid"),
    "DOT-Helper": PatternFill(start_color="81D4FA", end_color="81D4FA", fill_type="solid"),
    "XL": PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid"),
    "Standby": PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"),
}

# ---------------------------------------------
# File Upload
# ---------------------------------------------
uploaded_file = st.file_uploader("📤 Upload your weekly Excel schedule (rows 14–90, cols D–L):", type=["xlsx"])
if not uploaded_file:
    st.stop()

# ---------------------------------------------
# Helper Functions
# ---------------------------------------------
def read_schedule(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None, skiprows=13, nrows=77)
    drivers = []
    for _, row in df.iterrows():
        first, last = str(row[3]).strip(), str(row[4]).strip()
        if not first or not last or first == "nan" or last == "nan":
            continue
        days = {DAY_NAMES[i]: row[5 + i] for i in range(7)}
        drivers.append({"name": f"{first} {last}".strip(), "days": days})
    return drivers

def scheduled_on_day(driver, day):
    v = driver["days"].get(day)
    return str(v).strip().lower() in ("1", "dot", "dot-commingled")

# ---------------------------------------------
# User Inputs
# ---------------------------------------------
week_input = st.number_input("📅 Enter the week number (e.g., 45):", min_value=1, max_value=53, value=45)
today_year = datetime.today().year
wk_start = datetime.fromisocalendar(today_year, week_input, 1).date() - timedelta(days=1)
week_dates = [wk_start + timedelta(days=i) for i in range(7)]
st.info(f"📆 Week {week_input} detected – starting Sunday {wk_start}")

dot_commingled_eligible = st.text_area("👷 Enter DOT-Commingled eligible drivers (one per line):").splitlines()
new_drivers = st.text_area("🆕 Enter new drivers (for XL routes only, one per line):").splitlines()
semi_drivers = st.text_area("🚫 Enter semi-restricted drivers (cannot do DOT/HelperRoute, one per line):").splitlines()

drivers = read_schedule(uploaded_file)
dot_map = {d["name"]: any(str(v).strip().lower() == "dot" for v in d["days"].values()) for d in drivers}

dot_weekly_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
dot_stepvan_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
standby_tracker = {n: 0 for n in dot_map}

# ---------------------------------------------
# Assignment Generator
# ---------------------------------------------
results = {}
for tgt in week_dates:
    day_name = DAY_NAMES[(tgt.weekday() + 1) % 7]
    scheduled = [d["name"] for d in drivers if scheduled_on_day(d, day_name)]

    # Input only DOT, DOT-Commingled, XL
    dot_routes = st.number_input(f"🚛 {day_name}: Number of DOT routes", min_value=0, value=4)
    comm_routes = st.number_input(f"🟣 {day_name}: Number of DOT-Commingled routes", min_value=0, value=2)
    xl_routes = st.number_input(f"📦 {day_name}: Number of XL routes", min_value=0, value=5)

    eligible_dots = [n for n in scheduled if dot_map.get(n, False) and n not in semi_drivers]
    comm_eligible = [n for n in dot_commingled_eligible if n in scheduled]

    dot_assigned = random.sample(eligible_dots, min(dot_routes, len(eligible_dots)))
    comm_assigned = random.sample(comm_eligible, min(comm_routes, len(comm_eligible)))

    helper_route = random.sample(
        [n for n in eligible_dots if n not in dot_assigned + comm_assigned],
        min(4, len(eligible_dots)),
    )
    helper = random.sample(
        [n for n in scheduled if n not in dot_assigned + comm_assigned + helper_route],
        min(4, len(scheduled)),
    )

    xl_assigned = [n for n in new_drivers if n in scheduled][:xl_routes]

    standby = [
        n for n in scheduled
        if n not in set(dot_assigned + comm_assigned + helper_route + helper + xl_assigned)
    ]
    standby = [n for n in standby if standby_tracker[n] < 2]
    for s in standby:
        standby_tracker[s] += 1

    results[day_name] = {
        "DOT": dot_assigned,
        "DOT-Commingled": comm_assigned,
        "DOT-HelperRoute": helper_route,
        "DOT-Helper": helper,
        "XL": xl_assigned,
        "Standby": standby,
    }

# ---------------------------------------------
# Update Workbook
# ---------------------------------------------
wb = load_workbook(uploaded_file)
ws = wb[wb.sheetnames[0]]

for r in range(14, 91):
    first, last = ws.cell(r, 4).value, ws.cell(r, 5).value
    if not first and not last:
        continue
    name = f"{str(first).strip()} {str(last).strip()}".strip()
    for i, day in enumerate(DAY_NAMES):
        val = ws.cell(r, 6 + i).value
        if str(val).strip().lower() in ("1", "dot", "dot-commingled"):
            for k, v in results.get(day, {}).items():
                if name in v:
                    ws.cell(r, 6 + i).value = k
                    ws.cell(r, 6 + i).fill = COLOR_FILL.get(k)

# ---------------------------------------------
# Export
# ---------------------------------------------
output = io.BytesIO()
wb.save(output)
output.seek(0)

st.success("✅ All assignments generated and weekly sheet updated successfully!")
st.download_button(
    label="📥 Download Updated Weekly Schedule",
    data=output,
    file_name=f"SMKL_Week_{week_input}_updated_assignments.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
