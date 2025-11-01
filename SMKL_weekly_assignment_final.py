import streamlit as st
from datetime import datetime, timedelta, date
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.styles import Font
import io
import random
import pandas as pd

# ---------------------------
# App header
# ---------------------------
st.set_page_config(page_title="SMKL Scheduling Assistant", layout="centered")
st.markdown(
    """
    <h1 style='text-align:center; color:#A3BE8C; margin-bottom:0.25rem;'>
      SMKL Scheduling Assistant — Designed by Timmy Nguyen 😎
    </h1>
    <p style='text-align:center; color:#E0E0E6; opacity:0.9;'>
      Your weekly dispatch planner and assignment assistant.
    </p>
    """,
    unsafe_allow_html=True,
)

# ---------------------------
# Constants
# ---------------------------
DAY_NAMES = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]  # F..L
START_ROW = 14
END_ROW = 90
COL_FIRST = 4  # D
COL_LAST  = 5  # E
COL_SUN   = 6  # F (then +i)

# Helper/HR fixed counts
FIXED_HELPER_ROUTE = 4
FIXED_HELPER       = 4

# ---------------------------
# Utilities
# ---------------------------
def week_start_for(d: date) -> date:
    # Sunday start
    return d - timedelta(days=(d.weekday()+1) % 7)

def normalize_name(s: str) -> str:
    return " ".join(s.split()).strip().casefold()

def read_schedule(uploaded_bytes) -> list[dict]:
    # Read rows 14–90, cols D..L to build driver list
    df = pd.read_excel(uploaded_bytes, sheet_name=0, header=None, skiprows=13, nrows=77)
    drivers = []
    for _, row in df.iterrows():
        first = "" if pd.isna(row[3]) else str(row[3]).strip()
        last  = "" if pd.isna(row[4]) else str(row[4]).strip()
        if not first and not last:
            continue
        days = {DAY_NAMES[i]: row[5+i] for i in range(7)}
        drivers.append({
            "name": f"{first} {last}".strip(),
            "norm": normalize_name(f"{first} {last}"),
            "days": days
        })
    return drivers

def scheduled_on_day(driver: dict, day: str) -> bool:
    v = driver["days"].get(day)
    if pd.isna(v): return False
    s = str(v).strip().lower()
    return s in ("1","dot","dot-commingled")

def dot_certified(driver: dict) -> bool:
    for v in driver["days"].values():
        if pd.isna(v): 
            continue
        if str(v).strip().lower() == "dot":
            return True
    return False

def create_or_replace_sheet(wb, title: str):
    if title in wb.sheetnames:
        ws_old = wb[title]
        wb.remove(ws_old)
    ws = wb.create_sheet(title)
    return ws

def write_daily_sheet(wb, title: str, assignments: dict):
    ws = create_or_replace_sheet(wb, title)

    def write_group(header: str, names: list[str], row: int) -> int:
        ws.cell(row=row, column=1, value=header).font = Font(bold=True)
        row += 1
        if names:
            for n in names:
                ws.cell(row=row, column=1, value="")
                ws.cell(row=row, column=2, value=n)
                row += 1
        else:
            ws.cell(row=row, column=2, value="(none)")
            row += 1
        row += 1
        return row

    # small summary on top
    route_count = len(assignments["DOT"]) + len(assignments["DOT-Commingled"]) + len(assignments["XL"])
    ws.cell(row=1, column=1, value=f"Routes assigned: {route_count}").font = Font(bold=True)
    ws.cell(row=2, column=1, value=f"DOT ({len(assignments['DOT'])})").font = Font(bold=True)
    ws.cell(row=3, column=1, value=f"DOT-Commingled ({len(assignments['DOT-Commingled'])})").font = Font(bold=True)
    ws.cell(row=4, column=1, value=f"DOT-HelperRoute ({len(assignments['DOT-HelperRoute'])})").font = Font(bold=True)
    ws.cell(row=5, column=1, value=f"DOT-Helper ({len(assignments['DOT-Helper'])})").font = Font(bold=True)
    ws.cell(row=6, column=1, value=f"XL ({len(assignments['XL'])})").font = Font(bold=True)
    ws.cell(row=7, column=1, value=f"Standby ({len(assignments['Standby'])})").font = Font(bold=True)

    row = 9
    row = write_group("DOT Routes", assignments["DOT"], row)
    row = write_group("DOT-Commingled", assignments["DOT-Commingled"], row)
    row = write_group("DOT-HelperRoute", assignments["DOT-HelperRoute"], row)
    row = write_group("DOT-Helper", assignments["DOT-Helper"], row)
    row = write_group("XL Routes", assignments["XL"], row)
    row = write_group("Standby", assignments["Standby"], row)

def update_weekly_sheet_values_only(wb, day_label_maps: dict):
    """Replace '1'/'DOT'/'DOT-Commingled' by explicit assignment labels; keep existing cell fills/colors."""
    ws = wb[wb.sheetnames[0]]
    for r in range(START_ROW, END_ROW+1):
        first = ws.cell(r, COL_FIRST).value
        last  = ws.cell(r, COL_LAST ).value
        if not first and not last:
            continue
        name = f"{str(first).strip()} {str(last).strip()}".strip()
        norm = normalize_name(name)
        for day_idx, c in enumerate(range(COL_SUN, COL_SUN+7)):
            val = ws.cell(r, c).value
            s = str(val).strip().lower() if val is not None else ""
            if s in ("1","dot","dot-commingled"):
                label = day_label_maps.get(day_idx, {}).get(norm)
                if label:
                    ws.cell(r, c, value=label)  # value only; do NOT change fill

# ---------------------------
# Inputs
# ---------------------------
uploaded = st.file_uploader("📤 Upload your weekly Excel (.xlsx) — rows 14–90, cols D–L", type=["xlsx"])
if not uploaded:
    st.stop()

week_input = st.number_input("📅 Enter the week number (e.g., 45):", min_value=1, max_value=53, value=45)
today_year = datetime.today().year
wk_start = datetime.fromisocalendar(today_year, week_input, 1).date() - timedelta(days=1)
week_dates = [wk_start + timedelta(days=i) for i in range(7)]
st.info(f"📆 Week {week_input} — starting Sunday {wk_start}")

dot_commingled_eligible_raw = st.text_area("🟣 DOT-Commingled eligible drivers (one per line):", height=120)
new_drivers_raw = st.text_area("🆕 New drivers (XL only; one per line):", height=120)
semi_drivers_raw = st.text_area("🚫 Semi-restricted (cannot do DOT/HelperRoute; one per line):", height=120)

dot_commingled_eligible = [normalize_name(x) for x in dot_commingled_eligible_raw.splitlines() if x.strip()]
new_drivers = [normalize_name(x) for x in new_drivers_raw.splitlines() if x.strip()]
semi_drivers = [normalize_name(x) for x in semi_drivers_raw.splitlines() if x.strip()]

# ---------------------------
# Parse schedule
# ---------------------------
drivers = read_schedule(uploaded)
if not drivers:
    st.error("No drivers parsed. Check your sheet layout (rows 14–90, cols D–L).")
    st.stop()

# Build maps
norm_to_name = {d["norm"]: d["name"] for d in drivers}
dot_map = {d["norm"]: dot_certified(d) for d in drivers}

# ---------------------------
# Assignment loop (per day)
# ---------------------------
results = {}            # { "Sun": {...}, ...}
day_label_maps = {}     # { day_idx: {norm_name: label,...}, ...}

# simple weekly standby cap tracker
standby_cap = {n: 0 for n in dot_map.keys()}  # count per norm name

for tgt in week_dates:
    day_idx = (tgt.weekday()+1) % 7
    day_name = DAY_NAMES[day_idx]

    # Inputs for the day
    st.subheader(f"**{day_name}**")
    dot_routes = st.number_input(f"🚛 {day_name} — Number of DOT routes", min_value=0, value=4, key=f"dot_{day_name}")
    comm_routes = st.number_input(f"🟣 {day_name} — Number of DOT-Commingled routes", min_value=0, value=0, key=f"comm_{day_name}")
    xl_routes = st.number_input(f"📦 {day_name} — Number of XL routes", min_value=0, value=5, key=f"xl_{day_name}")

    # Who is scheduled?
    scheduled_norm = [d["norm"] for d in drivers if scheduled_on_day(d, day_name)]

    # Eligible pools
    eligible_dot = [n for n in scheduled_norm if dot_map.get(n, False) and n not in semi_drivers]
    eligible_comm = [n for n in scheduled_norm if n in dot_commingled_eligible]

    # Assign DOT
    random.shuffle(eligible_dot)
    dot_assigned = eligible_dot[:min(dot_routes, len(eligible_dot))]

    # Assign DOT-Commingled
    remaining_for_step = [n for n in eligible_dot if n not in dot_assigned]
    # Prefer comm-eligible first but ensure they are scheduled
    random.shuffle(eligible_comm)
    comm_assigned = []
    for n in eligible_comm:
        if len(comm_assigned) >= comm_routes:
            break
        if n not in dot_assigned:
            comm_assigned.append(n)

    # Assign DOT-HelperRoute (auto 4)
    remaining_for_step = [n for n in eligible_dot if n not in dot_assigned and n not in comm_assigned]
    random.shuffle(remaining_for_step)
    helper_route = remaining_for_step[:min(FIXED_HELPER_ROUTE, len(remaining_for_step))]

    # Assign DOT-Helper (auto 4)
    non_step_taken = set(dot_assigned + comm_assigned + helper_route)
    helper_pool = [n for n in scheduled_norm if n not in non_step_taken]
    random.shuffle(helper_pool)
    helper = helper_pool[:min(FIXED_HELPER, len(helper_pool))]

    # Assign XL — new drivers (scheduled) first
    xl_assigned = []
    new_sched = [n for n in new_drivers if n in scheduled_norm and n not in non_step_taken and n not in helper]
    random.shuffle(new_sched)
    take_new = min(len(new_sched), xl_routes)
    xl_assigned.extend(new_sched[:take_new])

    remain_need = xl_routes - len(xl_assigned)
    if remain_need > 0:
        others = [n for n in scheduled_norm if n not in set(dot_assigned + comm_assigned + helper_route + helper + xl_assigned)]
        random.shuffle(others)
        xl_assigned.extend(others[:remain_need])

    # Standby = everyone scheduled not assigned, cap ≤ 2 per week
    assigned_today = set(dot_assigned + comm_assigned + helper_route + helper + xl_assigned)
    standby_pool = [n for n in scheduled_norm if n not in assigned_today]
    # apply cap
    standby_final = []
    for n in standby_pool:
        if standby_cap[n] < 2:
            standby_final.append(n)
            standby_cap[n] += 1

    # Store pretty names
    def pretty(list_norm):
        return [norm_to_name.get(n, n) for n in list_norm]

    results[day_name] = {
        "DOT": pretty(dot_assigned),
        "DOT-Commingled": pretty(comm_assigned),
        "DOT-HelperRoute": pretty(helper_route),
        "DOT-Helper": pretty(helper),
        "XL": pretty(xl_assigned),
        "Standby": pretty(standby_final),
    }

    # Build label map (values only; keep colors)
    label_map = {}
    for n in dot_assigned:
        label_map[n] = "DOT"
    for n in comm_assigned:
        label_map[n] = "DOT-Commingled"
    for n in helper_route:
        label_map[n] = "DOT-HelperRoute"
    for n in helper:
        label_map[n] = "DOT-Helper"
    for n in xl_assigned:
        label_map[n] = "XL"
    for n in standby_final:
        label_map[n] = "Standby"
    day_label_maps[day_idx] = label_map

st.markdown("<hr>", unsafe_allow_html=True)

# ---------------------------
# Build output workbook:
# - Update weekly sheet values only
# - Add per-day sheets
# ---------------------------
uploaded.seek(0)
wb = load_workbook(uploaded)
# Update weekly main sheet
update_weekly_sheet_values_only(wb, day_label_maps)

# Add daily sheets "Sun mm_dd", ...
for i, d in enumerate(week_dates):
    title = d.strftime("%a %m_%d")
    write_daily_sheet(wb, title, results[DAY_NAMES[(d.weekday()+1)%7]])

# Save to buffer and offer download
output = io.BytesIO()
wb.save(output)
output.seek(0)

st.success("✅ Daily sheets generated and weekly sheet updated (values only).")
st.download_button(
    "📥 Download: SMKL_Week_{}_updated_assignments.xlsx".format(week_input),
    data=output,
    file_name=f"SMKL_Week_{week_input}_updated_assignments.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
