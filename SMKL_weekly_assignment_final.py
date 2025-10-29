import io
import random
from datetime import datetime, timedelta, date

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# -------------------------------
# App Title / Header / Reset
# -------------------------------
st.set_page_config(page_title="SMKL Scheduling Assistant", layout="wide")
left, right = st.columns([0.8, 0.2])
with left:
    st.markdown("## SMKL Scheduling Assistant — Designed by Timmy Nguyen 😎")
with right:
    if st.button("🔄 Reset / Start Over", use_container_width=True):
        st.experimental_rerun()

st.write("Upload your weekly schedule and enter counts per day. This app will assign routes, "
         "enforce step-van fairness for DOT drivers, cap Standby at 2 per week, and export a "
         "color-coded Excel workbook with daily sheets + Fairness_Audit (if swaps occurred).")

# -------------------------------
# Constants / Layout
# -------------------------------
SKIPROWS = 13            # rows 14..90 inclusive
NROWS = 90 - 14 + 1
COL_FIRST = 3            # D
COL_LAST  = 4            # E
COL_DAY_START = 5        # F is Sunday
DAY_NAMES = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]

# Caps / Rules
STANDBY_CAP = 2
MAX_DOT_WEEK = 2   # Counts DOT / DOT-HelperRoute / DOT-Helper. XL & Standby don't count.
ENFORCE_MAX_DOT = True  # Set False if you want fairness only.

# Colors
FILL_DOT        = PatternFill(start_color="C5E1A5", end_color="C5E1A5", fill_type="solid")  # light green
FILL_DOT_HR     = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")  # darker green
FILL_HELPER     = PatternFill(start_color="81D4FA", end_color="81D4FA", fill_type="solid")  # light blue
FILL_XL         = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")  # yellow
FILL_STANDBY    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")  # grey
HEADER_FILL     = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")  # banner
BOLD_FONT       = Font(bold=True)

ASSIGN_LABEL_TO_FILL = {
    "DOT": FILL_DOT,
    "DOT-HelperRoute": FILL_DOT_HR,
    "DOT-Helper": FILL_HELPER,
    "XL": FILL_XL,
    "Standby": FILL_STANDBY,
}

# -------------------------------
# Helpers
# -------------------------------
def read_schedule_df(uploaded_file):
    # First sheet only; sheet shape is fixed by your template
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None,
                       skiprows=SKIPROWS, nrows=NROWS, engine="openpyxl")
    return df

def build_driver_records(df):
    drivers = []
    for _, row in df.iterrows():
        first = str(row[COL_FIRST]).strip() if not pd.isna(row[COL_FIRST]) else ""
        last  = str(row[COL_LAST]).strip()  if not pd.isna(row[COL_LAST])  else ""
        if not first and not last:
            continue
        days = {DAY_NAMES[i]: row[COL_DAY_START + i] for i in range(7)}
        drivers.append({"name": f"{first} {last}".strip(), "days": days})
    return drivers

def infer_dot_map(drivers):
    dot_map = {}
    for d in drivers:
        is_dot = any((not pd.isna(v)) and str(v).strip().lower() == "dot"
                     for v in d["days"].values())
        dot_map[d["name"]] = is_dot
    return dot_map

def scheduled_on_day(driver, day):
    v = driver["days"].get(day)
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    return s in ("1", "dot")

def week_start_from_weeknum(week_num, year=None):
    if year is None:
        year = datetime.today().year
    monday = datetime.fromisocalendar(year, int(week_num), 1).date()
    return monday - timedelta(days=1)  # Sunday

def choose(pool, k):
    if k <= 0:
        return []
    pool2 = pool[:]  # copy
    random.shuffle(pool2)
    return pool2[:k]

# -------------------------------
# Fairness & Assignment Engine
# -------------------------------
def generate_week(drivers, dot_map, new_drivers, semi_restricted, week_dates, daily_counts):
    """
    Returns:
      - day_assignments: dict day_idx -> dict{ group -> [names], 'Pairings': [(driver, helper)] }
      - swaps_log: list of text lines describing fairness swaps
      - day_summaries: list of dict rows for the UI table
    """
    # Trackers
    dot_weekly_count = {n: 0 for n, isdot in dot_map.items() if isdot}
    stepvan_count    = {n: 0 for n, isdot in dot_map.items() if isdot}  # DOT or DOT-HelperRoute
    standby_count    = {}  # all drivers

    day_assignments = {}
    day_summaries = []
    swaps_log = []

    # First pass: day-by-day assignment
    for idx, dt in enumerate(week_dates):
        day_name = DAY_NAMES[(dt.weekday() + 1) % 7]
        counts = daily_counts[idx]  # dict with dot, dot_hr, dot_h, xl
        scheduled = [d["name"] for d in drivers if scheduled_on_day(d, day_name)]

        # Pools
        new_sched = [n for n in new_drivers if n in scheduled]
        dot_avail = [n for n in scheduled if dot_map.get(n, False) and (n not in semi_restricted)]
        helpers_avail = [n for n in scheduled]  # helpers can be anyone scheduled

        # Enforce weekly DOT cap by filtering (optional)
        if ENFORCE_MAX_DOT:
            dot_avail = [n for n in dot_avail if dot_weekly_count.get(n, 0) < MAX_DOT_WEEK]

        # Prioritize DOT with fewer step-van so far
        zero_step = [n for n in dot_avail if stepvan_count.get(n, 0) == 0]
        one_step  = [n for n in dot_avail if stepvan_count.get(n, 0) == 1]
        more_step = [n for n in dot_avail if stepvan_count.get(n, 0) >= 2]
        dot_priority = zero_step + one_step + more_step

        # Assign DOT and DOT-HelperRoute
        dot_assigned = []
        dot_needed = counts["dot"]
        for n in dot_priority:
            if len(dot_assigned) >= dot_needed:
                break
            dot_assigned.append(n)

        rem_dot_pool = [n for n in dot_priority if n not in dot_assigned]
        dot_hr_assigned = choose(rem_dot_pool, counts["dot_hr"])

        # Helpers (1:1 pairing for helperroute will be done after selecting helpers)
        # First choose helpers that are not already DOT step-van drivers this day
        occupied = set(dot_assigned + dot_hr_assigned)
        helpers_pool = [n for n in helpers_avail if n not in occupied]
        dot_h_assigned = choose(helpers_pool, counts["dot_h"])

        # Update DOT weekly counters
        for n in dot_assigned + dot_hr_assigned:
            if dot_map.get(n, False):
                dot_weekly_count[n] = dot_weekly_count.get(n, 0) + 1
                stepvan_count[n]    = stepvan_count.get(n, 0) + 1
        for n in dot_h_assigned:
            if dot_map.get(n, False):
                dot_weekly_count[n] = dot_weekly_count.get(n, 0) + 1

        # XL: new scheduled first, then others scheduled
        occupied = set(dot_assigned + dot_hr_assigned + dot_h_assigned)
        xl_assigned = []
        need_xl = counts["xl"]

        take_new = min(len(new_sched), need_xl)
        if take_new > 0:
            xl_assigned.extend(choose(new_sched, take_new))
            need_xl -= take_new

        if need_xl > 0:
            remaining_sched = [n for n in scheduled if n not in occupied and n not in xl_assigned]
            xl_assigned.extend(choose(remaining_sched, need_xl))

        # Standby = scheduled not assigned anywhere
        occupied = set(dot_assigned + dot_hr_assigned + dot_h_assigned + xl_assigned)
        standby = [n for n in scheduled if n not in occupied]

        # Pair DOT-HelperRoute with helpers from dot_h_assigned (1:1)
        pairs = []
        helpers_for_pair = dot_h_assigned[:]
        random.shuffle(helpers_for_pair)
        paired_helpers = []
        for d in dot_hr_assigned:
            h = helpers_for_pair.pop(0) if helpers_for_pair else None
            if h:
                pairs.append((d, h))
                paired_helpers.append(h)

        # Build assignments for the day
        assignments = {
            "DOT": dot_assigned,
            "DOT-HelperRoute": dot_hr_assigned,
            "DOT-Helper": dot_h_assigned,
            "XL": xl_assigned,
            "Standby": standby,
            "Pairings": pairs
        }
        day_assignments[idx] = assignments

        # Update standby counts
        for n in standby:
            standby_count[n] = standby_count.get(n, 0) + 1

        # Save summary counts
        day_summaries.append({
            "Day": dt.strftime("%a %m/%d"),
            "DOT": len(dot_assigned),
            "DOT-HelperRoute": len(dot_hr_assigned),
            "DOT-Helper": len(dot_h_assigned),
            "XL": len(xl_assigned),
            "Standby": len(standby),
            "Scheduled": len(scheduled),
            "Routes": len(dot_assigned) + len(dot_hr_assigned) + len(xl_assigned)
        })

    # -------------------------------
    # Fairness: ensure each DOT driver gets ≥1 step-van this week
    # -------------------------------
    st.session_state.progress_text.write("✅ Step-van fairness: applying...")
    unserved_dot = [n for n in stepvan_count if stepvan_count[n] == 0]
    # Try to swap from days with step-van assignments
    for driver in unserved_dot:
        swapped = False
        for idx, dt in enumerate(week_dates):
            day_name = DAY_NAMES[(dt.weekday() + 1) % 7]
            # driver must be scheduled that day
            try:
                drv_obj = next(d for d in drivers if d["name"] == driver)
            except StopIteration:
                continue
            if not scheduled_on_day(drv_obj, day_name):
                continue

            # find someone to swap OUT of step-van that day (DOT or DOT-HR)
            todays = day_assignments[idx]
            candidates = todays["DOT"] + todays["DOT-HelperRoute"]
            if not candidates:
                continue

            # Prefer candidate with stepvan_count >= 2
            c_sorted = sorted(candidates, key=lambda x: -stepvan_count.get(x, 0))
            for victim in c_sorted:
                if victim == driver:
                    continue
                # Swap: give driver a step-van slot; move victim somewhere else (XL or Standby)
                # Remove victim from whichever group they are in
                if victim in todays["DOT"]:
                    todays["DOT"].remove(victim)
                    target_group = "DOT"
                else:
                    todays["DOT-HelperRoute"].remove(victim)
                    target_group = "DOT-HelperRoute"

                # Place driver into that group
                todays[target_group].append(driver)
                stepvan_count[driver] = stepvan_count.get(driver, 0) + 1

                # Re-assign victim: try XL first, else Standby (respect standby cap)
                if victim not in todays["XL"]:
                    todays["XL"].append(victim)
                # If standby cap exceeded later, we'll repair below

                swaps_log.append(f"{dt.strftime('%a %m/%d')}: swapped in {driver} (step-van), moved {victim} to XL")
                swapped = True
                break
            if swapped:
                break

    # -------------------------------
    # Standby cap repair (≤2 per week per driver)
    # -------------------------------
    st.session_state.progress_text.write("✅ Standby cap: enforcing ≤ 2 per driver...")
    # Recompute weekly standby tallies from day_assignments and fix violations
    tally = {}
    for idx in day_assignments:
        for n in day_assignments[idx]["Standby"]:
            tally[n] = tally.get(n, 0) + 1

    # For anyone > 2, try moving them into XL (or DOT-Helper) where possible
    for name, cnt in list(tally.items()):
        while cnt > STANDBY_CAP:
            # find a day they were standby and move them out if any slot free
            moved = False
            for idx in range(7):
                if name in day_assignments[idx]["Standby"]:
                    # Prefer XL
                    day_assignments[idx]["Standby"].remove(name)
                    day_assignments[idx]["XL"].append(name)
                    cnt -= 1
                    moved = True
                    swaps_log.append(f"{week_dates[idx].strftime('%a %m/%d')}: reduced standby for {name} → moved to XL")
                    break
            if not moved:
                # no move possible
                break

    # Done
    return day_assignments, swaps_log, day_summaries


def write_colored_workbook(day_assignments, week_dates, swaps_log, filename_bytes):
    """
    Build an Excel workbook with Sun..Sat sheets and an optional Fairness_Audit sheet.
    Color code rows per group.
    """
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    def add_sheet(title, rows):
        ws = wb.create_sheet(title[:31])
        # Header
        ws.append(["Group", "Driver"])
        for cell in ws[1]:
            cell.font = BOLD_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")
        # Data with fills
        for group, driver in rows:
            ws.append([group, driver])
            r = ws.max_row
            fill = ASSIGN_LABEL_TO_FILL.get(group)
            if fill:
                ws[f"A{r}"].fill = fill
                ws[f"B{r}"].fill = fill
        # Widths
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 34

    # Create daily sheets
    for idx, dt in enumerate(week_dates):
        title = dt.strftime("%a %m_%d")
        asg = day_assignments[idx]
        rows = []
        # Group order
        for g in ["DOT", "DOT-HelperRoute", "DOT-Helper", "XL", "Standby"]:
            rows.append((g, ""))  # header spacer row
            for n in asg[g]:
                rows.append((g, n))
            rows.append(("", ""))  # blank spacer
        add_sheet(title, rows)

    # Fairness_Audit
    if swaps_log:
        ws = wb.create_sheet("Fairness_Audit")
        ws.append(["Action"])
        ws["A1"].font = BOLD_FONT
        ws["A1"].fill = HEADER_FILL
        for line in swaps_log:
            ws.append([line])
        ws.column_dimensions["A"].width = 80

    # Write to BytesIO
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename_bytes.write(bio.getvalue())


# -------------------------------
# Sidebar Inputs
# -------------------------------
with st.sidebar:
    st.markdown("### 1) Upload Weekly Excel")
    uploaded = st.file_uploader("Excel (rows 14–90, cols D–L)", type=["xlsx"])

    st.markdown("---")
    st.markdown("### 2) Drivers")
    new_drivers_text = st.text_area("New drivers (one per line)", height=120,
                                    placeholder="Example:\nJohn Doe\nMaria Perez")
    semi_text = st.text_area("Semi-restricted (no DOT / no DOT-HelperRoute)", height=120,
                             placeholder="Example:\nAlex Pham\nNina Tran")

    st.markdown("---")
    st.markdown("### 3) Week Number")
    week_num = st.number_input("Enter week number (ISO, e.g., 45)", min_value=1, max_value=53, value=45, step=1)

    st.markdown("---")
    st.markdown("### 4) Daily Route Counts")
    counts = {}
    for d in DAY_NAMES:
        st.markdown(f"**{d}**")
        dot  = st.number_input(f"{d} — DOT", min_value=0, value=3, step=1, key=f"{d}_dot")
        dhr  = st.number_input(f"{d} — DOT-HelperRoute", min_value=0, value=3, step=1, key=f"{d}_dhr")
        dh   = st.number_input(f"{d} — DOT-Helper", min_value=0, value=2, step=1, key=f"{d}_dh")
        xl   = st.number_input(f"{d} — XL", min_value=0, value=10, step=1, key=f"{d}_xl")
        counts[d] = {"dot": dot, "dot_hr": dhr, "dot_h": dh, "xl": xl}

    st.markdown("---")
    run_btn = st.button("🚀 Generate Week", type="primary", use_container_width=True)

# Placeholders for progress and logs
st.markdown("---")
progress = st.progress(0)
st.session_state.progress_text = st.empty()

# -------------------------------
# Run Generation
# -------------------------------
if run_btn:
    if not uploaded:
        st.error("Please upload the weekly Excel file.")
        st.stop()

    try:
        df = read_schedule_df(uploaded)
    except Exception as e:
        st.error(f"Could not read Excel: {e}")
        st.stop()

    drivers = build_driver_records(df)
    if not drivers:
        st.error("No drivers parsed. Check Excel layout (rows 14–90, D–L).")
        st.stop()

    dot_map = infer_dot_map(drivers)
    new_drivers = [x.strip() for x in new_drivers_text.split("\n") if x.strip()]
    semi_restricted = [x.strip() for x in semi_text.split("\n") if x.strip()]

    wk_start = week_start_from_weeknum(week_num)
    week_dates = [wk_start + timedelta(days=i) for i in range(7)]

    daily_counts = [counts[d] for d in DAY_NAMES]

    progress.progress(5); st.session_state.progress_text.write("📥 Input loaded.")
    day_assignments, swaps_log, day_summaries = generate_week(
        drivers, dot_map, new_drivers, semi_restricted, week_dates, daily_counts
    )

    progress.progress(60); st.session_state.progress_text.write("🔁 Step-van fairness applied & Standby cap enforced.")

    # --- UI: per-day panels
    st.markdown("### Daily Assignments")
    for idx, dt in enumerate(week_dates):
        label = dt.strftime("%a %m/%d")
        asg = day_assignments[idx]
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(f"**🚛 DOT ({len(asg['DOT'])})**")
            st.write("\n".join(asg["DOT"]) or "_none_")
        with c2:
            st.markdown(f"**🚐 DOT-HelperRoute ({len(asg['DOT-HelperRoute'])})**")
            st.write("\n".join(asg["DOT-HelperRoute"]) or "_none_")
        with c3:
            st.markdown(f"**🧑‍🤝‍🧑 DOT-Helper ({len(asg['DOT-Helper'])})**")
            st.write("\n".join(asg["DOT-Helper"]) or "_none_")
        with c4:
            st.markdown(f"**📦 XL ({len(asg['XL'])})**")
            st.write("\n".join(asg["XL"]) or "_none_")
        with c5:
            st.markdown(f"**💤 Standby ({len(asg['Standby'])})**")
            st.write("\n".join(asg["Standby"]) or "_none_")
        st.caption(label)
        st.divider()

    progress.progress(80); st.session_state.progress_text.write("🧾 Building summary & workbook...")

    # --- Summary table
    st.markdown("### Weekly Summary (counts by day)")
    summary_df = pd.DataFrame(day_summaries)
    st.dataframe(summary_df, use_container_width=True)

    # --- Fairness_Audit log
    if swaps_log:
        with st.expander("Fairness_Audit (swaps performed)"):
            st.write("\n".join(swaps_log))

    # --- Build colored Excel in-memory
    buf = io.BytesIO()
    write_colored_workbook(day_assignments, week_dates, swaps_log, buf)
    progress.progress(100); st.session_state.progress_text.write("✅ All done! Ready to download.")

    out_name = f"Week{int(week_num)}_Schedule.xlsx"
    st.download_button(
        label=f"⬇️ Download {out_name}",
        data=buf.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

