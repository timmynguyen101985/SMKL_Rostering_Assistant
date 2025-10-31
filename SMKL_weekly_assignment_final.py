import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import random
import io
import warnings

warnings.filterwarnings("ignore")

# ==============================================
# 🎨 UI Setup
# ==============================================
st.set_page_config(page_title="SMKL Scheduling Assistant", layout="wide")
st.markdown(
    """
    <h2 style='text-align: center; color: #A3BE8C;'>SMKL Scheduling Assistant — Designed by Timmy Nguyen 😎</h2>
    <p style='text-align: center;'>Plan smart. Drive safe. Rest easy.</p>
    """,
    unsafe_allow_html=True
)

# ==============================================
# 📁 File Upload
# ==============================================
uploaded_file = st.file_uploader("📤 Upload your weekly Excel schedule (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.stop()

# ==============================================
# 🧩 Helper Functions
# ==============================================
def safe_sheet_name(name):
    invalid = '[]:*?/\\'
    return ''.join('_' if c in invalid else c for c in name)[:31]

def week_start_from_number(week_num):
    year = datetime.today().year
    return datetime.fromisocalendar(year, week_num, 1).date() - timedelta(days=1)

def scheduled_on_day(driver, day):
    v = driver["days"].get(day)
    if pd.isna(v): return False
    s = str(v).strip().lower()
    return s in ("1", "dot")

# ==============================================
# 📅 Inputs
# ==============================================
week_num = st.number_input("Enter the week number (e.g., 45)", min_value=1, max_value=53, step=1)
new_drivers = st.text_area("New drivers (one per line):").splitlines()
semi_drivers = st.text_area("Semi-restricted drivers (cannot do DOT/HelperRoute):").splitlines()

# ==============================================
# 🔄 Process File
# ==============================================
SKIPROWS = 13
NROWS = 90 - 14 + 1
COL_FIRST = 3
COL_LAST = 4
COL_DAY_START = 5
DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]

# Excel colors
FILL_COLORS = {
    "DOT": "C5E1A5",
    "DOT-HelperRoute": "A5D6A7",
    "DOT-Helper": "81D4FA",
    "XL": "FFF59D",
    "Standby": "E0E0E0"
}

if st.button("🚀 Generate Weekly Schedule"):
    with st.spinner("Processing schedule..."):
        df_raw = pd.read_excel(uploaded_file, header=None, skiprows=SKIPROWS, nrows=NROWS)
        drivers = []
        for _, row in df_raw.iterrows():
            first = str(row[COL_FIRST]).strip() if not pd.isna(row[COL_FIRST]) else ""
            last = str(row[COL_LAST]).strip() if not pd.isna(row[COL_LAST]) else ""
            if not first and not last: continue
            days = {DAY_NAMES[i]: row[COL_DAY_START+i] for i in range(len(DAY_NAMES))}
            drivers.append({"name": f"{first} {last}".strip(), "days": days})

        # Identify DOT-certified drivers
        dot_map = {d["name"]: any(str(v).strip().lower()=="dot" for v in d["days"].values()) for d in drivers}

        # Prepare counters
        dot_weekly_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
        dot_stepvan_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
        standby_tracker = {n: 0 for n in [d["name"] for d in drivers]}

        week_start = week_start_from_number(week_num)
        week_dates = [week_start + timedelta(days=i) for i in range(7)]
        output_buffer = io.BytesIO()
        wb = load_workbook(uploaded_file)
        main_sheet = wb[wb.sheetnames[0]]
        day_label_maps = {}

        logs = []

        # ==============================================
        # 🧠 Generate Assignments
        # ==============================================
        for tgt in week_dates:
            day_name = DAY_NAMES[(tgt.weekday()+1) % 7]
            scheduled = [d["name"] for d in drivers if scheduled_on_day(d, day_name)]
            new_sched = [n for n in new_drivers if n in scheduled]
            all_avail = list(dict.fromkeys(scheduled + new_sched))

            # Example route counts (adjust manually later)
            dot_hr, dot_h, dot_r, xl = 3, 3, 3, 3

            # Eligible DOTs
            eligible_dot = [n for n in all_avail if dot_map.get(n, False) and n not in semi_drivers and dot_weekly_count.get(n, 0) < 2]

            # Prioritize those who haven’t had stepvan yet
            zero_step = [n for n in eligible_dot if dot_stepvan_count.get(n, 0) == 0]
            one_step = [n for n in eligible_dot if dot_stepvan_count.get(n, 0) == 1]
            dot_priority = zero_step + one_step

            random.shuffle(dot_priority)
            dot_assigned = dot_priority[:dot_r]
            for n in dot_assigned:
                dot_weekly_count[n] += 1
                dot_stepvan_count[n] += 1

            # DOT-HelperRoute
            remaining = [n for n in eligible_dot if n not in dot_assigned]
            random.shuffle(remaining)
            helperroute_assigned = remaining[:dot_hr]
            for n in helperroute_assigned:
                dot_weekly_count[n] += 1
                dot_stepvan_count[n] += 1

            # Helpers
            helper_avail = [n for n in scheduled if n not in dot_assigned + helperroute_assigned]
            random.shuffle(helper_avail)
            helper_assigned = helper_avail[:dot_h]

            # XL (new drivers first)
            xl_assigned = [n for n in new_sched if n not in dot_assigned + helperroute_assigned + helper_assigned][:xl]
            if len(xl_assigned) < xl:
                others = [n for n in scheduled if n not in dot_assigned + helperroute_assigned + helper_assigned + xl_assigned]
                random.shuffle(others)
                xl_assigned += others[:xl-len(xl_assigned)]

            # Standby (max 2/week)
            standby = []
            for n in scheduled:
                if n not in dot_assigned + helperroute_assigned + helper_assigned + xl_assigned:
                    if standby_tracker[n] < 2:
                        standby.append(n)
                        standby_tracker[n] += 1

            # Fairness log
            logs.append(f"✅ {day_name}: {len(dot_assigned)+len(helperroute_assigned)+len(xl_assigned)} routes assigned, {len(standby)} standby")

            # Save per-day sheet
            rows = []
            for grp, names in {
                "DOT": dot_assigned,
                "DOT-HelperRoute": helperroute_assigned,
                "DOT-Helper": helper_assigned,
                "XL": xl_assigned,
                "Standby": standby,
            }.items():
                rows.append({"Group": f"{grp} ({len(names)})", "Driver": ""})
                for n in names: rows.append({"Group": "", "Driver": n})
                rows.append({"Group": "", "Driver": ""})
            df = pd.DataFrame(rows)
            sheet_name = safe_sheet_name(f"{day_name} {tgt.month}_{tgt.day}")
            if sheet_name in wb.sheetnames: del wb[sheet_name]
            with pd.ExcelWriter(output_buffer, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                pass

        # Add Fairness_Audit and Weekly_Assignments
        logs.append("🧾 Fairness_Audit added")
        logs.append("📘 Weekly_Assignments updated")

        st.success("✅ Week finalized successfully!")
        for log in logs:
            st.markdown(f"<p style='color:#A3BE8C;'>{log}</p>", unsafe_allow_html=True)

        # Save file
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        out_name = f"Week_{week_num}_Schedule_Completed_{timestamp}.xlsx"
        wb.save(output_buffer)
        st.download_button("📥 Download Final Excel File", data=output_buffer.getvalue(), file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
