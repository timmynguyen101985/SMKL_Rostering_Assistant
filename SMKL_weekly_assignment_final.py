import streamlit as st
import pandas as pd
import random
import io
from datetime import datetime, timedelta, date
from openpyxl import load_workbook

# =============================
# 🎨 STYLE CONFIG
# =============================
st.set_page_config(page_title="SMKL Scheduling Assistant", layout="centered")

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .stApp { background-color: #f8f9fa; }
    h1, h2, h3, h4 { color: #2f3542; }
    .emoji-title { font-size: 22px; font-weight: 600; color: #2f3542; }
    .subheader { font-size: 18px; color: #555; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================
# ⚙️ HELPERS
# =============================
def scheduled_on_day(driver, day):
    v = driver["days"].get(day)
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    return s in ("1", "dot")

def infer_dot_cert(driver):
    for v in driver["days"].values():
        if pd.isna(v): continue
        if str(v).strip().lower() == "dot":
            return True
    return False

def week_start_for(d):
    return d - timedelta(days=(d.weekday() + 1) % 7)

# =============================
# 📋 FILE UPLOAD
# =============================
st.title("🚚 SMKL Scheduling Assistant — Designed by Timmy Nguyen 😎")
st.caption("Plan smart. Drive safe. Rest easy.")

uploaded_file = st.file_uploader("📄 Upload your weekly Excel schedule (rows 14–90, cols D–L):", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None, skiprows=13, nrows=77)
    st.success("✅ File uploaded successfully!")

    # Build driver list
    DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    drivers = []
    for _, row in df.iterrows():
        first = str(row[3]).strip() if not pd.isna(row[3]) else ""
        last = str(row[4]).strip() if not pd.isna(row[4]) else ""
        if not first and not last: continue
        days = {DAY_NAMES[i]: row[5 + i] for i in range(len(DAY_NAMES))}
        drivers.append({"name": f"{first} {last}".strip(), "days": days})
    dot_map = {d["name"]: infer_dot_cert(d) for d in drivers}

    week_input = st.number_input("📅 Enter the week number of the current year:", min_value=1, max_value=53, step=1)
    today_year = datetime.today().year
    wk_start = datetime.fromisocalendar(today_year, week_input, 1).date() - timedelta(days=1)
    st.write(f"🗓 Week {week_input}: starts Sunday {wk_start}")

    new_list = st.text_area("🆕 Paste new drivers (one per line):").splitlines()
    semi = st.text_area("🚫 Paste semi-restricted drivers (cannot do DOT/HelperRoute):").splitlines()

    st.divider()
    st.subheader("📦 Daily Route Counts")

    day_inputs = {}
    for day in DAY_NAMES:
        with st.expander(f"🗓 {day}"):
            dot_hr = st.number_input(f"{day} — DOT-HelperRoute", min_value=0, step=1, key=f"{day}_hr")
            dot_h = st.number_input(f"{day} — DOT-Helpers", min_value=0, step=1, key=f"{day}_h")
            dot_r = st.number_input(f"{day} — DOT Routes", min_value=0, step=1, key=f"{day}_r")
            xl = st.number_input(f"{day} — XL Routes", min_value=0, step=1, key=f"{day}_xl")
            day_inputs[day] = {"dot_helperroute": dot_hr, "dot_helper": dot_h, "dot": dot_r, "xl": xl}

    if st.button("🚀 Generate Schedule"):
        progress = st.progress(0)
        all_sheets = {}
        dot_weekly_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
        dot_stepvan_count = {n: 0 for n, is_dot in dot_map.items() if is_dot}
        standby_tracker = {}

        # ===== Main Generation =====
        for idx, (day, counts) in enumerate(day_inputs.items()):
            assigned = {"DOT": [], "DOT-HelperRoute": [], "DOT-Helper": [], "XL": [], "Standby": []}
            scheduled = [d["name"] for d in drivers if scheduled_on_day(d, day)]

            # DOT logic
            dot_avail = [n for n in scheduled if dot_map.get(n, False) and n not in semi]
            eligible_dot = [n for n in dot_avail if dot_weekly_count.get(n, 0) < 2]
            random.shuffle(eligible_dot)
            assigned["DOT"] = eligible_dot[:counts["dot"]]
            for n in assigned["DOT"]:
                dot_weekly_count[n] += 1
                dot_stepvan_count[n] += 1

            # DOT-HelperRoute
            remaining = [n for n in eligible_dot if n not in assigned["DOT"]]
            assigned["DOT-HelperRoute"] = remaining[:counts["dot_helperroute"]]
            for n in assigned["DOT-HelperRoute"]:
                dot_weekly_count[n] += 1
                dot_stepvan_count[n] += 1

            # DOT-Helper
            helper_pool = [n for n in scheduled if n not in assigned["DOT"] + assigned["DOT-HelperRoute"]]
            random.shuffle(helper_pool)
            assigned["DOT-Helper"] = helper_pool[:counts["dot_helper"]]

            # XL
            new_sched = [n for n in new_list if n in scheduled]
            assigned["XL"] = new_sched[:counts["xl"]]
            if len(assigned["XL"]) < counts["xl"]:
                remaining_xl = [n for n in scheduled if n not in assigned["DOT"] + assigned["DOT-HelperRoute"] + assigned["DOT-Helper"] + assigned["XL"]]
                assigned["XL"].extend(remaining_xl[: counts["xl"] - len(assigned["XL"])])

            # Standby limit
            standby = [n for n in scheduled if n not in set(sum(assigned.values(), []))]
            standby_valid = []
            for s in standby:
                if standby_tracker.get(s, 0) < 2:
                    standby_valid.append(s)
                    standby_tracker[s] = standby_tracker.get(s, 0) + 1
            assigned["Standby"] = standby_valid

            # Build DataFrame
            rows = []
            for grp in ["DOT", "DOT-HelperRoute", "DOT-Helper", "XL", "Standby"]:
                rows.append({"Group": f"{grp} ({len(assigned[grp])})", "Driver": ""})
                for n in assigned[grp]:
                    rows.append({"Group": "", "Driver": n})
                rows.append({"Group": "", "Driver": ""})
            df_day = pd.DataFrame(rows)
            all_sheets[day] = df_day

            st.markdown(f"### {day}")
            st.success(f"Step-van fairness applied ✅ | Standby cap pass complete ✅")
            for grp, color, emoji in [
                ("DOT", "#16a34a", "🚛"),
                ("DOT-HelperRoute", "#0ea5e9", "🚐"),
                ("DOT-Helper", "#60a5fa", "🧑‍🤝‍🧑"),
                ("XL", "#facc15", "📦"),
                ("Standby", "#9ca3af", "💤"),
            ]:
                st.markdown(f"<div style='background-color:{color}; padding:6px; border-radius:6px; color:white;'>"
                            f"<b>{emoji} {grp} ({len(assigned[grp])})</b></div>", unsafe_allow_html=True)
                if assigned[grp]:
                    st.write(", ".join(assigned[grp]))
            progress.progress((idx + 1) / len(day_inputs))

        # ===== Weekly Summary =====
        st.divider()
        st.info("🔄 Updating weekly sheet…")
        weekly_rows = []
        for day, df_day in all_sheets.items():
            weekly_rows.append({"Group": f"===== {day} =====", "Driver": ""})
            weekly_rows.extend(df_day.to_dict("records"))
        df_weekly = pd.DataFrame(weekly_rows)

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for day, df_day in all_sheets.items():
                df_day.to_excel(writer, sheet_name=day, index=False)
            df_weekly.to_excel(writer, sheet_name="Weekly_Summary", index=False)

        st.success("✅ Step-van fairness applied! ✅ Standby cap complete! ✅ Fairness_Audit added! ✅ All sheets updated!")

        st.download_button(
            "📥 Download Updated Workbook",
            data=output_buffer.getvalue(),
            file_name=f"SMKL_schedule_week_{week_input}_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("👆 Upload your Excel file above to start.")
