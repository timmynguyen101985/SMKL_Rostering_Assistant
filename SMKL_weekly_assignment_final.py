"""
SMKL_weekly_assignment_v10.py

Lightweight dispatch assignment generator
- Full week (Sun–Sat), calm console UI
- No JSON / snapshots
- Enforces DOT weekly limits (max 2 per week across DOT / DOT-HelperRoute / DOT-Helper)
  * Standby and XL DO NOT count toward the DOT limit
- Prioritizes giving every DOT driver at least one step-van slot weekly
- New drivers are always assigned to XL
- Standby capped to 2 per driver per week
- Writes per-day sheets, updates weekly schedule (replaces '1'/'DOT' with labels + colors)
- Clean and simple output
"""

# =========================
# 🔹 IMPORTS & CONSTANTS
# =========================
import os
import random
import time
import warnings
from datetime import datetime, timedelta, date
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
# import tkinter as tk
# from tkinter import filedialog
from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel


st.title("📘 SMKL Dispatch Assistant — Web Version")
uploaded_file = st.file_uploader("Upload your weekly Excel schedule", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.success("✅ File uploaded successfully!")
else:
    st.warning("Please upload a file to continue.")


warnings.filterwarnings(
    "ignore",
    message="Unknown extension is not supported and will be removed",
    category=UserWarning
)

console = Console()

# Layout constants
SKIPROWS = 13
NROWS = 90 - 14 + 1
COL_FIRST = 3
COL_LAST = 4
COL_DAY_START = 5
DAY_NAMES = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]

# Calm UI colors
COLOR_TEXT  = "#E0E0E6"
COLOR_HEADER= "#A3BE8C"

# Excel cell fills
FILL_DOT        = PatternFill(start_color="C5E1A5", end_color="C5E1A5", fill_type="solid")
FILL_DOT_HR     = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")
FILL_HELPER     = PatternFill(start_color="81D4FA", end_color="81D4FA", fill_type="solid")
FILL_XL         = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
FILL_STANDBY    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

ASSIGN_LABEL_TO_TEXT = {
    "DOT": "DOT Route",
    "DOT-HelperRoute": "DOT-HelperRoute",
    "DOT-Helper": "DOT-Helper",
    "XL": "XL",
    "Standby": "Standby"
}
ASSIGN_LABEL_TO_FILL = {
    "DOT": FILL_DOT,
    "DOT-HelperRoute": FILL_DOT_HR,
    "DOT-Helper": FILL_HELPER,
    "XL": FILL_XL,
    "Standby": FILL_STANDBY
}
ASSIGN_PRIORITY = ["DOT-HelperRoute","DOT","DOT-Helper","XL","Standby"]

# =========================
# 🗂️ UTILITIES
# =========================

def safe_sheet_name(name):
    invalid = '[]:*?/\\'
    return ''.join('_' if c in invalid else c for c in name)[:31]

def week_start_for(d:date):
    return d - timedelta(days=(d.weekday()+1)%7)

def read_schedule_df(path):
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet = xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet, header=None,
                         skiprows=SKIPROWS, nrows=NROWS, engine="openpyxl"), sheet

def save_df_to_sheet(workbook_path, df, sheet_name):
    book = load_workbook(workbook_path)
    if sheet_name in book.sheetnames:
        del book[sheet_name]; book.save(workbook_path)
    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="a") as w:
        df.to_excel(w, sheet_name=sheet_name, index=False)

def safe_save_workbook(wb, path):
    try:
        wb.save(path)
        return True
    except PermissionError:
        console.print(f"[bold red]❌ Close the Excel file before running the program again.[/bold red]")
        return False

# =========================
# 🧩 DRIVER LOGIC
# =========================
def build_driver_records(df):
    drivers=[]
    for _, row in df.iterrows():
        first=str(row[COL_FIRST]).strip() if not pd.isna(row[COL_FIRST]) else ""
        last=str(row[COL_LAST]).strip() if not pd.isna(row[COL_LAST]) else ""
        if not first and not last: continue
        days={DAY_NAMES[i]:row[COL_DAY_START+i] for i in range(len(DAY_NAMES))}
        drivers.append({"name":f"{first} {last}".strip(),"days":days})
    return drivers

def infer_dot_cert(driver):
    for v in driver["days"].values():
        if pd.isna(v): continue
        if str(v).strip().lower()=="dot": return True
    return False

def scheduled_on_day(driver,day):
    v=driver["days"].get(day)
    if pd.isna(v): return False
    s=str(v).strip().lower()
    return s in ("1","dot")

def prioritize_and_sample(pool,need,priority=None):
    if need<=0: return []
    chosen=[]
    if priority:
        pri = [n for n in priority if n in pool]
        random.shuffle(pri)
        take=min(len(pri),need)
        chosen.extend(pri[:take])
    remaining=need-len(chosen)
    if remaining>0:
        rest=[n for n in pool if n not in chosen]
        random.shuffle(rest)
        chosen.extend(rest[:remaining])
    return chosen

# =========================
# 🔁 GENERATION LOGIC
# =========================
def generate_day_assignments(drivers, dot_map, new_list, semi, counts, target_date,
                             dot_weekly_count, dot_stepvan_count, standby_tracker):
    day_idx=(target_date.weekday()+1)%7
    day_name=DAY_NAMES[day_idx]
    scheduled=[d["name"] for d in drivers if scheduled_on_day(d,day_name)]
    new_sched=[n for n in new_list if n in scheduled]

    all_avail=list(dict.fromkeys(scheduled+new_sched))
    dot_avail=[n for n in all_avail if dot_map.get(n,False) and n not in semi]

    # --- New Drivers get XL ---
    xl_assigned=[]
    need_xl=counts.get("xl",0)
    take_new=min(len(new_sched),need_xl)
    if take_new>0:
        random.shuffle(new_sched)
        xl_assigned.extend(new_sched[:take_new])
        need_xl-=take_new
    unassigned_new=[n for n in new_sched if n not in xl_assigned]

    # --- DOT step-van logic ---
    eligible_dot=[n for n in dot_avail if dot_weekly_count.get(n,0)<2]
    zero_step=[n for n in eligible_dot if dot_stepvan_count.get(n,0)==0]
    one_step=[n for n in eligible_dot if dot_stepvan_count.get(n,0)==1]
    dot_priority=zero_step+one_step

    dot_assigned=prioritize_and_sample(eligible_dot,counts.get("dot",0),priority=dot_priority)
    remaining_dot=[n for n in eligible_dot if n not in dot_assigned]
    zero_step2=[n for n in remaining_dot if dot_stepvan_count.get(n,0)==0]
    one_step2=[n for n in remaining_dot if dot_stepvan_count.get(n,0)==1]
    helperroute_assigned=prioritize_and_sample(remaining_dot,counts.get("dot_helperroute",0),
                                               priority=zero_step2+one_step2)

    helper_pool=[n for n in scheduled if n not in dot_assigned+helperroute_assigned+xl_assigned]
    helper_assigned=prioritize_and_sample(helper_pool,counts.get("dot_helper",0))

    for n in dot_assigned+helperroute_assigned:
        if dot_map.get(n,False): dot_weekly_count[n]+=1; dot_stepvan_count[n]+=1
    for n in helper_assigned:
        if dot_map.get(n,False): dot_weekly_count[n]+=1

    # --- Remaining XL ---
    if need_xl>0:
        remain=[n for n in scheduled if n not in dot_assigned+helperroute_assigned+helper_assigned+xl_assigned]
        random.shuffle(remain)
        xl_assigned.extend(remain[:need_xl])

    # --- Standby capped at 2 per week ---
    assigned_now=set(dot_assigned+helperroute_assigned+helper_assigned+xl_assigned)
    standby_candidates=[n for n in scheduled if n not in assigned_now]
    standby=[]
    for n in standby_candidates:
        if standby_tracker.get(n,0)<2:
            standby.append(n)
            standby_tracker[n]=standby_tracker.get(n,0)+1

    # --- Pair DOT-HelperRoute with helpers ---
    pairs=[]
    helpers_for_pair=helper_assigned.copy()
    random.shuffle(helpers_for_pair)
    for d in helperroute_assigned:
        h=helpers_for_pair.pop(0) if helpers_for_pair else None
        pairs.append((d,h))

    return {
        "DOT": dot_assigned,
        "DOT-HelperRoute": helperroute_assigned,
        "DOT-Helper": helper_assigned,
        "XL": xl_assigned,
        "Standby": standby,
        "Unassigned_New": unassigned_new,
        "Pairings": pairs
    }, dot_weekly_count, standby_tracker

# =========================
# 🖋️ SHEET UPDATES
# =========================
def build_day_label_map(assignments):
    label_map={}
    for d,h in assignments.get("Pairings",[]):
        if d: label_map[d]="DOT-HelperRoute"
    for label in ASSIGN_PRIORITY:
        if label=="DOT-HelperRoute": continue
        for n in assignments.get(label,[]):
            if n not in label_map:
                label_map[n]=label
    for n in assignments.get("Unassigned_New",[]):
        label_map[n]="Unassigned (Need XL)"
    return label_map

def update_schedule_with_assignments(source_path,target_path,day_maps):
    wb=load_workbook(source_path)
    ws=wb[wb.sheetnames[0]]
    for r in range(14,91):
        first=ws.cell(r,4).value; last=ws.cell(r,5).value
        if not first and not last: continue
        name=f"{str(first).strip()} {str(last).strip()}".strip()
        for day_idx,c in enumerate(range(6,13)):
            cell=ws.cell(r,c); v=cell.value
            s=str(v).strip().lower() if v else ""
            if s in ("1","dot"):
                label=day_maps.get(day_idx,{}).get(name)
                if label:
                    cell.value=ASSIGN_LABEL_TO_TEXT.get(label,label)
                    fill=ASSIGN_LABEL_TO_FILL.get(label)
                    if fill: cell.fill=fill
    safe_save_workbook(wb,target_path)

# =========================
# 🚀 MAIN
# =========================
def main():
    console.clear()
    banner = (
        f"🌙 {'━'*45} 🌙\n"
        f"      SMKL Dispatch Assistant\n"
        f"   “Plan smart. Drive safe. Rest easy.”\n"
        f"   This program is designed by Timmy Nguyen 😎\n"
        f"{'━'*55}"
    )
    console.print(Markdown(f"```{banner}```"))

    console.print("\nSelect your weekly Excel schedule file (rows 14–90, cols D–L):",style=COLOR_TEXT)
    uploaded_file = st.file_uploader("📘 Upload your weekly Excel schedule (rows 14–90, cols D–L):", type=["xlsx", "xls"])
	
	if uploaded_file is None:
    st.warning("Please upload a file to continue.")
    st.stop()

    path = uploaded_file

    if not path:
        console.print("[bold red]No file selected. Exiting.[/bold red]"); input(); return

    df_raw = pd.read_excel(uploaded_file, header=None, skiprows=SKIPROWS, nrows=NROWS, engine="openpyxl")
    drivers=build_driver_records(df_raw)
    dot_map={d["name"]:infer_dot_cert(d) for d in drivers}

    console.print("\nPaste new drivers (one per line). Press Enter immediately if none:",style=COLOR_TEXT)
    new_list=[]
    while True:
        ln=input().strip()
        if not ln: break
        new_list.append(ln)

    console.print("\n(Optional) Paste semi-restricted drivers (cannot do DOT/HelperRoute). Press Enter if none:",style=COLOR_TEXT)
    semi=[]
    while True:
        ln=input().strip()
        if not ln: break
        semi.append(ln.strip())

    console.print("\nEnter the week number (e.g., 45 for Week 45):", style=COLOR_TEXT)
    try:
        week_input=int(input().strip())
        today_year=datetime.today().year
        wk_start=datetime.fromisocalendar(today_year,week_input,1).date()-timedelta(days=1)
        console.print(f"[{COLOR_HEADER}]Week {week_input} detected: starting Sunday {wk_start}[/]")
    except Exception:
        console.print("[yellow]Invalid or blank input — defaulting to current week.[/yellow]")
        wk_start=week_start_for(date.today())

    week_dates=[wk_start+timedelta(days=i) for i in range(7)]
    day_label_maps={}
    dot_weekly_count={n:0 for n,is_dot in dot_map.items() if is_dot}
    dot_stepvan_count={n:0 for n,is_dot in dot_map.items() if is_dot}
    standby_tracker={}

    console.print(f"\n[{COLOR_HEADER}]Starting full week generation...[/]")

    for tgt in week_dates:
        day_label=tgt.strftime("%a %m_%d")
        console.print(f"\n[{COLOR_TEXT}]Generating assignments for [bold]{day_label}[/bold] ({tgt.isoformat()})[/{COLOR_TEXT}]")

        try:
            dot_hr=int(input("Number of DOT-HelperRoute: ").strip())
            dot_h=int(input("Number of DOT-Helpers: ").strip())
            dot_r=int(input("Number of DOT routes: ").strip())
            xl=int(input("Number of XL routes: ").strip())
        except Exception:
            console.print("[bold red]Invalid input. Skipping this day.[/bold red]"); continue

        counts={"dot_helperroute":dot_hr,"dot_helper":dot_h,"dot":dot_r,"xl":xl}
        assignments, dot_weekly_count, standby_tracker=generate_day_assignments(
            drivers,dot_map,new_list,semi,counts,tgt,dot_weekly_count,dot_stepvan_count,standby_tracker
        )
        day_idx=(tgt.weekday()+1)%7
        day_label_maps[day_idx]=build_day_label_map(assignments)

        def panel(title,names,emoji):
            lines="\n".join(f"- {x}" for x in names) if names else "[dim](none)[/dim]"
            return Panel(lines,title=f"{emoji} {title} ({len(names)})",expand=False)
        console.print(panel("DOT Routes",assignments["DOT"],"🚛"))
        console.print(panel("DOT-HelperRoute",assignments["DOT-HelperRoute"],"🚐"))
        console.print(panel("DOT-Helper",assignments["DOT-Helper"],"🧑‍🤝‍🧑"))
        console.print(panel("XL Routes",assignments["XL"],"📦"))
        console.print(panel("Standby",assignments["Standby"],"💤"))
        if assignments["Unassigned_New"]:
            console.print(f"[yellow]⚠️ Unassigned new drivers (need XL): {', '.join(assignments['Unassigned_New'])}[/yellow]")

        scheduled_count=len([d["name"] for d in drivers if scheduled_on_day(d,DAY_NAMES[(tgt.weekday()+1)%7])])
        route_count=len(assignments["DOT"])+len(assignments["DOT-HelperRoute"])+len(assignments["XL"])
        console.print(f"[{COLOR_HEADER}]📊 {day_label}: {route_count} routes assigned / {scheduled_count} scheduled drivers[/]")

        rows=[{"Group":f"Routes assigned: {route_count}","Driver":""},
              {"Group":f"Scheduled drivers: {scheduled_count}","Driver":""},
              {"Group":"","Driver":""}]
        for grp in ["DOT","DOT-HelperRoute","DOT-Helper","XL","Standby"]:
            rows.append({"Group":f"{grp} ({len(assignments[grp])})","Driver":""})
            for n in assignments[grp]: rows.append({"Group":"","Driver":n})
            rows.append({"Group":"","Driver":""})
        if assignments["Unassigned_New"]:
            rows.append({"Group":"Unassigned (Need XL)","Driver":""})
            for n in assignments["Unassigned_New"]:
                rows.append({"Group":"","Driver":n})
        df=pd.DataFrame(rows)
        save_df_to_sheet(path,df,safe_sheet_name(day_label))
        console.print("✅ Daily sheets updated")

        # --- Fairness Enforcement & Console Logs ---
    console.print(f"\n[{COLOR_HEADER}]⚖️ Starting fairness enforcement...[/]")

    # Track DOT fairness and standby counts
    unassigned_dots = [n for n, v in dot_stepvan_count.items() if v == 0]
    standby_count = {name: sum(1 for dm in day_label_maps.values() if dm.get(name) == "Standby")
                     for name in dot_map.keys()}

    if unassigned_dots:
        console.print(f"[yellow]⚠️ {len(unassigned_dots)} DOT-certified drivers did not drive a step van this week. Rebalancing...[/yellow]")
        swaps = []
        reassignments = []

        for driver in unassigned_dots:
            for day_idx, date_obj in enumerate(week_dates):
                dname = DAY_NAMES[(date_obj.weekday() + 1) % 7]
                try:
                    driver_data = next(d for d in drivers if d["name"] == driver)
                except StopIteration:
                    continue
                if scheduled_on_day(driver_data, dname):
                    replacements = day_label_maps[day_idx]
                    replaced = False
                    for group in ["DOT", "DOT-HelperRoute"]:
                        for assigned_driver, label in list(replacements.items()):
                            if label in ["DOT", "DOT-HelperRoute"]:
                                # Swap drivers
                                replacements.pop(assigned_driver)
                                replacements[driver] = group
                                swaps.append((assigned_driver, driver, dname))
                                dot_stepvan_count[driver] = 1

                                # Reassign replaced driver (if scheduled that day)
                                if scheduled_on_day(driver_data, dname):
                                    if standby_count.get(assigned_driver, 0) < 2:
                                        replacements[assigned_driver] = "Standby"
                                        standby_count[assigned_driver] = standby_count.get(assigned_driver, 0) + 1
                                        reassignments.append((assigned_driver, dname, "Standby"))
                                    else:
                                        # Skip if already at max standby
                                        console.print(f"[dim]- {assigned_driver} already at Standby cap (2), skipped reassign[/dim]")
                                replaced = True
                                break
                        if replaced:
                            break
                    if replaced:
                        break

        # Console logs
        for a, b, d in swaps:
            console.print(f"[green]→ Swapped {a} with {b} on {d}[/green]")
        for n, d, grp in reassignments:
            console.print(f"[cyan]↪ Reassigned {n} to {grp} on {d}[/cyan]")

        console.print(f"[green]✅ Step-van fairness applied! ({len(swaps)} swaps, {len(reassignments)} reassignments)[/green]")
    else:
        console.print(f"[green]✅ Step-van fairness already satisfied![/green]")

    console.print(f"[{COLOR_HEADER}]✅ Standby cap pass complete![/]")

    # --- Fairness_Audit summary ---
    console.print(f"[{COLOR_HEADER}]🧾 Adding Fairness_Audit sheet...[/]")
    audit_rows = []
    for day_idx, date_obj in enumerate(week_dates):
        day_label = date_obj.strftime("%a %m_%d")
        replacements = day_label_maps.get(day_idx, {})
        for name, label in replacements.items():
            if name in unassigned_dots and label in ["DOT", "DOT-HelperRoute"]:
                audit_rows.append({"Driver": name, "Assigned Label": label, "Day": day_label})
    if audit_rows:
        audit_df = pd.DataFrame(audit_rows)
        save_df_to_sheet(path, audit_df, safe_sheet_name("Fairness_Audit"))
        console.print(f"[green]✅ Fairness_Audit sheet added successfully![/green]")
    else:
        console.print(f"[yellow]ℹ️ No fairness swaps were recorded for audit.[/yellow]")

    # --- Refresh per-day sheets ---
    console.print(f"[{COLOR_HEADER}]🔁 Updating per-day sheets after swaps...[/]")
    for day_idx, date_obj in enumerate(week_dates):
        day_label = date_obj.strftime("%a %m_%d")
        assignments = {"DOT": [], "DOT-HelperRoute": [], "DOT-Helper": [], "XL": [], "Standby": []}
        for name, label in day_label_maps.get(day_idx, {}).items():
            if label in assignments:
                assignments[label].append(name)

        route_count = len(assignments["DOT"]) + len(assignments["DOT-HelperRoute"]) + len(assignments["XL"])
        scheduled_count = len([d for d in drivers if scheduled_on_day(d, DAY_NAMES[day_idx])])
        rows = [
            {"Group": f"Routes assigned: {route_count}", "Driver": ""},
            {"Group": f"Scheduled drivers: {scheduled_count}", "Driver": ""},
            {"Group": "", "Driver": ""},
        ]
        for grp in ["DOT", "DOT-HelperRoute", "DOT-Helper", "XL", "Standby"]:
            rows.append({"Group": f"{grp} ({len(assignments[grp])})", "Driver": ""})
            for n in assignments[grp]:
                rows.append({"Group": "", "Driver": n})
            rows.append({"Group": "", "Driver": ""})
        df = pd.DataFrame(rows)
        save_df_to_sheet(path, df, safe_sheet_name(day_label))

    console.print(f"[green]✅ All per-day sheets refreshed and fairness adjustments applied![/green]")


    # --- Final workbook save and confirmation ---
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    base = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(os.path.dirname(path), f"{base}_updated_with_assignments_{ts}.xlsx")
    update_schedule_with_assignments(path, out, day_label_maps)

    # Add "Generated on" note at the top of the first sheet
    from openpyxl.utils import get_column_letter

    wb = load_workbook(out)
    ws = wb[wb.sheetnames[0]]
    ws.insert_rows(1)  # Add a new top row
    ws["A1"] = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} – Baseline Version"
    ws["A1"].font = ws["A1"].font.copy(bold=True)
    safe_save_workbook(wb, out)
    console.print(f"[green]📅 Added generation timestamp to '{out}'[/green]") 
	
    console.print(f"\n[{COLOR_HEADER}]✅ Weekly file saved successfully![/]")
    console.print("\nNice work, Timmy! Week assigned and ready. 🚚💪😊",style=COLOR_TEXT)
    input("\nPress Enter to exit...")

if __name__=="__main__":
    try:
        main()
    except Exception as e:
        import traceback
        console.print("[bold red]Unexpected error:[/bold red]",str(e))
        console.print(traceback.format_exc(),style="dim")
        input("\nPress Enter to exit...")


