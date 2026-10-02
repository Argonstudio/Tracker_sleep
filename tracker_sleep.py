# sleep_tracker.py
# Developer: Ivan Voitkov
# GitHub: https://github.com/Argonstudio/Tracker_sleep/
# MIT License
#
# Automatic sleep tracking based on activity logs (Chrome history and Windows power events).
# This script analyzes periods of inactivity to detect sleep periods, saves them to an Excel file,
# and generates a chart.

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import os
import msvcrt
from openpyxl.styles import Font, Border, Side, Alignment
import sqlite3
import shutil
import tempfile
import subprocess

# ---------- previous functions (manual input and formatting) ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(BASE_DIR, 'sleep_log.xlsx')
CHART_NAME = os.path.join(BASE_DIR, 'sleep_chart.png')

def smart_input_time(prompt):
    """Input time in HH:MM format"""
    print(f"\n[ {prompt} ]")
    template = "__:__"
    input_str = ""
    idx_map = [0, 1, 3, 4]
    while len(input_str) < 4:
        display = list(template)
        for i, char in enumerate(input_str):
            display[idx_map[i]] = char
        print(f"\r│ {''.join(display)}", end='', flush=True)
        char = msvcrt.getch()
        if char == b'\x08':            # Backspace
            input_str = input_str[:-1]
        elif char.isdigit():
            input_str += char.decode('utf-8')
    try:
        h, m = int(input_str[:2]), int(input_str[2:])
        if 0 <= h <= 23 and 0 <= m <= 59:
            time_str = f"{input_str[:2]}:{input_str[2:]}"
            print(f"\r│ {time_str} ✅")
            return time_str
        else:
            print(f"\r│ {input_str[:2]}:{input_str[2:]} ❌ Invalid time")
            return smart_input_time(prompt)
    except:
        print(f"\r│ {input_str[:2]}:{input_str[2:]} ❌ Error")
        return smart_input_time(prompt)

def smart_input_date(prompt):
    """Input date DD.MM.YYYY, if only day and month entered + Enter, year = current"""
    print(f"\n[ {prompt} ]")
    template = "__.__.____"
    curr_year = str(datetime.now().year)
    input_str = ""
    idx_map = [0, 1, 3, 4, 6, 7, 8, 9]
    while True:
        display = list(template)
        for i, char in enumerate(input_str):
            display[idx_map[i]] = char
        print(f"\r│ {''.join(display)}", end='', flush=True)
        char = msvcrt.getch()
        if char == b'\x08':
            input_str = input_str[:-1]
        elif char == b'\r':                     # Enter
            if len(input_str) == 4:
                input_str += curr_year
                break
            elif len(input_str) == 8:
                break
        elif char.isdigit() and len(input_str) < 8:
            input_str += char.decode('utf-8')
    date_str = f"{input_str[:2]}.{input_str[2:4]}.{input_str[4:]}"
    try:
        datetime.strptime(date_str, '%d.%m.%Y')
        print(f"\r│ {date_str} ✅")
        return date_str
    except:
        print(f"\r│ {date_str} ❌ Date error")
        return smart_input_date(prompt)

def format_excel(file_path):
    """Beautiful Excel formatting"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    header_font = Font(size=12, bold=True)
    bold_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    if 'Data' in wb.sheetnames:
        ws = wb['Data']
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=7):
            for cell in row:
                cell.border = thin_border
                if cell.column < 7:
                    cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['G'].width = 45

    if 'Analytics' in wb.sheetnames:
        ws_stat = wb['Analytics']
        ws_stat.column_dimensions['A'].width = 35
        for row in ws_stat.iter_rows():
            row[0].font = bold_font
            for cell in row:
                cell.border = thin_border
    wb.save(file_path)

def process_and_save(df):
    """Sorting, calculating metrics, statistics, Excel and chart"""
    df['Bedtime'] = pd.to_datetime(df['Bedtime'], dayfirst=True)
    df['Wake up'] = pd.to_datetime(df['Wake up'], dayfirst=True)
    df = df.sort_values(by='Bedtime').reset_index(drop=True)

    awakes, shifts, totals = [0.0]*len(df), [0.0]*len(df), [0.0]*len(df)
    for i in range(len(df)):
        sleep_dur = (df.loc[i, 'Wake up'] - df.loc[i, 'Bedtime']).total_seconds() / 3600
        df.loc[i, 'Sleep'] = round(sleep_dur, 2)
        if i > 0:
            p_wake = df.loc[i-1, 'Wake up']
            p_bed = df.loc[i-1, 'Bedtime']
            c_bed = df.loc[i, 'Bedtime']
            gap = (c_bed - p_wake).total_seconds() / 3600
            if 0 < gap < 200:
                awakes[i] = round(gap, 2)
                shifts[i] = round((c_bed - p_bed).total_seconds() / 3600 - 24, 2)
                totals[i] = round(sleep_dur + awakes[i], 2)

    df['Awake'] = awakes
    df['Shift'] = shifts
    df['Daily total'] = totals

    valid = df[df['Awake'] > 0]
    stats_data = [
        ("Average sleep", df['Sleep'].mean()),
        ("Minimum sleep", df['Sleep'].min()),
        ("Maximum sleep", df['Sleep'].max()),
        ("Average awake", valid['Awake'].mean() if not valid.empty else 0),
        ("Minimum awake", valid['Awake'].min() if not valid.empty else 0),
        ("Maximum awake", valid['Awake'].max() if not valid.empty else 0),
        ("Average daily total", valid['Daily total'].mean() if not valid.empty else 0),
        ("Minimum daily total", valid['Daily total'].min() if not valid.empty else 0),
        ("Maximum daily total", valid['Daily total'].max() if not valid.empty else 0),
        ("Minimum shift (earlier)", valid['Shift'].min() if not valid.empty else 0),
        ("Maximum shift (later)", valid['Shift'].max() if not valid.empty else 0),
        ("Total sleep periods", len(df)),
        ("Total sleep hours", df['Sleep'].sum())
    ]
    stats_df = pd.DataFrame(stats_data)
    stats_df[1] = stats_df[1].apply(lambda x: round(x, 1) if isinstance(x, (float, int)) else x)

    df_excel = df.copy()
    df_excel['Bedtime'] = df_excel['Bedtime'].dt.strftime('%d.%m.%Y %H:%M')
    df_excel['Wake up'] = df_excel['Wake up'].dt.strftime('%d.%m.%Y %H:%M')

    with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
        df_excel.to_excel(writer, sheet_name='Data', index=False)
        stats_df.to_excel(writer, sheet_name='Analytics', index=False, header=False)

    format_excel(FILE_NAME)

    plt.figure(figsize=(10, 5))
    plt.plot(df.index, df['Sleep'], 'b-o', label='Sleep')
    plt.plot(df.index, df['Awake'].replace(0, float('nan')), 'r-s', label='Awake')
    plt.title('Sleep and activity history')
    plt.legend()
    plt.grid(alpha=0.3)
    plt.savefig(CHART_NAME)
    plt.close()

# ---------- automatic data collection ----------
def get_chrome_visits():
    """Extracts visit times from Chrome history."""
    visits = []
    chrome_dir = os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\User Data')
    if not os.path.isdir(chrome_dir):
        return visits

    for profile in ['Default', 'Profile 1', 'Profile 2']:
        history_path = os.path.join(chrome_dir, profile, 'History')
        if not os.path.isfile(history_path):
            continue
        try:
            # Try to open directly
            conn = sqlite3.connect(history_path)
            _read_chrome_timestamps(conn, visits)
            conn.close()
        except Exception:
            # File is locked – copy to temporary
            try:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.sqlite')
                tmp.close()
                shutil.copy2(history_path, tmp.name)
                conn = sqlite3.connect(tmp.name)
                _read_chrome_timestamps(conn, visits)
                conn.close()
                os.unlink(tmp.name)
            except Exception:
                pass
    return visits

def _read_chrome_timestamps(conn, out_list):
    """Reads last_visit_time from open connection and appends datetime to out_list."""
    cursor = conn.cursor()
    cursor.execute("SELECT last_visit_time FROM urls WHERE last_visit_time > 0")
    for row in cursor:
        t = row[0]
        # Chrome time: microseconds since 1601-01-01 → Unix seconds → local time
        try:
            unix_ts = (t / 1_000_000) - 11644473600
            if unix_ts > 0:
                dt = datetime.fromtimestamp(unix_ts)
                out_list.append(dt)
        except:
            pass

def get_windows_power_events():
    """Gets power on/off event times from System log (ID 6005/6006)."""
    events = []
    try:
        cmd = [
            'powershell', '-Command',
            "Get-WinEvent -FilterHashtable @{LogName='System'; ID=6005,6006} -MaxEvents 500 "
            "| ForEach-Object { $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') }"
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                line = line.strip()
                if line:
                    try:
                        dt = datetime.strptime(line, '%Y-%m-%d %H:%M:%S')
                        events.append(dt)
                    except:
                        pass
    except Exception as e:
        print(f"Error getting Windows events: {e}")
    return events

def find_sleep_periods(activity_times, min_gap_h=3.0, max_gap_h=24.0):
    """
    Finds continuous periods without activity lasting
    from min_gap_h to max_gap_h hours and converts them to sleep periods.
    Returns a list of tuples (bed_time, wake_time).
    """
    if len(activity_times) < 2:
        return []

    # Remove duplicates and sort
    unique = sorted(set(activity_times))
    periods = []
    for i in range(len(unique) - 1):
        t1 = unique[i]
        t2 = unique[i + 1]
        gap_hours = (t2 - t1).total_seconds() / 3600.0
        if min_gap_h < gap_hours <= max_gap_h:
            bed = t1 + timedelta(minutes=30)      # time of falling asleep
            wake = t2 - timedelta(hours=1)        # time of waking up
            if bed < wake:
                periods.append((bed, wake))
    return periods

# ---------- main block ----------
def main():
    print("=== AUTOMATIC SLEEP TRACKING (based on activity logs) ===\n")

    # 1. Load existing log
    if os.path.exists(FILE_NAME):
        df = pd.read_excel(FILE_NAME, sheet_name='Data')
    else:
        df = pd.DataFrame(columns=[
            'Bedtime', 'Wake up', 'Sleep', 'Awake',
            'Shift', 'Daily total', 'Comment'
        ])

    # Determine date after which to look for new entries
    if df.empty:
        max_date = None
    else:
        # Convert to datetime for reliability
        df['Wake up_dt'] = pd.to_datetime(df['Wake up'], dayfirst=True)
        max_date = df['Wake up_dt'].max()
        df.drop(columns=['Wake up_dt'], inplace=True, errors='ignore')

    # 2. Collect activity events
    print("Collecting Chrome history...")
    chrome_events = get_chrome_visits()
    print(f"Found Chrome visits: {len(chrome_events)}")

    print("Getting Windows power on/off events...")
    power_events = get_windows_power_events()
    print(f"Found power events: {len(power_events)}")

    all_events = chrome_events + power_events
    if not all_events:
        print("\nNo activity data found. Program terminated.")
        input("Press Enter to exit...")
        return

    # 3. Find sleep periods
    sleep_periods = find_sleep_periods(all_events)
    print(f"Detected sleep periods (gap >3h and ≤16h): {len(sleep_periods)}")

    # 4. Filter new ones (after last record)
    if max_date is not None:
        new_periods = [p for p in sleep_periods if p[0] > max_date]
    else:
        new_periods = sleep_periods

    # 5. Output results
    # Saved log (last 5 entries)
    if not df.empty:
        print("\n1. Saved log")
        for _, row in df.tail(5).iterrows():
            bed_str = row['Bedtime'] if isinstance(row['Bedtime'], str) else pd.to_datetime(row['Bedtime']).strftime('%d.%m.%Y %H:%M')
            wake_str = row['Wake up'] if isinstance(row['Wake up'], str) else pd.to_datetime(row['Wake up']).strftime('%d.%m.%Y %H:%M')
            sleep_val = row['Sleep'] if not pd.isna(row['Sleep']) else 0
            awake_val = row['Awake'] if not pd.isna(row['Awake']) else 0
            print(f"{bed_str} - {wake_str}   sleep {round(sleep_val,2)} awake {round(awake_val,2)}")
    else:
        print("\n1. Saved log is empty")

    print("\n2. New dates.")
    if not new_periods:
        print("No new entries to add.")
        input("Press Enter to exit...")
        return

    prev_wake = max_date
    for i, (bed, wake) in enumerate(new_periods, 1):
        sleep_dur = round((wake - bed).total_seconds() / 3600, 2)
        if prev_wake is not None:
            awake_gap = round((bed - prev_wake).total_seconds() / 3600, 2)
            awake_str = f"awake {awake_gap}"
        else:
            awake_str = "awake -"
        print(f"{i}. {bed.strftime('%d.%m.%Y %H:%M')} - {wake.strftime('%d.%m.%Y %H:%M')}   sleep {sleep_dur} {awake_str}")
        prev_wake = wake

    # 6. Confirmation
    print("\nAre all dates correct? Save?\n1. Yes\n2. No")
    choice = input(">> ").strip()

    if choice == '2':
        try:
            nums = input("Enter the numbers of incorrect dates separated by spaces: ").strip()
            remove_indices = sorted([int(x) for x in nums.split()], reverse=True)
            for idx in remove_indices:
                if 1 <= idx <= len(new_periods):
                    del new_periods[idx - 1]
        except:
            print("Input error. New dates unchanged.")

    if new_periods:
        # Add remaining entries to the main DataFrame
        new_rows = []
        for bed, wake in new_periods:
            new_rows.append({
                'Bedtime': bed.strftime('%d.%m.%Y %H:%M'),
                'Wake up': wake.strftime('%d.%m.%Y %H:%M'),
                'Comment': ''
            })
        new_df = pd.DataFrame(new_rows)
        df = pd.concat([df, new_df], ignore_index=True)
        process_and_save(df)
        print("\n✅ Sleep log updated!")
    else:
        print("No new dates to save.")

    input("\nPress Enter to close the window...")

if __name__ == "__main__":
    main()
