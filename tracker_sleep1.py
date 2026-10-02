# sleep_tracker.py
# Developer: Ivan Voitkov
# GitHub: https://github.com/Argonstudio/Tracker_sleep/
# MIT License
#
# A console application for manual sleep tracking.
# It allows entering bedtime and wake-up time, calculates sleep duration,
# updates statistics, saves data to an Excel file, and generates a chart.

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import msvcrt
from openpyxl.styles import Font, Border, Side, Alignment

# ---------- Path settings ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(BASE_DIR, 'sleep_log.xlsx')
CHART_NAME = os.path.join(BASE_DIR, 'sleep_chart.png')

def smart_input(prompt):
    """
    Interactive input of date and time in the format DD.MM.YYYY HH:MM.
    Uses msvcrt for character-by-character input with a template.
    """
    print(f"\n[ {prompt} ]")
    template = "__.__.____ __:__"
    input_str = ""
    # Positions in the template where digits should be placed
    idx_map = [0, 1, 3, 4, 6, 7, 8, 9, 11, 12, 14, 15]
    
    while len(input_str) < 12:
        display = list(template)
        for i, char in enumerate(input_str):
            display[idx_map[i]] = char
        print(f"\r│ {''.join(display)}", end='', flush=True)
        
        char = msvcrt.getch()
        if char == b'\x08':  # Backspace
            input_str = input_str[:-1]
        elif char.isdigit():
            input_str += char.decode('utf-8')
            
    final_date = f"{input_str[:2]}.{input_str[2:4]}.{input_str[4:8]} {input_str[8:10]}:{input_str[10:]}"
    print(f"\r│ {final_date} ✅")
    return datetime.strptime(final_date, '%d.%m.%Y %H:%M')

def format_excel(file_path):
    """
    Applies formatting to the Excel file: fonts, borders, column widths.
    """
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    header_font = Font(size=12, bold=True)
    bold_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    ws = wb['Data']
    # Format header row
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
    
    # Apply borders and center alignment to all cells except the last column
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.border = thin_border
            if cell.column < 7:
                cell.alignment = Alignment(horizontal='center')

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['G'].width = 45

    ws_stat = wb['Analytics']
    ws_stat.column_dimensions['A'].width = 35
    for row in ws_stat.iter_rows():
        row[0].font = bold_font
        for cell in row:
            cell.border = thin_border
    wb.save(file_path)

def process_and_save(df):
    """
    Sorts data, recalculates awake time, shift, and daily total,
    saves everything to an Excel file, and updates the chart.
    """
    # 1. Sorting and recalculation
    df['Bedtime'] = pd.to_datetime(df['Bedtime'], dayfirst=True)
    df['Wake up'] = pd.to_datetime(df['Wake up'], dayfirst=True)
    df = df.sort_values(by='Bedtime').reset_index(drop=True)

    awakes, shifts, totals = [0.0]*len(df), [0.0]*len(df), [0.0]*len(df)
    for i in range(1, len(df)):
        p_wake, p_bed = df.loc[i-1, 'Wake up'], df.loc[i-1, 'Bedtime']
        c_bed, c_sleep = df.loc[i, 'Bedtime'], df.loc[i, 'Sleep']
        gap = (c_bed - p_wake).total_seconds() / 3600
        if 0 < gap < 48:
            awakes[i] = round(gap, 2)
            shifts[i] = round((c_bed - p_bed).total_seconds() / 3600 - 24, 2)
            totals[i] = round(c_sleep + awakes[i], 2)

    df['Awake'], df['Shift'], df['Daily total'] = awakes, shifts, totals
    df['Bedtime'] = df['Bedtime'].dt.strftime('%d.%m.%Y %H:%M')
    df['Wake up'] = df['Wake up'].dt.strftime('%d.%m.%Y %H:%M')

    # 2. Extended statistics
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
    # Round all numeric values to 1 decimal place
    stats_df[1] = stats_df[1].apply(lambda x: round(x, 1) if isinstance(x, (float, int)) else x)

    # 3. Writing to Excel
    with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        stats_df.to_excel(writer, sheet_name='Analytics', index=False, header=False)
    
    format_excel(FILE_NAME)

    # 4. Generating the chart
    plt.figure(figsize=(10, 5))
    plt.plot(df.index, df['Sleep'], 'b-o', label='Sleep')
    plt.plot(df.index, df['Awake'].replace(0, float('nan')), 'r-s', label='Awake')
    plt.axhline(0, color='gray', linestyle='--', alpha=0.5)
    plt.title('Sleep chart')
    plt.legend()
    plt.grid(alpha=0.3)
    plt.savefig(CHART_NAME)
    plt.close()

def main():
    """
    Main loop: prompts user for bedtime, wake-up time, and comment,
    then updates the log and statistics.
    """
    print("=== SLEEP TRACKING (Close window to exit) ===")
    while True:
        try:
            # Load existing data or create an empty DataFrame
            if os.path.exists(FILE_NAME):
                df = pd.read_excel(FILE_NAME, sheet_name='Data')
            else:
                df = pd.DataFrame(columns=['Bedtime', 'Wake up', 'Sleep', 'Awake', 'Shift', 'Daily total', 'Comment'])

            bed_t = smart_input("When did you go to bed?")
            wake_t = smart_input("When did you wake up?")
            comment = input("\n[ Comment ] >> ")
            
            sleep_dur = round((wake_t - bed_t).total_seconds() / 3600, 2)
            new_row = {
                'Bedtime': bed_t.strftime('%d.%m.%Y %H:%M'), 
                'Wake up': wake_t.strftime('%d.%m.%Y %H:%M'), 
                'Sleep': sleep_dur, 
                'Comment': comment
            }
            
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            process_and_save(df)
            print("\n✅ Saved! Press Enter for a new entry or close the window.")
            input()

        except Exception as e:
            print(f"\n🚨 Error: {e}")
            input("Press Enter to try again...")

if __name__ == "__main__":
    main()
