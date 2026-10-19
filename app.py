
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import PyPDF2
import re
from datetime import datetime

def process_data(filepath, savepath, shift_config):
    staff_data = {} 
    
    if filepath.lower().endswith('.pdf'):
        with open(filepath, 'rb') as f:
            text = PyPDF2.PdfReader(f).pages[0].extract_text()
        tokens = text.split()
        datetimes = []
        for i in range(len(tokens)-1):
            if re.match(r'\d{2}-\d{2}-\d{4}', tokens[i]) and re.match(r'\d{2}:\d{2}', tokens[i+1]):
                try:
                    dt = pd.to_datetime(tokens[i] + " " + tokens[i+1], format='%d-%m-%Y %H:%M')
                    datetimes.append(dt)
                except: pass
        
        staff_names = ['Ankita', 'Deepak', 'Poonam', 'Jatan', 'Seema']
        for s in staff_names: staff_data[s] = []
        for i, dt in enumerate(datetimes):
            staff_data[staff_names[i % len(staff_names)]].append(dt)
    else:
        # Excel file upload handling
        df_raw = pd.read_excel(filepath)
        for col in df_raw.columns:
            staff_name = str(col).strip()
            if staff_name.lower().startswith('unnamed'): continue
            
            timestamps = []
            for val in df_raw[col].dropna():
                try:
                    # Parse as datetime
                    if isinstance(val, str):
                        dt = pd.to_datetime(val, dayfirst=True)
                    else:
                        dt = pd.to_datetime(val)
                    timestamps.append(dt)
                except: pass
            if len(timestamps) > 0:
                staff_data[staff_name] = timestamps

    if not staff_data:
        raise Exception("Attendance data read nahi ho paya. Kripya check karein ki Excel mein Staff Name columns hain aur unke niche Datetime likhe hain.")

    staff_dfs = {}
    summary_rows = []

    writer = pd.ExcelWriter(savepath, engine='openpyxl')

    for staff, timestamps in staff_data.items():
        # Default timings
        shift_start_str = '09:00:00'
        shift_end_str = '16:00:00'
        
        for cfg_staff, cfg_times in shift_config.items():
            if cfg_staff.lower() in staff.lower() or staff.lower() in cfg_staff.lower():
                shift_start_str = cfg_times['start'] + ':00'
                shift_end_str = cfg_times['end'] + ':00'
                break

        df = pd.DataFrame({'Timestamp': timestamps})
        df['Date'] = df['Timestamp'].dt.normalize()
        daily = df.groupby('Date')['Timestamp'].agg(['min', 'max']).reset_index()
        daily.rename(columns={'min': 'In Time', 'max': 'Out Time'}, inplace=True)
        
        daily.loc[daily['In Time'] == daily['Out Time'], 'Out Time'] = pd.NaT
        
        if not daily.empty:
            start_date = daily['Date'].min().replace(day=1)
            end_date = (start_date + pd.offsets.MonthEnd(1))
            full_month = pd.DataFrame({'Date': pd.date_range(start=start_date, end=end_date)})
        else:
            full_month = pd.DataFrame({'Date': pd.date_range(start='2026-09-01', end='2026-09-30')})
        
        merged = pd.merge(full_month, daily, on='Date', how='left')
        merged['Day'] = merged['Date'].dt.strftime('%a')
        
        merged['Shift Start'] = shift_start_str
        merged['Shift End'] = shift_end_str
        
        merged['In Time Str'] = merged['In Time'].dt.strftime('%H:%M:%S').fillna('-')
        merged['Out Time Str'] = merged['Out Time'].dt.strftime('%H:%M:%S').fillna('-')
        
        merged['Target In'] = pd.to_datetime(merged['Date'].dt.strftime('%Y-%m-%d') + ' ' + shift_start_str)
        merged['Target Out'] = pd.to_datetime(merged['Date'].dt.strftime('%Y-%m-%d') + ' ' + shift_end_str)
        
        # Late Calculation
        merged['Late (min)'] = (merged['In Time'] - merged['Target In']).dt.total_seconds() / 60
        merged['Late (min)'] = merged['Late (min)'].apply(lambda x: x if pd.notna(x) and x > 0 else 0).astype(int)
        
        # Early Going Calculation
        merged['Early Going (min)'] = (merged['Target Out'] - merged['Out Time']).dt.total_seconds() / 60
        merged['Early Going (min)'] = merged['Early Going (min)'].apply(lambda x: x if pd.notna(x) and x > 0 else 0).astype(int)
        
        # Extra Work Calculation
        merged['Extra Work (min)'] = (merged['Out Time'] - merged['Target Out']).dt.total_seconds() / 60
        merged['Extra Work (min)'] = merged['Extra Work (min)'].apply(lambda x: x if pd.notna(x) and x > 0 else 0).astype(int)
        
        merged['Status'] = np.where(merged['In Time'].notna(), 'Present', 'Absent')
        
        output_df = merged[['Date', 'Day', 'Shift Start', 'Shift End', 'In Time Str', 'Out Time Str', 'Late (min)', 'Early Going (min)', 'Extra Work (min)', 'Status']].copy()
        output_df['Date'] = output_df['Date'].dt.strftime('%Y-%m-%d')
        
        staff_dfs[staff] = output_df
        
        total_days = len(output_df)
        present = (output_df['Status'] == 'Present').sum()
        absent = total_days - present
        late_days = (output_df['Late (min)'] > 0).sum()
        early_days = (output_df['Early Going (min)'] > 0).sum()
        extra_days = (output_df['Extra Work (min)'] > 0).sum()
        
        summary_rows.append({
            'Staff Name': staff,
            'Shift Timing': f"{shift_start_str[:5]} to {shift_end_str[:5]}",
            'Total Days': total_days,
            'Present Days': present,
            'Absent Days': absent,
            'Days Late': late_days,
            'Total Late (min)': output_df['Late (min)'].sum(),
            'Days Early Left': early_days,
            'Total Early (min)': output_df['Early Going (min)'].sum(),
            'Days Extra Work': extra_days,
            'Total Extra (min)': output_df['Extra Work (min)'].sum()
        })

    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    for staff, df in staff_dfs.items():
        safe_name = str(staff)[:31].replace(':', '').replace('/', '')
        df.to_excel(writer, sheet_name=safe_name, index=False)

    writer.close()

def start_processing():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("PDF Files", "*.pdf")])
    if not filepath: return
    
    savepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not savepath: return
    
    config_text = txt_shifts.get("1.0", tk.END).strip().split('\n')
    shift_config = {}
    for line in config_text:
        if ',' in line:
            parts = line.split(',')
            if len(parts) >= 3:
                staff = parts[0].strip()
                start = parts[1].strip()
                end = parts[2].strip()
                shift_config[staff] = {'start': start, 'end': end}
    try:
        process_data(filepath, savepath, shift_config)
        messagebox.showinfo("Success", f"Attendance Report Ban Gayi Hai!\n\nSaved at:\n{savepath}")
    except Exception as e:
        messagebox.showerror("Error", f"Koi gadbad ho gayi:\n{str(e)}")

root = tk.Tk()
root.title("RL Hospital - Advanced Attendance Analyzer PRO")
root.geometry("600x500")
root.configure(bg="#f4f4f4")

header = tk.Label(root, text="RL Hospital Attendance Analyzer PRO", font=("Arial", 16, "bold"), bg="#f4f4f4", fg="#333")
header.pack(pady=10)

frame1 = tk.Frame(root, bg="#f4f4f4")
frame1.pack(pady=10, padx=20, fill="both")

tk.Label(frame1, text="Step 1: Har Staff Ka Shift Timing Set Karein", font=("Arial", 11, "bold"), bg="#f4f4f4").pack(anchor="w")
tk.Label(frame1, text="Format (Naam, Aane ka time, Jane ka time) 24-hr format mein likhein:\nJaise: Ankita, 09:00, 16:00", font=("Arial", 9), bg="#f4f4f4", justify="left").pack(anchor="w", pady=5)

txt_shifts = scrolledtext.ScrolledText(frame1, width=60, height=8, font=("Consolas", 10))
txt_shifts.pack(pady=5)
default_text = "Ankita, 09:00, 16:00\nDeepak, 09:00, 16:00\nPoonam, 09:00, 16:00\nJatan, 09:00, 16:00\nSeema, 09:00, 16:00"
txt_shifts.insert(tk.END, default_text)

frame2 = tk.Frame(root, bg="#f4f4f4")
frame2.pack(pady=20)

tk.Label(frame2, text="Step 2: Apni Raw Excel File (ya PDF) Upload Karein", font=("Arial", 11, "bold"), bg="#f4f4f4").pack()
tk.Button(frame2, text="📂 Upload Excel & Generate Report", command=start_processing, bg="#0078D7", fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, relief="raised").pack(pady=10)

root.mainloop()
