import os
import re
import sys
import pandas as pd
from pathlib import Path
from datetime import date
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Set the constants for the age assessment in years
TOO_OLD = 2
OLD = 1.5


##### GUI TIME
#
# Root window
root = tk.Tk()
root.title('Select Fish-DB-file')
root.withdraw()

file_path = filedialog.askopenfilename()

# Read the fish DB export
try:
    if file_path[-3:] == 'tab':
        fdb = pd.read_csv(file_path, sep='\t')
    elif file_path[-3:] == 'csv':
        fdb = pd.read_csv(file_path)
except:
    messagebox.showerror("Error", "The chosen file format is not supported")

# Create headers if the composition of the things is reliable and exactly the same
ch_dob = 'DoB'
death = 'Dead'
ch_newdob = 'DoB new'
ch_age = 'Age assessment'
ch_gen = 'Generator'
ch_ct = 'Caretaker'
fdb.columns = [ch_ct, 'Stock number', ch_dob, death, 'Dead of Death', ch_gen, 'Genotype', 'Trash', 'Line ID', 'Tank']
fdb.ffill(inplace=True)

# Select only those fish that are alive, drop the column
fdb_alive = fdb[fdb['Dead'] == 'NO']
fdb_alive.reset_index(inplace=True)
fdb_alive = fdb_alive.copy()
# Create new column with Date of Birth corrected into format that is readable for Python
fdb_alive[ch_newdob] = pd.to_datetime(fdb_alive['DoB'], dayfirst=True)


# Calculate the date cut off date for old fish
def year_to_months(years):
    if years < 0:
        return 0
    else:
        return int(years * 12)

old = relativedelta(months=year_to_months(OLD))
date_old = pd.Timestamp(date.today() - old)
# Calculate the cut off date for the too old fish
too_old = relativedelta(months=year_to_months(TOO_OLD))
date_too_old = pd.Timestamp(date.today() - too_old)


def is_old(row):
    if row[ch_newdob] < date_too_old:
        val = 'Too old'
    elif date_too_old < row[ch_newdob] < date_old:
        val = 'Old'
    elif date_old < row[ch_newdob]:
        val = 'Fine'
    else:
        val = 'NaN'
    return val


fdb_alive[ch_age] = fdb_alive.apply(is_old, axis=1)

fdb_report = fdb_alive[fdb_alive[ch_age] != 'Fine']
fdb_export = fdb_report[[ch_ct, ch_gen, 'Tank', 'Line ID', 'Stock number', ch_dob, ch_age]]
fdb_export = fdb_export.copy()
fdb_export.sort_values(by=[ch_ct, 'Tank'], inplace=True)
fdb_export.reset_index(inplace=True, drop=True)

root = tk.Tk()
root.title('Select export location')
root.withdraw()

sve_path = filedialog.askdirectory(initialdir=file_path)
sve_path = Path(sve_path)
# fdb_report.to_csv(sve_path + os.path.sep + 'fishpatrol_detailed.csv')
fdb_export.to_csv(sve_path / 'fishpatol_list.csv')
fdb.to_csv(sve_path / 'fishdb_filled.csv')



## Excel export
def clean_for_excel(df):
    """
    Clean DataFrame for Excel export by removing illegal characters
    """
    df_clean = df.copy()

    # Define illegal characters for Excel (control characters)
    illegal_chars = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

    # Clean all string columns
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':  # String columns
            df_clean[col] = df_clean[col].astype(str).apply(
                lambda x: illegal_chars.sub('', x) if isinstance(x, str) else x
            )

    return df_clean

fdb_export_clean = clean_for_excel(fdb_export)

try:
    fdb_export_clean.to_excel(sve_path / 'fishpatrole_age_list.xlsx', index=False)
    print("Excel file exported successfully!")
except Exception as e:
    print(f"Excel export failed: {e}")
    # Fallback: try with even more aggressive cleaning
    try:
        # Replace any remaining problematic values with empty strings
        fdb_export_fallback = fdb_export_clean.copy()
        for col in fdb_export_fallback.columns:
            if fdb_export_fallback[col].dtype == 'object':
                fdb_export_fallback[col] = fdb_export_fallback[col].apply(
                    lambda x: '' if not str(x).isprintable() else x
                )
        fdb_export_fallback.to_excel(sve_path / 'fishpatrole_age_list.xlsx', index=False)
        print("Excel file exported successfully with fallback cleaning!")
    except Exception as e2:
        print(f"Excel export failed even with fallback cleaning: {e2}")
        messagebox.showerror("Export Error", f"Could not export to Excel: {e2}")

excel_file = sve_path / 'fishpatrole_age_list.xlsx'
try:
    os.system(f"open -a '/Applications/Microsoft Excel.app' '{excel_file}' ")
except:
    print('oops mistake happened')
    messagebox.showerror('ERROR',
                        f'FISHpy couldn\'t identify Excel on your machine. Please go to the export-folder ({excel_file}) and open the excel file')




# OLD export code 24.07.2025
""" 
fdb_export.to_excel(sve_path / 'fishpatrole_age_list.xlsx')
# fdb_export.to_html(sve_path + os.path.sep + 'table.html')


excel_file = sve_path / 'fishpatrole_age_list.xlsx'
try:
    os.system(f"open -a '/Applications/Microsoft Excel.app' '{excel_file}' ")
except:
    print('oops mistake happened')
    messagebox.ERROR('ERROR',
                     f'FISHpy couldn\'t identify Excel on your machine. Please go the the export-folder ({excel_file} and open the excel file')
"""
