import os
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

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


# Create headders if the composition of the things is reliable and exactly the same
ch_dob = 'DoB'
death = 'Dead'
ch_newdob = 'DoB new'
ch_age = 'Age assessment'
ch_gen = 'Generator'
ch_ct = 'Caretaker'
fdb.columns = [ch_ct, 'Stock number', ch_dob, death, 'Dead of Death', ch_gen, 'Genotype', 'Trash','Line ID', 'Tank']

# Select only those fish that are alive, drop the column
fdb_alive = fdb[fdb['Dead'] == 'NO']
fdb_alive.reset_index(inplace=True)
# Create new column with Date of Birth corrected into format that is readable for Python
fdb_alive[ch_newdob] = pd.to_datetime(fdb_alive['DoB'], dayfirst=True)

# Calculate the date cut off date for old fish
old = relativedelta(years=1, months=6)
date_old = date.today() - old
# Calculate the cut off date for the too old fish
too_old = relativedelta(years=2)
date_too_old = date.today() - too_old

def f(row):
    if row[ch_newdob] < date_too_old:
        val = 'Too old'
    elif date_too_old < row[ch_newdob] < date_old:
        val = 'Old'
    elif  date_old < row[ch_newdob]:
        val = 'Fine'
    else:
        val = 'NaN'
    return val

fdb_alive[ch_age] = fdb_alive.apply(f, axis=1)

fdb_report = fdb_alive[fdb_alive[ch_age] != 'Fine']
fdb_export = fdb_report[[ch_ct, ch_gen, 'Tank', 'Line ID', 'Stock number', ch_dob, ch_age]]
fdb_export.sort_values(by=[ch_ct, 'Tank'], inplace=True)
fdb_export.reset_index(inplace=True, drop=True)


root = tk.Tk()
root.title('Select export location')
root.withdraw()

sve_path = filedialog.askdirectory(initialdir=file_path)

#fdb_report.to_csv(sve_path + os.path.sep + 'fishpatrol_detailed.csv')
fdb_export.to_csv(sve_path + os.path.sep + 'fishpatol_list.csv')
fdb_export.to_excel(sve_path + os.path.sep + 'fishpatrole_age_list.xlsx')
#fdb_export.to_html(sve_path + os.path.sep + 'table.html')


# Does creates a plot in order to create an easy to read table
# plt.figure()
# cell_text = []
# for row in range(len(fdb_export)):
#     cell_text.append(fdb_export.iloc[row])
# fdb_export.to_dict
# plt.table(cellText=cell_text, colLabels=fdb_export.columns)
# plt.axis('off')
# #plt.show()
# plt.savefig(sve_path + os.path.sep + 'fishpatrole.pdf', Bbox='tight')

excel_file = sve_path + os.path.sep + 'fishpatrole_age_list.xlsx'
try:
    os.system(f"open -a '/Applications/Microsoft Excel.app' '{excel_file}' ")
except:
    print('oops mistake happened')
    messagebox.ERROR('ERROR', f'FISHpy couldn\'t identify Excel on your machine. Please go the the export-folder ({excel_file} and open the excel file')