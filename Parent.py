import win32com.client as win32
import time
from datetime import datetime
import optuna
import os
import shutil
import csv
import pandas as pd
from optuna.samplers import NSGAIISampler
import sys
import subprocess

# Clean up from previous runs
os.system("TASKKILL /F /IM excel.exe")
time.sleep(5)

# Get current file locations
dest_folder = os.getcwd()

# Settings
# Number of workers on this pc
cpu_num=1

# Excel file locations and destinations
file_path = r'C:\Database\James\ValeSA_HoV_V1.20_base_Met.xlsm'
_name=file_path.split('\\')[-1].split('.xl')[0]
base_name ="node"

# Variable storage location
csv_path = r'C:\Database\James\vars.csv'

# Script with objective function
script_relative_path = "Optuna_/nodex.py"
script_to_execute=os.path.join(dest_folder,script_relative_path)

# Storage url
journal_path=r'C:\Database\James\journal.log'
jounal_lock=r'C:\Database\James\journal.log.lock'

# Excel mappings
sheet_name="Solver_Stack"
range_address="F13:F63"
target="J13"

# Optimisation settings
max_time=14400
max_iter = 20000

def get_params(file_path, sheet_name, range_address):

    excel = win32.DispatchEx('Excel.Application')
    workbook = excel.Workbooks.Open(file_path)
    excel.Visible = False  # Keep Excel hidden  

    # Reference the specified sheet and range
    worksheet = workbook.Sheets(sheet_name)
    variable_range = worksheet.Range(range_address)

    # Read the number of variables based on the number of rows in the range
    num_vars = variable_range.Rows.Count

    # Read the lower bounds from Excel (assuming they are in column -3)
    lower_bounds_range = variable_range.GetOffset(0,-3)
    lower_bounds_list = [int(cell.Value) for cell in lower_bounds_range]

    # Read the upper bounds from Excel (assuming they are in column -2)
    upper_bounds_range = variable_range.GetOffset(0, -2)
    upper_bounds_list = [int(cell.Value) for cell in upper_bounds_range]

    # Read the upper bounds from Excel (assuming they are in column -1)
    guess_range = variable_range.GetOffset(0, -1)
    guess_list = [int(cell.Value) for cell in guess_range]

    # Set calculation mode to manual
    excel.Application.Calculation = -4135

    # Save the workbook (optional)
    workbook.Save()
    excel.DisplayAlerts=False
    workbook.Close()
    excel.Quit()

    return(num_vars,lower_bounds_list,upper_bounds_list,guess_list)
num_vars,lower_bounds_list,upper_bounds_list,guess_list = get_params(file_path, sheet_name, range_address)

def copy_excel_files(file_path, num_copies, base_name, dest_folder):
    # Copy the source file to the destination folder
    for i in range(num_copies):
        dest_file = os.path.join(dest_folder, f'{base_name}_{i}')
        shutil.copy(file_path, dest_file)
copy_excel_files(file_path, cpu_num, base_name, dest_folder)

# Save parameters in a .csv for use in subprocess
prompts=[sheet_name,range_address,target,max_time,lower_bounds_list,upper_bounds_list,guess_list,num_vars,journal_path, max_iter]
with open(csv_path, 'w') as file:
    csv.writer(file).writerow(prompts)
  
try: 
    os.remove(journal_path)
except OSError:
    pass

try: 
    os.remove(jounal_lock)
except OSError:
    pass
    
lock_obj = optuna.storages.JournalFileOpenLock(journal_path)
storage = optuna.storages.JournalStorage(
optuna.storages.JournalFileStorage(journal_path, lock_obj=lock_obj),
)

# Create a study
study = optuna.create_study(study_name="multi_cpu", sampler=NSGAIISampler(), direction="maximize", storage=storage, load_if_exists=True)

procs = [subprocess.Popen([sys.executable, script_to_execute, str(i), csv_path]) for i in range(cpu_num)]

for p in procs:
    p.wait()

# Close all excel instances
os.system("TASKKILL /F /IM excel.exe")

study = optuna.load_study(
    study_name="multi_cpu", storage=storage
)

# Save results
timestamp=datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
f_name=f"{_name}_results_{timestamp}.csv"
df = study.trials_dataframe()
df.to_csv(f_name, index=False)

print(study.best_params)
print(study.best_value)

