import win32com.client as win32
import optuna
import sys
import os
import csv
import ast

node = int(sys.argv[1])
csv_path = sys.argv[2]

dest_folder = os.getcwd()
base_name = "node"

with open(csv_path, 'r') as file:
    reader=csv.reader(file)
    prompts=[]
    for row in reader:
        prompts.append(row)

    sheet_name=prompts[0][0]
    range_address=prompts[0][1] 
    target=prompts[0][2]
    max_time=int(prompts[0][3])
    lower_bounds_list= ast.literal_eval(prompts[0][4])
    upper_bounds_list=ast.literal_eval(prompts[0][5])
    guess_list=ast.literal_eval(prompts[0][6])
    num_vars=int(prompts[0][7])
    url=prompts[0][8]
    max_iter=int(prompts[0][9])

# Connect to Excel
excel = win32.DispatchEx('Excel.Application')
file_path = os.path.join(dest_folder, f'{base_name}_{node}')
workbook = excel.Workbooks.Open(file_path)
excel.Visible = False  # Keep Excel hidden    

worksheet = workbook.Sheets(sheet_name)
variable_range = worksheet.Range(range_address)
objective_cell = worksheet.Range(target)
formula=objective_cell.Formula

def fun(trial):

    objective_cell.Value=None

    for i in range(num_vars):
        var = trial.suggest_int(f"x{i}", lower_bounds_list[i], upper_bounds_list[i], step=1)

    for i, (key, value) in enumerate(trial.params.items()):
        variable_range.Cells(1, 1).GetOffset(i, 0).Value = int(value)  # Update the variable value
    
    objective_cell.Formula=formula

    excel.Calculate()

    while objective_cell.Value == None:
        pass

    objective_value = objective_cell.Value

    return objective_value

optuna.logging.set_verbosity(optuna.logging.WARNING)

lock_obj = optuna.storages.JournalFileOpenLock(url)
storage = optuna.storages.JournalStorage(
    optuna.storages.JournalFileStorage(url, lock_obj=lock_obj),
)

study = optuna.load_study(
    study_name="multi_cpu", storage=storage
)

study.optimize(fun, timeout=max_time)
