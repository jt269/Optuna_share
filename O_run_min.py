import optuna
import pandas as pd
from optuna.samplers import NSGAIISampler
import sys
import subprocess
import os

# Settings
# Number of workers on this pc
cpu_num=4

# Storage path
dest_folder = os.getcwd()
journal_path=os.path.join(dest_folder,'journal.log')
script_relative_path = "Optuna_/O_worker_min.py"
worker_path=os.path.join(dest_folder,script_relative_path)

# Optimisation settings
max_time=14400
max_iter = 20000

lock_obj = optuna.storages.JournalFileOpenLock(journal_path)
storage = optuna.storages.JournalStorage(
optuna.storages.JournalFileStorage(journal_path, lock_obj=lock_obj),
)

# Create a study
study = optuna.create_study(study_name="multi_cpu", sampler=NSGAIISampler(), direction="maximize", storage=storage, load_if_exists=True)

procs = [subprocess.Popen([sys.executable, worker_path, str(i)]) for i in range(cpu_num)]

for p in procs:
    p.wait()

study = optuna.load_study(
    study_name="multi_cpu", storage=storage
)

# Save results
f_name=f"test_results.csv"
df = study.trials_dataframe()
df.to_csv(f_name, index=False)

print(study.best_params)
print(study.best_value)
