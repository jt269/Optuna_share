import optuna
import os

num_vars=10
max_time=3600
dest_folder = os.getcwd()
journal_path=os.path.join(dest_folder,'journal.log')

def fun(trial):

    for i in range(num_vars):
        var = trial.suggest_int(f"x{i}", 1, 12, step=1)

    objective_value = sum(trial.params.values())

    return objective_value

optuna.logging.set_verbosity(optuna.logging.WARNING)

lock_obj = optuna.storages.JournalFileOpenLock(journal_path)
storage = optuna.storages.JournalStorage(
    optuna.storages.JournalFileStorage(journal_path, lock_obj=lock_obj),
)

study = optuna.load_study(
    study_name="multi_cpu", storage=storage
)

study.optimize(fun, timeout=max_time)
