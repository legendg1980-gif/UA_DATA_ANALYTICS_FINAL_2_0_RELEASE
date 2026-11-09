from pathlib import Path
import csv
ROOT=Path(__file__).resolve().parents[1]

def count_csv(path, delimiter=","):
    with open(path,encoding="utf-8-sig",newline="") as f:
        return sum(1 for _ in csv.reader(f,delimiter=delimiter))-1

print("QA starter")
print("stg_contracts:", count_csv(ROOT/"04_STAGING/stg_contracts.csv"))
print("stg_payments:", count_csv(ROOT/"04_STAGING/stg_payments.csv"))
# TODO: додайте перевірки DQ001-DQ020
