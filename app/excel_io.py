import pandas as pd
from app.config import INPUT_FILE, OUTPUT_FILE

def read_employees():
    df = pd.read_excel(INPUT_FILE)
    required = {"employee_id", "employee_name"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in employee_input.xlsx: {missing}")
    return df.fillna("")

def write_results(results):
    df = pd.DataFrame(results)
    df.to_excel(OUTPUT_FILE, index=False)