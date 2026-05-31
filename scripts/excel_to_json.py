import pandas as pd
import json

INPUT_FILE = "patient.xlsx"
OUTPUT_FILE = "assets/data.json"

ANON_COLUMNS = [
    "ФИО", "ФИО пациента", "Пациент", "Имя", "Фамилия", "Отчество",
    "Name", "Patient", "Patient Name", "Full Name"
]

df = pd.read_excel(INPUT_FILE)
df = df.fillna("")

for col in df.columns:
    if str(col).strip() in ANON_COLUMNS:
        df[col] = "Аноним"

records = df.to_dict(orient="records")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=2)
