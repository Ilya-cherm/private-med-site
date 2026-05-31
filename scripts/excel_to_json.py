import pandas as pd
import json

INPUT_FILE = "patient.xlsx"
OUTPUT_FILE = "assets/data.json"

def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

all_sheets = pd.read_excel(INPUT_FILE, sheet_name=None, header=0, dtype=str)

result = []

for sheet_name, df in all_sheets.items():
    df.columns = [clean(c) for c in df.columns]

    for _, row in df.iterrows():
        values = [clean(v) for v in row.tolist()]
        non_empty = [v for v in values if v]

        if len(non_empty) < 2:
            continue

        marker = non_empty[0]
        value = non_empty[1]
        unit = non_empty[2] if len(non_empty) > 2 else ""
        comment = non_empty[3] if len(non_empty) > 3 else ""

        if marker.lower() in [
            "пациент", "возраст", "пол", "дата рождения",
            "дата составления", "passport", "цифровой регистр пациента",
            "версия от"
        ]:
            continue

        result.append({
            "patient": "Аноним",
            "date": sheet_name,
            "marker": marker,
            "value": value,
            "unit": unit,
            "refLow": "",
            "refHigh": "",
            "comment": comment,
            "sheet": sheet_name
        })

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
