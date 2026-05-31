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

        if len(non_empty) < 3:
            continue

        date = non_empty[0]
        marker = non_empty[1]
        value = non_empty[2]

        if marker.lower() in ["пациент", "возраст", "пол", "дата рождения", "дата составления", "passport"]:
            continue

        result.append({
            "patient": "Аноним",
            "date": date,
            "marker": marker,
            "value": value,
            "unit": non_empty[3] if len(non_empty) > 3 else "",
            "refLow": non_empty[4] if len(non_empty) > 4 else "",
            "refHigh": non_empty[5] if len(non_empty) > 5 else "",
            "comment": non_empty[6] if len(non_empty) > 6 else "",
            "sheet": sheet_name
        })

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
