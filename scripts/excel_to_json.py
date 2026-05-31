import pandas as pd
import json
import math

INPUT_FILE = "patient.xlsx"
OUTPUT_FILE = "assets/data.json"

ANON_COLUMNS = [
    "ФИО", "ФИО пациента", "Пациент", "Имя", "Фамилия", "Отчество",
    "Name", "Patient", "Patient Name", "Full Name"
]

MARKER_KEYS = ["Показатель", "Анализ", "marker", "parameter", "name", "Marker", "Parameter"]
DATE_KEYS = ["Дата", "date", "Date"]
VALUE_KEYS = ["Значение", "value", "Value"]
UNIT_KEYS = ["Единицы", "Ед", "unit", "units", "Unit"]
REF_LOW_KEYS = ["Нижняя граница", "РефНиз", "refLow", "referenceLow", "low"]
REF_HIGH_KEYS = ["Верхняя граница", "РефВерх", "refHigh", "referenceHigh", "high"]
COMMENT_KEYS = ["Комментарий", "comment", "notes", "Note"]

def is_empty(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if str(v).strip() == "":
        return True
    return False

def first_existing(row, keys):
    for key in keys:
        if key in row and not is_empty(row[key]):
            return row[key]
    return ""

def normalize_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def anonymize(df):
    for col in df.columns:
        if str(col).strip() in ANON_COLUMNS:
            df[col] = "Аноним"
    return df

all_sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
result = []

for sheet_name, df in all_sheets.items():
    df = normalize_columns(df)
    df = anonymize(df)
    df = df.fillna("")

    for _, row in df.iterrows():
        row = row.to_dict()

        marker = first_existing(row, MARKER_KEYS)
        date = first_existing(row, DATE_KEYS)
        value = first_existing(row, VALUE_KEYS)
        unit = first_existing(row, UNIT_KEYS)
        ref_low = first_existing(row, REF_LOW_KEYS)
        ref_high = first_existing(row, REF_HIGH_KEYS)
        comment = first_existing(row, COMMENT_KEYS)

        if is_empty(marker) or is_empty(date) or is_empty(value):
            continue

        record = {
            "patient": "Аноним",
            "marker": str(marker).strip(),
            "date": str(date).strip(),
            "value": value,
            "unit": "" if is_empty(unit) else str(unit).strip(),
            "refLow": "" if is_empty(ref_low) else ref_low,
            "refHigh": "" if is_empty(ref_high) else ref_high,
            "comment": "" if is_empty(comment) else str(comment).strip(),
            "sheet": str(sheet_name).strip()
        }

        result.append(record)

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
