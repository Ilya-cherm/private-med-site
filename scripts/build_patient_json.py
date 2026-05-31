from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_XLSX = ROOT / 'patient_data.xlsx'
INFO_XLSX = ROOT / 'patient_info.xlsx'
ASSETS = ROOT / 'assets'
ASSETS.mkdir(exist_ok=True)


def clean_value(v):
    if pd.isna(v):
        return ''
    if hasattr(v, 'isoformat'):
        try:
            return v.isoformat()
        except Exception:
            pass
    return str(v).strip()


def clean_number(v):
    if pd.isna(v):
        return None
    try:
        f = float(v)
        if f.is_integer():
            return int(f)
        return f
    except Exception:
        s = str(v).strip().replace(' ', '').replace(',', '.')
        try:
            f = float(s)
            if f.is_integer():
                return int(f)
            return f
        except Exception:
            return s


def normalize_date(v):
    if pd.isna(v):
        return ''
    s = str(v).strip()
    if not s:
        return ''
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return s


# patient_data.xlsx -> assets/data.json
raw = pd.read_excel(DATA_XLSX, sheet_name='data')
rows = []
for _, r in raw.iterrows():
    row = {
        'Дата': normalize_date(r.get('Дата')),
        'Группа': clean_value(r.get('Группа')),
        'Показатель': clean_value(r.get('Показатель')),
        'Значение': clean_number(r.get('Значение')),
        'Единицы': clean_value(r.get('Единицы')),
        'Нижняя граница': clean_number(r.get('Нижняя граница')),
        'Верхняя граница': clean_number(r.get('Верхняя граница')),
        'Комментарий': clean_value(r.get('Комментарий')),
    }
    if row['Дата'] and row['Показатель']:
        rows.append(row)

rows.sort(key=lambda x: (x['Дата'], x['Группа'], x['Показатель']))

with open(ASSETS / 'data.json', 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)

# patient_info.xlsx -> assets/info.json
passport = pd.read_excel(INFO_XLSX, sheet_name='passport').fillna('')
diagnoses = pd.read_excel(INFO_XLSX, sheet_name='diagnoses').fillna('')
studies = pd.read_excel(INFO_XLSX, sheet_name='studies').fillna('')
therapy = pd.read_excel(INFO_XLSX, sheet_name='therapy').fillna('')

info = {
    'passport': passport.to_dict(orient='records'),
    'diagnoses': diagnoses.to_dict(orient='records'),
    'studies': studies.to_dict(orient='records'),
    'therapy': therapy.to_dict(orient='records'),
}

with open(ASSETS / 'info.json', 'w', encoding='utf-8') as f:
    json.dump(info, f, ensure_ascii=False, indent=2)
