from pathlib import Path
import json
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / 'assets'
ASSETS.mkdir(exist_ok=True)

CANDIDATES = [
    'patient.xlsx',
    'patient_data.xlsx',
    'patient-data.xlsx',
    'data.xlsx',
]

PATIENT_XLSX = None
for name in CANDIDATES:
    p = ROOT / name
    if p.exists():
        PATIENT_XLSX = p
        break

if PATIENT_XLSX is None:
    found = sorted([p.name for p in ROOT.glob('*.xlsx')])
    raise FileNotFoundError(
        'Не найден Excel-файл пациента. Ищутся имена: '
        + ', '.join(CANDIDATES)
        + (f'. Найдены только: {", ".join(found)}' if found else '. В корне репозитория .xlsx файлов нет.')
    )


def norm_key(s):
    s = '' if s is None else str(s)
    s = s.strip().lower().replace('ё', 'е')
    s = re.sub(r'\s+', ' ', s)
    return s


def slugify(s):
    s = norm_key(s)
    s = re.sub(r'[^a-zа-я0-9]+', '_', s)
    return s.strip('_') or 'metric'


def val(x):
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp,)):
        return x.strftime('%Y-%m-%d')
    if hasattr(x, 'item'):
        try:
            x = x.item()
        except Exception:
            pass
    return x


def find_sheet(xls, wants):
    names = list(xls.sheet_names)
    normalized = {name: norm_key(name) for name in names}
    for want in wants:
        w = norm_key(want)
        for original, n in normalized.items():
            if n == w or w in n:
                return original
    return names[0] if names else None


xls = pd.ExcelFile(PATIENT_XLSX)
sheet = find_sheet(xls, ['data', 'данные', 'labs', 'анализы', 'sheet1'])
df = pd.read_excel(PATIENT_XLSX, sheet_name=sheet)
df.columns = [norm_key(c) for c in df.columns]

col_map = {
    'date': ['date', 'дата'],
    'section': ['section', 'group', 'группа', 'раздел'],
    'marker': ['marker', 'metric', 'name', 'показатель', 'анализ'],
    'result': ['result', 'value', 'результат', 'значение'],
    'unit': ['unit', 'units', 'ед', 'ед.', 'unit_name'],
    'lower_ref': ['lower_ref', 'ref_low', 'нижняя граница', 'нижняя_граница', 'min'],
    'upper_ref': ['upper_ref', 'ref_high', 'верхняя граница', 'верхняя_граница', 'max'],
    'reference_text': ['reference_text', 'reference', 'референс', 'референсные значения', 'норма'],
    'clinical_note': ['clinical_note', 'comment', 'комментарий', 'заметка', 'assessment', 'trend'],
}

resolved = {}
for target, variants in col_map.items():
    resolved[target] = None
    for variant in variants:
      nv = norm_key(variant)
      for col in df.columns:
        if col == nv:
          resolved[target] = col
          break
      if resolved[target]:
        break

records = []
for _, row in df.iterrows():
    marker = val(row.get(resolved['marker'])) if resolved['marker'] else None
    if not marker:
        continue
    section = val(row.get(resolved['section'])) if resolved['section'] else None
    marker_key = slugify(f"{section or ''}_{marker}")
    date_val = val(row.get(resolved['date'])) if resolved['date'] else None
    if isinstance(date_val, str):
        try:
            date_val = pd.to_datetime(date_val).strftime('%Y-%m-%d')
        except Exception:
            pass
    rec = {
        'date': date_val,
        'section': section,
        'group': section,
        'marker': marker,
        'marker_key': marker_key,
        'result': val(row.get(resolved['result'])) if resolved['result'] else None,
        'unit': val(row.get(resolved['unit'])) if resolved['unit'] else None,
        'lower_ref': val(row.get(resolved['lower_ref'])) if resolved['lower_ref'] else None,
        'upper_ref': val(row.get(resolved['upper_ref'])) if resolved['upper_ref'] else None,
        'reference_text': val(row.get(resolved['reference_text'])) if resolved['reference_text'] else None,
        'clinical_note': val(row.get(resolved['clinical_note'])) if resolved['clinical_note'] else None,
    }
    records.append(rec)

metrics_map = {}
for r in records:
    if r['marker_key'] not in metrics_map:
        metrics_map[r['marker_key']] = {
            'key': r['marker_key'],
            'label': r['marker'],
            'group': r['group'],
        }

payload = {
    'patient': {},
    'diagnoses': [],
    'therapy': [],
    'metrics': list(metrics_map.values()),
    'labs': records,
}

(ASSETS / 'patient-data.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'Built {ASSETS / "patient-data.json"} from {PATIENT_XLSX.name}')
