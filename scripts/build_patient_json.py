import json
import math
import re
from pathlib import Path

import pandas as pd

ROOT = Path('.').resolve()
ASSETS = ROOT / 'assets'
ASSETS.mkdir(exist_ok=True)

PATIENT_XLSX = ROOT / 'patient.xlsx'


def clean_text(value):
    if value is None:
        return ''
    if isinstance(value, float) and math.isnan(value):
        return ''
    text = str(value).replace('\xa0', ' ').strip()
    return '' if text.lower() == 'nan' else text


def clean_num(value):
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(',', '.')
    m = re.fullmatch(r'-?\d+(?:\.\d+)?', text)
    return float(text) if m else None


def to_iso_date(value):
    text = clean_text(value)
    if not text:
        return ''
    dt = pd.to_datetime(text, dayfirst=True, errors='coerce')
    if pd.notna(dt):
        return dt.strftime('%Y-%m-%d')
    m = re.fullmatch(r'(\d{4})', text)
    if m:
        return f'{m.group(1)}-01-01'
    return text


def normalize_marker(value):
    text = clean_text(value)
    text = re.sub(r'\s+', ' ', text)
    return text.strip(' .,-')


def marker_slug(value):
    text = normalize_marker(value).lower()
    text = text.replace('25-oh', '25 oh').replace('hs-crp', 'hs crp')
    text = re.sub(r'[^a-zа-я0-9]+', '_', text, flags=re.IGNORECASE)
    return text.strip('_')


def parse_reference(ref_text):
    text = clean_text(ref_text)
    if not text:
        return None, None, ''
    normalized = text.replace('–', '-').replace('—', '-').replace(',', '.')
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    m = re.fullmatch(r'(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)', normalized)
    if m:
        return float(m.group(1)), float(m.group(2)), text
    m = re.fullmatch(r'(-?\d+(?:\.\d+)?)', normalized)
    if m:
        return None, float(m.group(1)), text
    m = re.search(r'(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)', normalized)
    if m and normalized.count('-') == 1:
        return float(m.group(1)), float(m.group(2)), text
    return None, None, text


xls = pd.ExcelFile(PATIENT_XLSX)

passport = pd.read_excel(xls, sheet_name='passport').fillna('')
diagnoses = pd.read_excel(xls, sheet_name='diagnoses').fillna('')
labs = pd.read_excel(xls, sheet_name='labsfull').fillna('')
vitals = pd.read_excel(xls, sheet_name='vitals').fillna('')
studies = pd.read_excel(xls, sheet_name='studies').fillna('')
therapy = pd.read_excel(xls, sheet_name='therapyplan').fillna('')

passport_info = {}
if not passport.empty:
    row = passport.iloc[0].to_dict()
    passport_info = {str(k): clean_text(v) for k, v in row.items() if clean_text(v)}

lab_rows = []
metric_map = {}
for _, row in labs.iterrows():
    marker = normalize_marker(row.get('marker', ''))
    result = clean_num(row.get('result'))
    date_iso = to_iso_date(row.get('date'))
    ref_text = clean_text(row.get('reference'))
    if not marker or result is None or not date_iso:
        continue
    lower_ref, upper_ref, reference_text = parse_reference(ref_text)
    item = {
        'date': date_iso,
        'section': clean_text(row.get('section')),
        'marker': marker,
        'marker_key': marker_slug(marker),
        'result': result,
        'unit': clean_text(row.get('unit')),
        'reference_text': reference_text,
        'lower_ref': lower_ref,
        'upper_ref': upper_ref,
        'assessment': clean_text(row.get('assessment')),
        'trend': clean_text(row.get('trend')),
        'clinical_note': clean_text(row.get('clinicalnote')),
        'source_sheet': 'labsfull'
    }
    lab_rows.append(item)
    mk = item['marker_key']
    metric_map.setdefault(mk, {
        'key': mk,
        'label': marker,
        'unit': item['unit'],
        'group': clean_text(row.get('section')) or 'Другое',
        'reference_text': reference_text,
        'lower_ref': lower_ref,
        'upper_ref': upper_ref
    })
    if reference_text and not metric_map[mk].get('reference_text'):
        metric_map[mk]['reference_text'] = reference_text
    if lower_ref is not None and metric_map[mk].get('lower_ref') is None:
        metric_map[mk]['lower_ref'] = lower_ref
    if upper_ref is not None and metric_map[mk].get('upper_ref') is None:
        metric_map[mk]['upper_ref'] = upper_ref

vitals_map = {
    'weightkg': ('weightkg', 'Вес', 'kg', 'Витальные'),
    'sbp': ('sbp', 'САД', 'mmHg', 'Витальные'),
    'dbp': ('dbp', 'ДАД', 'mmHg', 'Витальные'),
    'pulse': ('pulse', 'Пульс', 'bpm', 'Витальные'),
}
for _, row in vitals.iterrows():
    date_iso = to_iso_date(row.get('date'))
    if not date_iso:
        continue
    for col, meta in vitals_map.items():
        val = clean_num(row.get(col))
        if val is None:
            continue
        key, label, unit, group = meta
        item = {
            'date': date_iso,
            'section': group,
            'marker': label,
            'marker_key': key,
            'result': val,
            'unit': unit,
            'reference_text': '',
            'lower_ref': None,
            'upper_ref': None,
            'assessment': '',
            'trend': '',
            'clinical_note': clean_text(row.get('comment')),
            'source_sheet': 'vitals'
        }
        lab_rows.append(item)
        metric_map.setdefault(key, {
            'key': key,
            'label': label,
            'unit': unit,
            'group': group,
            'reference_text': '',
            'lower_ref': None,
            'upper_ref': None
        })

lab_rows = sorted(lab_rows, key=lambda x: (x['marker_key'], x['date']))
metrics = sorted(metric_map.values(), key=lambda x: (x['group'], x['label']))

payload = {
    'generated_at': pd.Timestamp.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
    'patient': passport_info,
    'diagnoses': [
        {k: clean_text(v) for k, v in row.items() if clean_text(v)}
        for _, row in diagnoses.iterrows()
        if any(clean_text(v) for v in row.values)
    ],
    'studies': [
        {k: clean_text(v) for k, v in row.items() if clean_text(v)}
        for _, row in studies.iterrows()
        if any(clean_text(v) for v in row.values)
    ],
    'therapy': [
        {k: clean_text(v) for k, v in row.items() if clean_text(v)}
        for _, row in therapy.iterrows()
        if any(clean_text(v) for v in row.values)
    ],
    'metrics': metrics,
    'labs': lab_rows,
}

(ASSETS / 'patient-data.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
print('written assets/patient-data.json')
