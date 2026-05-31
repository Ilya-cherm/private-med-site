from pathlib import Path
import json
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / 'assets'
ASSETS.mkdir(exist_ok=True)

DATA_CANDIDATES = ['patient_data.xlsx', 'patient-data.xlsx', 'patient.xlsx', 'data.xlsx']
INFO_CANDIDATES = ['patient_info.xlsx', 'patient-info.xlsx', 'info.xlsx']


def pick(candidates):
    for name in candidates:
        p = ROOT / name
        if p.exists():
            return p
    return None


def norm(s):
    s = '' if s is None else str(s)
    s = s.strip().lower().replace('ё', 'е')
    s = re.sub(r'\s+', ' ', s)
    return s


def slug(s):
    s = norm(s)
    s = re.sub(r'[^a-zа-я0-9]+', '_', s)
    return s.strip('_') or 'metric'


def clean(v):
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.strftime('%Y-%m-%d')
    if hasattr(v, 'item'):
        try:
            v = v.item()
        except Exception:
            pass
    return v


def maybe_date(v):
    v = clean(v)
    if v is None:
        return None
    try:
        return pd.to_datetime(v).strftime('%Y-%m-%d')
    except Exception:
        return str(v)


def find_sheet(xls, wants):
    names = list(xls.sheet_names)
    mapping = {name: norm(name) for name in names}
    for want in wants:
        w = norm(want)
        for original, n in mapping.items():
            if n == w or w in n:
                return original
    return names[0] if names else None


def find_col(columns, variants):
    cols = [norm(c) for c in columns]
    for variant in variants:
        v = norm(variant)
        for i, c in enumerate(cols):
            if c == v:
                return columns[i]
        for i, c in enumerate(cols):
            if v in c or c in v:
                return columns[i]
    return None


def read_data_file(path):
    xls = pd.ExcelFile(path)
    sheet = find_sheet(xls, ['data', 'данные', 'labs', 'анализы', 'sheet1'])
    df = pd.read_excel(path, sheet_name=sheet)
    cols = list(df.columns)

    c_date = find_col(cols, ['date', 'дата'])
    c_group = find_col(cols, ['group', 'section', 'группа', 'раздел'])
    c_marker = find_col(cols, ['marker', 'metric', 'name', 'показатель', 'анализ'])
    c_result = find_col(cols, ['result', 'value', 'результат', 'значение'])
    c_unit = find_col(cols, ['unit', 'units', 'ед', 'ед.', 'единица'])
    c_low = find_col(cols, ['lower_ref', 'ref_low', 'нижняя граница', 'нижняя_граница', 'min'])
    c_high = find_col(cols, ['upper_ref', 'ref_high', 'верхняя граница', 'верхняя_граница', 'max'])
    c_ref = find_col(cols, ['reference_text', 'reference', 'референс', 'референсные значения', 'норма'])
    c_note = find_col(cols, ['clinical_note', 'comment', 'комментарий', 'assessment', 'trend', 'заметка'])

    records = []
    metrics = {}

    for _, row in df.iterrows():
        marker = clean(row[c_marker]) if c_marker else None
        if not marker:
            continue
        group = clean(row[c_group]) if c_group else None
        key = slug(f'{group or ""}_{marker}')
        rec = {
            'date': maybe_date(row[c_date]) if c_date else None,
            'section': group,
            'group': group,
            'marker': marker,
            'marker_key': key,
            'result': clean(row[c_result]) if c_result else None,
            'unit': clean(row[c_unit]) if c_unit else None,
            'lower_ref': clean(row[c_low]) if c_low else None,
            'upper_ref': clean(row[c_high]) if c_high else None,
            'reference_text': clean(row[c_ref]) if c_ref else None,
            'clinical_note': clean(row[c_note]) if c_note else None,
        }
        records.append(rec)
        if key not in metrics:
            metrics[key] = {'key': key, 'label': marker, 'group': group}

    return list(metrics.values()), records


def rows_to_dicts(df):
    out = []
    for _, row in df.iterrows():
        item = {}
        for c in df.columns:
            v = clean(row[c])
            if v is not None and str(v).strip() != '':
                item[str(c)] = v
        if item:
            out.append(item)
    return out


def read_info_file(path):
    patient = {}
    diagnoses = []
    therapy = []

    if not path or not path.exists():
        return patient, diagnoses, therapy

    xls = pd.ExcelFile(path)
    names = list(xls.sheet_names)

    patient_sheet = find_sheet(xls, ['passport', 'patient', 'паспорт', 'паспортная часть', 'info'])
    diag_sheet = find_sheet(xls, ['diagnosis', 'diagnoses', 'диагноз', 'диагнозы'])
    therapy_sheet = find_sheet(xls, ['therapy', 'терапия', 'treatment', 'лечение'])

    if patient_sheet:
        dfp = pd.read_excel(path, sheet_name=patient_sheet)
        if len(dfp.columns) >= 2:
            kcol, vcol = dfp.columns[0], dfp.columns[1]
            for _, row in dfp.iterrows():
                k = clean(row[kcol])
                v = clean(row[vcol])
                if k and v is not None:
                    patient[str(k)] = v
        elif len(dfp.columns) == 1:
            for idx, value in enumerate(dfp.iloc[:, 0].tolist(), start=1):
                v = clean(value)
                if v is not None:
                    patient[f'Поле {idx}'] = v

    if diag_sheet:
        dfd = pd.read_excel(path, sheet_name=diag_sheet)
        diagnoses = rows_to_dicts(dfd)

    if therapy_sheet:
        dft = pd.read_excel(path, sheet_name=therapy_sheet)
        therapy = rows_to_dicts(dft)

    return patient, diagnoses, therapy


data_file = pick(DATA_CANDIDATES)
info_file = pick(INFO_CANDIDATES)

if not data_file:
    found = sorted([p.name for p in ROOT.glob('*.xlsx')])
    raise FileNotFoundError('Не найден файл с анализами. Ищутся: ' + ', '.join(DATA_CANDIDATES) + (f'. Найдены: {", ".join(found)}' if found else '.'))

metrics, labs = read_data_file(data_file)
patient, diagnoses, therapy = read_info_file(info_file)

payload = {
    'patient': patient,
    'diagnoses': diagnoses,
    'therapy': therapy,
    'metrics': metrics,
    'labs': labs,
}

(ASSETS / 'patient-data.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'Built {ASSETS / "patient-data.json"} from {data_file.name}' + (f' and {info_file.name}' if info_file else ''))
