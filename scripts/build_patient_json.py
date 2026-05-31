from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
INFO_XLSX = ROOT / 'patient_data.xlsx'
DATA_XLSX = ROOT / 'patient_full.xlsx'
ASSETS = ROOT / 'assets'
ASSETS.mkdir(exist_ok=True)


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


def norm_cols(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def to_records(df):
    return [{k: clean(v) for k, v in r.items()} for r in df.to_dict(orient='records')]

# --- patient_data.xlsx: passport / diagnoses / studies / therapy ---
info_xl = pd.ExcelFile(INFO_XLSX)
info_sheets = {s.lower(): s for s in info_xl.sheet_names}

passport_raw = []
if 'passport' in info_sheets:
    ps = pd.read_excel(INFO_XLSX, sheet_name=info_sheets['passport'], header=None)
    for row in ps.itertuples(index=False):
        vals = [x for x in row if pd.notna(x) and str(x).strip()]
        if vals:
            passport_raw.append(str(vals[0]))

passport = {}
for item in passport_raw:
    if ':' in item:
        k, v = item.split(':', 1)
        passport[k.strip()] = v.strip()

if 'diagnoses' in info_sheets:
    diag_df = pd.read_excel(INFO_XLSX, sheet_name=info_sheets['diagnoses'])
    diagnoses = [str(x).strip() for x in diag_df.iloc[:, 0].dropna().tolist() if str(x).strip().lower() != 'diagnosis']
else:
    diagnoses = []

if 'studies' in info_sheets:
    studies_df = norm_cols(pd.read_excel(INFO_XLSX, sheet_name=info_sheets['studies']))
    studies = to_records(studies_df)
else:
    studies = []

if 'therapy' in info_sheets:
    th_df = norm_cols(pd.read_excel(INFO_XLSX, sheet_name=info_sheets['therapy']))
    therapy = to_records(th_df)
else:
    therapy = []

# --- patient_full.xlsx: единый лист data ---
flat_df = pd.read_excel(DATA_XLSX, sheet_name='data')
flat_df = norm_cols(flat_df)

rename_map = {
    'дата': 'date',
    'группа': 'group',
    'показатель': 'marker',
    'значение': 'value',
    'единица измерения': 'unit',
    'референс': 'reference',
    'комментарий': 'comment',
}
flat_df = flat_df.rename(columns=rename_map)

for col in ['date','group','marker','value','unit','reference','comment']:
    if col not in flat_df.columns:
        flat_df[col] = None

flat_df['marker'] = flat_df['marker'].astype(str).str.strip()
flat_df['group'] = flat_df['group'].astype(str).str.strip()
flat_df = flat_df[flat_df['marker'] != '']

flat_df['value'] = flat_df['value'].apply(clean)
flat_df['reference'] = flat_df['reference'].apply(clean)
flat_df['comment'] = flat_df['comment'].apply(clean)

rows = to_records(flat_df[['date','group','marker','value','unit','reference','comment']])

groups = sorted({r['group'] for r in rows if r.get('group')})
markers = sorted({r['marker'] for r in rows if r.get('marker')})

series = {}
for r in rows:
    m = r.get('marker')
    if not m:
        continue
    d = r.get('date')
    v = r.get('value')
    if d is None or v is None:
        continue
    series.setdefault(m, []).append({
        'date': d,
        'value': v,
        'group': r.get('group'),
        'unit': r.get('unit'),
    })

payload = {
    'passport_raw': passport_raw,
    'passport': passport,
    'diagnoses': diagnoses,
    'studies': studies,
    'therapy': therapy,
    'rows': rows,
    'groups': groups,
    'markers': markers,
    'series': series,
}

(ASSETS / 'patient-data.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
