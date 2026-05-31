from pathlib import Path
import json
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'patient.xlsx'
ASSETS = ROOT / 'assets'
ASSETS.mkdir(parents=True, exist_ok=True)


def norm_cols(df):
    df.columns = [str(c).strip().lower().replace(' ', '') for c in df.columns]
    return df


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


def records(df):
    return [{k: clean(v) for k, v in r.items()} for r in df.to_dict(orient='records')]

xl = pd.ExcelFile(SRC)
sheets = {s.lower(): s for s in xl.sheet_names}

passport = pd.read_excel(SRC, sheet_name=sheets['passport'], header=None)
passport_items = []
for row in passport.itertuples(index=False):
    val = next((clean(x) for x in row if pd.notna(x) and str(x).strip()), None)
    if val is not None:
        passport_items.append(str(val))

passport_map = {}
for item in passport_items:
    if ':' in item:
        k, v = item.split(':', 1)
        passport_map[k.strip()] = v.strip()

try:
    diagnoses_df = pd.read_excel(SRC, sheet_name=sheets['diagnoses'])
    diagnoses = [str(x).strip() for x in diagnoses_df.iloc[:,0].dropna().tolist() if str(x).strip() and str(x).strip().lower() != 'diagnosis']
except Exception:
    diagnoses = []

key_df = norm_cols(pd.read_excel(SRC, sheet_name=sheets.get('key2026', sheets.get('key_2026'))))
key_markers = records(key_df)

labs_df = norm_cols(pd.read_excel(SRC, sheet_name=sheets['labsfull']))
if 'result' in labs_df.columns:
    labs_df = labs_df[labs_df['result'].notna()]
labs = records(labs_df)

vitals = records(norm_cols(pd.read_excel(SRC, sheet_name=sheets['vitals'])))
studies = records(norm_cols(pd.read_excel(SRC, sheet_name=sheets['studies'])))
therapy = records(norm_cols(pd.read_excel(SRC, sheet_name=sheets['therapyplan'])))

graphs = {}
for lower, original in sheets.items():
    if lower.startswith('graph'):
        slug = re.sub(r'^graph_?', '', lower)
        gdf = norm_cols(pd.read_excel(SRC, sheet_name=original))
        cols = list(gdf.columns)
        if len(cols) >= 2:
            dcol, vcol = cols[0], cols[1]
            rows = []
            for _, r in gdf.iterrows():
                d = clean(r.get(dcol))
                v = clean(r.get(vcol))
                if d is None or v is None:
                    continue
                rows.append({'date': d, 'value': v})
            graphs[slug] = rows

payload = {
    'passport_raw': passport_items,
    'passport': passport_map,
    'diagnoses': diagnoses,
    'key_markers': key_markers,
    'labs': labs,
    'vitals': vitals,
    'studies': studies,
    'therapy': therapy,
    'graphs': graphs,
    'graph_names': sorted(graphs.keys())
}

(ASSETS / 'patient-data.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
