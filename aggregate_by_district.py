"""자치구별 월별 따릉이 이용 집계.

station_master.csv (대여소번호→자치구)를 원본 데이터와 조인해
25자치구 × 84개월 이용건수를 집계한다. 미매핑 대여소는 '미상'으로 처리.
"""
import glob, json, os, re
import pandas as pd

WORKDIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(WORKDIR)

master = pd.read_csv('station_master.csv')
master_map = dict(zip(master['대여소번호'].astype(str), master['자치구']))
print(f'Master: {len(master_map):,} stations, {master["자치구"].nunique()} districts')


def norm_ym(v):
    s = str(v)[:10]
    if re.fullmatch(r'\d{4}-\d{2}', s):
        return s[:4] + s[5:7]
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', s):
        return s[:4] + s[5:7]
    if re.fullmatch(r'\d{6}', s):
        return s
    if re.fullmatch(r'\d{8}', s):
        return s[:6]
    return None


def agg_df(df, label):
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    if '대여소' in df.columns and '대여소명' not in df.columns:
        df = df.rename(columns={'대여소': '대여소명'})
    if '이동시간(분)' in df.columns:
        df = df.rename(columns={'이동시간(분)': '이용시간(분)'})
    if '대여년월' in df.columns and '대여일자' not in df.columns:
        df = df.rename(columns={'대여년월': '대여일자'})
    df['ym'] = df['대여일자'].apply(norm_ym)
    df['대여소번호'] = df['대여소번호'].astype(str).str.strip()
    df['자치구'] = df['대여소번호'].map(master_map).fillna('미상')
    g = df.groupby(['ym', '자치구'], as_index=False)['이용건수'].sum()
    print(f'  {label}: {len(df):,} rows → {len(g):,} (ym,gu) pairs')
    return g


frames = []

for xl in sorted(glob.glob('2019/*.xlsx') + glob.glob('2020/*.xlsx') + glob.glob('2021/*.xlsx')):
    df = pd.read_excel(xl, usecols=['대여일자', '대여소번호', '이용건수'])
    frames.append(agg_df(df, os.path.basename(xl)))

def _read_csv_auto(p):
    for enc in ('utf-8-sig', 'utf-8', 'cp949', 'euc-kr'):
        try:
            return pd.read_csv(p, encoding=enc), enc
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f'encoding detect failed: {p}')

for csv in sorted(glob.glob('2019/*.csv') + glob.glob('2020/*.csv') + glob.glob('2021/*.csv')):
    full, enc = _read_csv_auto(csv)
    header = full.columns.tolist()
    date_col = '대여일자' if '대여일자' in header else ('대여년월' if '대여년월' in header else header[0])
    df = full[[date_col, '대여소번호', '이용건수']]
    frames.append(agg_df(df, os.path.basename(csv)))

for csv in sorted(glob.glob('2[2-5]_*.csv')):
    header = pd.read_csv(csv, encoding='cp949', nrows=0).columns.tolist()
    date_col = '대여일자' if '대여일자' in header else '대여년월'
    df = pd.read_csv(csv, encoding='cp949', usecols=[date_col, '대여소번호', '이용건수'])
    frames.append(agg_df(df, csv))

all_df = pd.concat(frames, ignore_index=True)
final = all_df.groupby(['ym', '자치구'], as_index=False)['이용건수'].sum()
final = final.sort_values(['ym', '자치구']).reset_index(drop=True)

print(f'\nTotal: {len(final):,} rows, {final["ym"].nunique()} months, {final["자치구"].nunique()} gus')
print('연도별 미상 비율:')
final['year'] = final['ym'].str[:4]
for y in sorted(final['year'].unique()):
    sub = final[final['year'] == y]
    tot = sub['이용건수'].sum()
    misu = sub[sub['자치구'] == '미상']['이용건수'].sum()
    print(f'  {y}: 총 {tot:>12,} / 미상 {misu:>11,} ({100*misu/tot:5.1f}%)')

final = final.drop(columns=['year'])
final.to_json('district_monthly.json', orient='records', force_ascii=False)
print('\nSaved district_monthly.json')

pivot = final.pivot(index='ym', columns='자치구', values='이용건수').fillna(0).astype(int)
pivot.to_csv('district_monthly_pivot.csv', encoding='utf-8-sig')
print('Saved district_monthly_pivot.csv (ym × gu)')
print('Shape:', pivot.shape)
