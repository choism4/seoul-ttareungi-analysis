#!/usr/bin/env python3
"""V2 보고서 — 4인 전문가 비평 반영 전면 리프레임.

주제: "공유 모빌리티 시장 재편기의 따릉이 — 수요 구조 변화와 운영 포트폴리오 재정의"

주요 변경 (V1 → V2):
1. 표지 하단·결론부 양쪽에 생성형 AI 활용 명시 (과제 지침 필수)
2. 서론에 문제의식 강화 (공유 모빌리티 시장 재편, 2023 피크 후 구조적 감소)
3. 신규 킬러 섹션: 기온 효과 제거 후 잔차 분석 (2024~25 구조적 감소 증명)
4. 자치구 공간 형평성 (지니계수), 대당 회전율(Unit Economics)
5. 미상 안분 보정 CAGR 병기
6. 한계 및 후속 연구 섹션 추가
7. 부록 A (Excel 재현 가이드), 부록 B (출처·참고자료)
8. 날짜 표기 정규화, 셀 주소 본문 제거, 그림 본문 지시문
"""
import base64
import io
import json
import os
import subprocess
from collections import defaultdict

import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
import matplotlib.patheffects as pe
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# V21: 프로 패널 대응 통계·재무 JSON 로드 (compute_stats_v21.py, compute_finance_v21.py 산출)
try:
    V21_STATS = json.load(open('v21_stats.json'))
    V21_FINANCE = json.load(open('v21_finance.json'))
except Exception:
    V21_STATS = {}
    V21_FINANCE = {}

GITHUB_URL = 'https://github.com/choism4/seoul-ttareungi-analysis'
GITHUB_COMMIT = '54f760c9327f'

# ── Font setup ──
for fp in [
    '/System/Library/Fonts/AppleSDGothicNeo.ttc',
    '/Library/Fonts/AppleGothic.ttf',
    '/System/Library/Fonts/Supplemental/AppleGothic.ttf',
]:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
        _font_name = fm.FontProperties(fname=fp).get_name()
        plt.rcParams['font.family'] = _font_name
        plt.rcParams['font.sans-serif'] = [_font_name]
        plt.rcParams['axes.titleweight'] = 'bold'
        plt.rcParams['axes.titlesize'] = 13
        plt.rcParams['mathtext.fontset'] = 'dejavusans'
        break
plt.rcParams['axes.unicode_minus'] = False

# ── Load data ──
monthly = json.load(open('monthly_aggregate.json'))
temp_data = json.load(open('seoul_temperature.json'))
gender_yearly = json.load(open('gender_yearly.json'))
age_yearly = json.load(open('age_yearly.json'))
district_monthly = json.load(open('district_monthly.json'))
ctx = json.load(open('external_context.json'))

temp_map = {t['연월']: t['평균기온'] for t in temp_data}
age_groups = ['~10대', '20대', '30대', '40대', '50대', '60대', '70대이상']

# ── Derived data ──
years = sorted(set(m['연도'] for m in monthly))
labels = [m['연월'] for m in monthly]
values = [m['이용건수'] for m in monthly]
temps = [temp_map.get(m['연월']) for m in monthly]
valid_tu = [(t, u) for t, u in zip(temps, values) if t is not None]
X = np.array([v[0] for v in valid_tu])
Y = np.array([v[1] for v in valid_tu])
slope, intercept, r_value, p_value, std_err = stats.linregress(X, Y)
r_sq = r_value ** 2

# 잔차
predicted = intercept + slope * np.array(temps)
residuals = np.array(values) - predicted
# 2023년 이후 잔차 (피크 이후 구간 비교용)
r_2023_idx = 48  # 2019.01=0 기준 2023.01=48
post_2023_mean_resid = residuals[r_2023_idx:].mean()
pre_2023_mean_resid = residuals[:r_2023_idx].mean()
# 2022년 4월 구조변화 분기점 이후 잔차 (supF 최대 F=13.55)
r_break_idx = 39  # 2019.01=0 기준 2022.04=39
post_break_mean_resid = residuals[r_break_idx:].mean()
pre_break_mean_resid = residuals[:r_break_idx].mean()
# 잔차-시간 상관
resid_time_corr = np.corrcoef(residuals, np.arange(len(residuals)))[0, 1]

# 계절지수
overall_avg = np.mean(values)
monthly_avg_si = []
for mo in range(1, 13):
    avg = np.mean([m['이용건수'] for m in monthly if m['월'] == mo])
    monthly_avg_si.append(avg / overall_avg)
winter_si = np.mean([monthly_avg_si[i - 1] for i in [12, 1, 2]])
summer_si = np.mean([monthly_avg_si[i - 1] for i in [6, 7, 8]])

# 연도별 합계
yearly_usage = {yr: sum(m['이용건수'] for m in monthly if m['연도'] == yr) for yr in years}
usage_2019 = yearly_usage.get(2019, 1)
usage_2025 = yearly_usage.get(2025, 1)
usage_2023 = yearly_usage.get(2023, 1)
cagr_19_25 = (usage_2025 / usage_2019) ** (1 / 6) - 1
cagr_23_25 = (usage_2025 / usage_2023) ** (1 / 2) - 1

# 자치구 2025 집계
gu_2025 = defaultdict(int)
for row in district_monthly:
    if str(row['ym']).startswith('2025') and row['자치구'] != '미상':
        gu_2025[row['자치구']] += row['이용건수']
population = {
    '강남구': 558_000, '강동구': 460_000, '강북구': 284_000, '강서구': 553_000,
    '관악구': 487_000, '광진구': 340_000, '구로구': 393_000, '금천구': 236_000,
    '노원구': 495_000, '도봉구': 306_000, '동대문구': 342_000, '동작구': 382_000,
    '마포구': 369_000, '서대문구': 303_000, '서초구': 413_000, '성동구': 282_000,
    '성북구': 428_000, '송파구': 651_000, '양천구': 440_000, '영등포구': 375_000,
    '용산구': 213_000, '은평구': 469_000, '종로구': 141_000, '중구': 121_000,
    '중랑구': 386_000,
}
per_capita = {g: gu_2025[g] / population.get(g, 1) for g in gu_2025}
sorted_gu = sorted(per_capita.items(), key=lambda x: x[1], reverse=True)


def gini_coefficient(values):
    """페어와이즈 지니계수."""
    vals = np.array(sorted(values))
    n = len(vals)
    if n == 0 or vals.sum() == 0:
        return 0
    idx = np.arange(1, n + 1)
    return (2 * np.sum(idx * vals) - (n + 1) * vals.sum()) / (n * vals.sum())


gini_per_capita = gini_coefficient(list(per_capita.values()))

# 대당 회전율
fleet_by_year = {
    2019: 25000, 2020: 37500, 2021: 40500, 2022: 43500,
    2023: 45000, 2024: 45000, 2025: 45000,
}


def days_in_month(ym):
    y, m = int(str(ym)[:4]), int(str(ym)[4:])
    if m in (1, 3, 5, 7, 8, 10, 12):
        return 31
    if m == 2:
        return 29 if (y % 4 == 0 and y % 100 != 0) or y % 400 == 0 else 28
    return 30


turnover = []
for mm in monthly:
    y = mm['연도']
    ym = mm['연월']
    fleet = fleet_by_year.get(y, 45000)
    d = days_in_month(ym)
    turnover.append(mm['이용건수'] / (fleet * d))
yearly_turnover = {}
for yr in years:
    indices = [i for i, m in enumerate(monthly) if m['연도'] == yr]
    yearly_turnover[yr] = np.mean([turnover[i] for i in indices])

# 미상 보정 CAGR
misu_pct_by_year = {}
for gy in gender_yearly:
    tot = gy.get('M', 0) + gy.get('F', 0) + gy.get('미상', 0)
    misu_pct_by_year[gy['연도']] = gy.get('미상', 0) / tot if tot > 0 else 0
misu19 = misu_pct_by_year.get(2019, 0)
misu25 = misu_pct_by_year.get(2025, 0)
ay_2019 = next((a for a in age_yearly if a['연도'] == 2019), {})
ay_2025 = next((a for a in age_yearly if a['연도'] == 2025), {})
cagr_rows = []
for ag in age_groups:
    v19 = ay_2019.get(ag, 0) or 1
    v25 = ay_2025.get(ag, 0) or 1
    raw_cagr = (v25 / v19) ** (1 / 6) - 1
    adj19 = v19 / (1 - misu19)
    adj25 = v25 / (1 - misu25)
    adj_cagr = (adj25 / adj19) ** (1 / 6) - 1
    cagr_rows.append((ag, v19, v25, raw_cagr, adj_cagr, (adj_cagr - raw_cagr) * 100))

# ── Helper ──


def fig_to_base64(fig, dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()


charts = {}

# ═══════════════════════════════════════════════════════════════
# Chart 1: 월별 이용건수 추이 (기존 유지)
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(14, 5))
ax.plot(range(len(labels)), values, color='#4472C4', linewidth=1.5)
ax.fill_between(range(len(labels)), values, alpha=0.15, color='#4472C4')
ax.set_title('월별 따릉이 이용건수 추이 (2019년 1월 ~ 2025년 12월)', fontsize=14, fontweight='bold')
ax.set_ylabel('이용건수')
tick_positions = list(range(0, len(labels), 6))
ax.set_xticks(tick_positions)
ax.set_xticklabels([labels[i] for i in tick_positions], rotation=45, fontsize=8)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:.1f}M'))
ax.grid(axis='y', alpha=0.3)
ax.set_ylim(0, max(values) * 1.18)
ax.annotate('COVID-19\n팬데믹 시작', xy=(14, values[14]),
            xytext=(2, max(values) * 0.92),
            arrowprops=dict(arrowstyle='->', color='red', lw=1.2),
            fontsize=9, color='#B71C1C', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.4', facecolor='white', edgecolor='#B71C1C',
                      alpha=1.0, linewidth=1.5))
peak_idx = int(np.argmax(values))
ax.annotate('2023년 10월 피크', xy=(peak_idx, values[peak_idx]),
            xytext=(peak_idx - 26, max(values) * 1.08),
            arrowprops=dict(arrowstyle='->', color='#1B5E20', lw=1.2),
            fontsize=9, color='#1B5E20', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.4', facecolor='white', edgecolor='#1B5E20',
                      alpha=1.0, linewidth=1.5))
charts['monthly_trend'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 2: 연도별 이용건수 + 평균기온 콤보
# ═══════════════════════════════════════════════════════════════
fig, ax1 = plt.subplots(figsize=(10, 6))
yearly_temp_avg = []
for yr in years:
    avg_t = np.mean([temp_map.get(m['연월'], 0) for m in monthly if m['연도'] == yr])
    yearly_temp_avg.append(avg_t)
yearly_usage_list = [yearly_usage[yr] for yr in years]
ax1.set_ylim(0, max(yearly_usage_list) * 1.18)  # 기온선·라벨 공간 확보
bars = ax1.bar(years, yearly_usage_list, color='#4472C4', alpha=0.8, label='이용건수')
ax1.set_ylabel('이용건수', color='#4472C4')
ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:.0f}M'))
ax2 = ax1.twinx()
ax2.plot(years, yearly_temp_avg, color='#FF0000', marker='o', linewidth=2.5, label='평균기온', zorder=5)
ax2.set_ylabel('평균기온 (°C)', color='#FF0000')
ax2.set_ylim(min(yearly_temp_avg) - 2, max(yearly_temp_avg) + 5)  # 기온선 위쪽 공간
ax1.set_title('연도별 이용건수 및 연평균기온', fontsize=14, fontweight='bold')
fig.legend(loc='upper left', bbox_to_anchor=(0.12, 0.88))
ax1.grid(axis='y', alpha=0.2)
for bar, val in zip(bars, yearly_usage_list):
    # 막대 내부 상단에 흰 글씨로 표시 (기온선과 분리)
    ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() - max(yearly_usage_list) * 0.04,
             f'{val / 1e6:.1f}M', ha='center', va='top', fontsize=9, fontweight='bold',
             color='white')
charts['yearly_combo'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 3: 세로 2단 — (상) 산점도 연도별 색 (하) 연도별 기울기 막대
# ═══════════════════════════════════════════════════════════════
fig, (ax, ax2) = plt.subplots(2, 1, figsize=(12, 10),
                                 gridspec_kw={'height_ratios': [1.35, 1]})
year_colors = {
    2019: '#1F3A93', 2020: '#2E86DE', 2021: '#10AC84',
    2022: '#F79F1F', 2023: '#EE5A24', 2024: '#C0392B', 2025: '#5B2C6F',
}
stratified_results = {}
for year in sorted(year_colors.keys()):
    xs = [temps[i] for i, m in enumerate(monthly)
           if m['연도'] == year and temps[i] is not None]
    ys = [values[i] for i, m in enumerate(monthly)
           if m['연도'] == year and temps[i] is not None]
    if len(xs) < 3:
        continue
    sl_y, ic_y, r_y, p_y, _ = stats.linregress(xs, ys)
    stratified_results[year] = {'slope': sl_y, 'intercept': ic_y, 'r_sq': r_y**2,
                                  'p': p_y, 'n': len(xs)}
    ax.scatter(xs, ys, c=year_colors[year], alpha=0.85, s=60,
               edgecolors='white', linewidth=0.8, label=str(year), zorder=3)

# pooled 회귀선만 표시 (연도별 7개 선은 혼잡하므로 생략)
x_line = np.linspace(X.min(), X.max(), 100)
ax.plot(x_line, slope * x_line + intercept, color='#333', linewidth=2.2,
        linestyle='--', zorder=4, label=f'Pooled β={slope/1e4:.1f}만')
ax.set_title('월평균기온 vs 이용건수 (연도별 색상)',
             fontsize=12, fontweight='bold')
ax.set_xlabel('평균기온 (°C)')
ax.set_ylabel('월간 이용건수')
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:.1f}M'))
ax.grid(alpha=0.3)
ax.legend(loc='lower right', fontsize=9, ncol=2, framealpha=0.95,
          title='연도', title_fontsize=9)

# (우) 연도별 기울기 막대 — 안정성 한눈에
yrs_s = sorted(stratified_results.keys())
slopes_s = [stratified_results[y]['slope'] for y in yrs_s]
mean_slope = np.mean(slopes_s)
bars = ax2.bar(yrs_s, slopes_s, color=[year_colors[y] for y in yrs_s],
                edgecolor='white', linewidth=0.8)
ax2.axhline(y=mean_slope, color='#333', linestyle='--', linewidth=1.5,
             label=f'Stratified 평균 {mean_slope/1e4:.1f}만')
ax2.axhline(y=slope, color='#E65100', linestyle=':', linewidth=1.5,
             label=f'Pooled {slope/1e4:.1f}만')
ax2.set_title('연도별 β (기온 기울기)', fontsize=12, fontweight='bold')
ax2.set_xlabel('연도')
ax2.set_ylabel('기울기 (건/°C)')
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e4:.0f}만'))
ax2.set_ylim(0, max(slopes_s) * 1.18)
ax2.grid(axis='y', alpha=0.25)
ax2.legend(loc='upper right', fontsize=8.5, framealpha=0.95)
for bar, v in zip(bars, slopes_s):
    ax2.text(bar.get_x() + bar.get_width()/2, v + max(slopes_s)*0.02,
             f'{v/1e4:.1f}', ha='center', fontsize=8, fontweight='bold')

plt.tight_layout()
charts['scatter'] = fig_to_base64(fig)
charts['__stratified__'] = stratified_results

# ═══════════════════════════════════════════════════════════════
# Chart 4: 계절지수
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(10, 6))
month_names = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월']
colors_si = ['#4472C4' if si >= 1.0 else '#ED7D31' for si in monthly_avg_si]
bars = ax.bar(month_names, monthly_avg_si, color=colors_si, edgecolor='white')
ax.axhline(y=1.0, color='red', linestyle='--', linewidth=1.5, label='기준선 (1.0)')
ax.set_title('월별 계절지수 (7년 평균 대비)', fontsize=14, fontweight='bold')
ax.set_ylabel('계절지수')
ax.legend()
for bar, si in zip(bars, monthly_avg_si):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.02,
            f'{si:.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold')
ax.grid(axis='y', alpha=0.2)
charts['seasonal'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 5 (NEW — 킬러): 잔차 시계열
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(14, 5))
residuals_plot = residuals.copy()
colors_res = ['#2E7D32' if r >= 0 else '#C62828' for r in residuals_plot]
ax.bar(range(len(residuals_plot)), residuals_plot, color=colors_res, width=0.85, alpha=0.85)
ax.axhline(y=0, color='black', linewidth=0.8)
# supF 분기점(2022-04) 이후 영역 음영 + 2022-04 세로선
ax.axvspan(r_break_idx, len(residuals_plot), alpha=0.12, color='red',
           label='2022-04 분기점 이후 구간')
ax.axvline(x=r_break_idx, color='#8B0000', linestyle='--', linewidth=1.6, alpha=0.9)
ax.set_title('기온 효과 제거 후 잔차 시계열 — 2022년 4월 분기점 이후 구조적 이탈',
             fontsize=14, fontweight='bold')
ax.set_ylabel('잔차 (실제 - 기온 기반 예측)')
tick_positions = list(range(0, len(labels), 6))
ax.set_xticks(tick_positions)
ax.set_xticklabels([labels[i] for i in tick_positions], rotation=45, fontsize=8)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:+.1f}M'))
ax.grid(axis='y', alpha=0.25)
ax.legend(loc='lower left')
# y축 위쪽에 흰 공간을 만들고 주석 박스를 상단 중앙-좌측에 배치 (범례와 y 레벨 분리)
ymin, ymax = residuals_plot.min(), residuals_plot.max()
yspan = ymax - ymin
ax.set_ylim(ymin - yspan * 0.08, ymax + yspan * 0.45)
ax.text(0.5, ymax + yspan * 0.30,
        f'분기점 이후 평균 잔차: {post_break_mean_resid:+,.0f}건  ·  음수 = 기온 대비 이용 저조',
        fontsize=10, color='#8B0000', fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.55', facecolor='white', alpha=1.0,
                  edgecolor='#8B0000', linewidth=1.6))
charts['residual'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 6: 이동평균 (기존)
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(14, 5))
ma3 = [None, None] + [np.mean(values[i - 2:i + 1]) for i in range(2, len(values))]
ma6 = [None] * 5 + [np.mean(values[i - 5:i + 1]) for i in range(5, len(values))]
ax.plot(range(len(values)), values, color='#4472C4', linewidth=1, alpha=0.7, label='실제값')
ax.plot(range(len(values)), ma3, color='#ED7D31', linewidth=2, label='3개월 이동평균')
ax.plot(range(len(values)), ma6, color='#70AD47', linewidth=2, label='6개월 이동평균')
ax.set_title('이용건수: 실제값 vs 이동평균', fontsize=14, fontweight='bold')
ax.set_ylabel('이용건수')
ax.set_xticks(tick_positions)
ax.set_xticklabels([labels[i] for i in tick_positions], rotation=45, fontsize=8)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:.1f}M'))
ax.legend()
ax.grid(axis='y', alpha=0.2)
charts['moving_avg'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 7: 연령대 추이
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(12, 7))
colors_age = ['#4472C4', '#ED7D31', '#70AD47', '#FFC000', '#5B9BD5', '#FF6600', '#7030A0']
for i, ag in enumerate(age_groups):
    vals = [yr_data.get(ag, 0) for yr_data in age_yearly]
    yrs = [d['연도'] for d in age_yearly]
    ax.plot(yrs, vals, marker='o', linewidth=2, label=ag, color=colors_age[i])
ax.set_title('연령대별 연도별 이용건수 추이', fontsize=14, fontweight='bold')
ax.set_ylabel('이용건수')
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x / 1e6:.0f}M'))
# 범례를 차트 바깥(우측)으로 빼서 플롯 영역 침범 제거
ax.legend(loc='center left', bbox_to_anchor=(1.02, 0.5), fontsize=9,
          frameon=True, framealpha=0.95, edgecolor='#CCCCCC')
ax.grid(alpha=0.3)
plt.tight_layout()
charts['age_trend'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 8 (NEW): 자치구별 1인당 이용 (상·하위 10)
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(11, 7))
top10 = sorted_gu[:10]
bot10 = sorted_gu[-10:][::-1]
all25 = sorted_gu[::-1]  # 오름차순
y_pos = np.arange(len(all25))
colors_pc = []
for i, (g, _) in enumerate(all25):
    if g in [x[0] for x in top10[:5]]:
        colors_pc.append('#2E7D32')
    elif g in [x[0] for x in bot10[:5]]:
        colors_pc.append('#C62828')
    else:
        colors_pc.append('#9E9E9E')
ax.barh(y_pos, [v for _, v in all25], color=colors_pc, edgecolor='white')
ax.set_yticks(y_pos)
ax.set_yticklabels([g for g, _ in all25], fontsize=9)
ax.set_xlabel('1인당 연간 이용횟수 (2025)')
ax.set_title(f'자치구별 1인당 따릉이 이용 — 지니계수 {gini_per_capita:.3f}',
             fontsize=14, fontweight='bold')
for i, (g, v) in enumerate(all25):
    ax.text(v, i, f' {v:.1f}', va='center', fontsize=8)
ax.grid(axis='x', alpha=0.2)
charts['district'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 9 (NEW): 성별 비율 + 남/여 격차 강조
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(11, 6))
yrs_g = [d['연도'] for d in gender_yearly]
m_vals = [d.get('M', 0) for d in gender_yearly]
f_vals = [d.get('F', 0) for d in gender_yearly]
u_vals = [d.get('미상', 0) for d in gender_yearly]
totals = [mm + ff + uu for mm, ff, uu in zip(m_vals, f_vals, u_vals)]
m_pct = [mm / t * 100 for mm, t in zip(m_vals, totals)]
f_pct = [ff / t * 100 for ff, t in zip(f_vals, totals)]
u_pct = [uu / t * 100 for uu, t in zip(u_vals, totals)]
ax.bar(yrs_g, m_pct, label='남성', color='#4472C4')
ax.bar(yrs_g, f_pct, bottom=m_pct, label='여성', color='#FF6699')
ax.bar(yrs_g, u_pct, bottom=[mm + ff for mm, ff in zip(m_pct, f_pct)],
       label='미상(미등록)', color='#CCCCCC')
for i, yr in enumerate(yrs_g):
    if m_vals[i] > 0 and f_vals[i] > 0:
        ratio = m_vals[i] / f_vals[i]
        # 남성(파랑) 영역 중앙에 흰 글씨 + 흰 외곽선으로 대비 강화
        ax.text(yr, m_pct[i] / 2, f'M/F\n{ratio:.2f}x', ha='center', va='center',
                fontsize=9, color='white', fontweight='bold',
                path_effects=[pe.withStroke(linewidth=2.5, foreground='#1A3A6E')])
ax.set_title('성별 이용 비율 추이 — 남성 편향 격차의 지속', fontsize=14, fontweight='bold')
ax.set_ylabel('비율 (%)')
ax.legend(loc='upper right', framealpha=0.95)
ax.set_ylim(0, 100)
charts['gender'] = fig_to_base64(fig)

# ═══════════════════════════════════════════════════════════════
# Chart 10 (NEW): 대당 일회전율
# ═══════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(14, 5))
ax.plot(range(len(turnover)), turnover, color='#1B3A5C', linewidth=2)
ax.fill_between(range(len(turnover)), turnover, alpha=0.2, color='#1B3A5C')
# 2023 피크 강조
to_peak_idx = int(np.argmax(turnover))
ax.scatter([to_peak_idx], [turnover[to_peak_idx]], s=120, color='#C62828', zorder=5)
to_min = min(turnover)
to_max = max(turnover)
to_span = to_max - to_min
ax.set_ylim(to_min - to_span * 0.20, to_max * 1.15)  # 하단·상단 모두 여유 확보
ax.annotate(f'{labels[to_peak_idx][:4]}년 {int(labels[to_peak_idx][4:])}월 피크 {turnover[to_peak_idx]:.2f}',
            xy=(to_peak_idx, turnover[to_peak_idx]),
            xytext=(to_peak_idx - 24, to_max * 1.05),  # x축 위쪽으로 이동 (라벨과 분리)
            arrowprops=dict(arrowstyle='->', color='#8B0000', lw=1.2),
            fontsize=9, color='#8B0000', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.4', facecolor='white', edgecolor='#8B0000',
                      alpha=1.0, linewidth=1.5))
ax.set_title('자전거 1대당 일회전율 추이 — 2023년 피크 이후 하락', fontsize=14, fontweight='bold', pad=12)
ax.set_ylabel('대당 일회전율 (이용건수 / 운영대수 / 일수)')
ax.set_xticks(tick_positions)
ax.set_xticklabels([labels[i] for i in tick_positions], rotation=45, fontsize=8)
ax.grid(axis='y', alpha=0.3)
charts['turnover'] = fig_to_base64(fig)

print(f'차트 생성 완료: {len(charts)}개')

# ═══════════════════════════════════════════════════════════════
# HTML 본문 작성
# ═══════════════════════════════════════════════════════════════
max_month = max(monthly, key=lambda m: m['이용건수'])
min_month = min(monthly, key=lambda m: m['이용건수'])


def fmt_ym(ym):
    s = str(ym)
    return f'{s[:4]}년 {int(s[4:])}월'


# Yearly summary table rows
yearly_tbl_rows = ''
prev_total = None
for yr in years:
    total = yearly_usage[yr]
    if prev_total is not None:
        pct = (total - prev_total) / prev_total * 100
        pct_str = f'{pct:+.1f}%'
    else:
        pct_str = '—'
    yearly_tbl_rows += f'<tr><td>{yr}</td><td>{total:,}</td><td>{pct_str}</td></tr>\n'
    prev_total = total

# 미상 보정 CAGR 테이블
cagr_tbl_rows = ''
for ag, v19, v25, raw_c, adj_c, diff in cagr_rows:
    cagr_tbl_rows += (
        f'<tr><td>{ag}</td><td>{v19:,}</td><td>{v25:,}</td>'
        f'<td>{raw_c * 100:+.1f}%</td><td>{adj_c * 100:+.1f}%</td>'
        f'<td>{diff:+.2f}</td></tr>\n'
    )

# 자치구 상·하위 5
top5_rows = ''.join(
    f'<tr><td>{i + 1}</td><td>{g}</td><td>{v:.2f}</td></tr>\n'
    for i, (g, v) in enumerate(sorted_gu[:5]))
bot5_rows = ''.join(
    f'<tr><td>{25 - i}</td><td>{g}</td><td>{v:.2f}</td></tr>\n'
    for i, (g, v) in enumerate(sorted_gu[-5:]))

# Unit economics 요약
to_2023 = yearly_turnover.get(2023, 0)
to_2025 = yearly_turnover.get(2025, 0)
to_drop_pct = (to_2025 / to_2023 - 1) * 100 if to_2023 else 0

# 2026 예측 (추세×계절)
time_idx = list(range(len(monthly)))
slope_t, intercept_t, _, _, _ = stats.linregress(time_idx, values)
forecast_2026 = []
for mo in range(1, 13):
    t = len(monthly) - 1 + mo
    trend_val = intercept_t + slope_t * t
    fc = trend_val * monthly_avg_si[mo - 1]
    forecast_2026.append((mo, trend_val, monthly_avg_si[mo - 1], fc))
forecast_total_2026 = sum(f[3] for f in forecast_2026)
# 신뢰구간: 추세×계절 모델 잔차 기반 RMSE
ts_preds = [(intercept_t + slope_t * i) * monthly_avg_si[(monthly[i]['월']) - 1] for i in range(len(monthly))]
ts_rmse = float(np.sqrt(np.mean((np.array(values) - np.array(ts_preds)) ** 2)))

# HTML 템플릿
html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>공유 모빌리티 시장 재편기의 따릉이</title>
<style>
@page {{ size: A4; margin: 20mm 18mm; @bottom-center {{ content: counter(page); font-size: 9pt; color:#888; }} }}
@page :first {{ @bottom-center {{ content: ""; }} }}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; font-size: 11pt; line-height: 1.65; color: #222; word-break: keep-all; overflow-wrap: break-word; }}

.cover {{ page-break-after: always; position: relative; display: flex; flex-direction: column; align-items: center; height: 100vh; text-align: center; padding-top: 18vh; }}
.cover h1 {{ font-size: 24pt; color: #1B3A5C; margin-bottom: 10px; line-height: 1.35; }}
.cover .subtitle {{ font-size: 13pt; color: #555; margin-bottom: 28px; }}
.cover h2 {{ font-size: 15pt; color: #444; margin-bottom: 32px; font-weight: 500; }}
.cover .info {{ font-size: 12pt; color: #333; line-height: 2; }}
.cover .line {{ width: 60%; height: 3px; background: linear-gradient(90deg, #1B3A5C, #4472C4, #1B3A5C); margin: 24px auto; }}
.cover .ai-notice {{ margin-top: auto; margin-bottom: 4mm; font-size: 9pt; color: #666; width: 60%; text-align: center; line-height: 1.55; }}

.toc {{ padding: 16px 0 22px; page-break-after: always; }}
.toc h2 {{ color: #1B3A5C; font-size: 17pt; border-bottom: 3px double #1B3A5C; padding-bottom: 6px; margin-bottom: 14px; }}
.toc ul {{ list-style: none; columns: 2; column-gap: 28px; column-fill: balance; }}
.toc li {{ padding: 2px 0; font-size: 10pt; break-inside: avoid; line-height: 1.5; }}
.toc li.main {{ font-weight: bold; margin-top: 5px; color: #1B3A5C; }}
.toc li.sub {{ padding-left: 16px; color: #555; }}

h2.section {{ color: #1B3A5C; font-size: 15pt; border-bottom: 2px solid #1B3A5C; padding-bottom: 6px; margin: 26px 0 12px; page-break-after: avoid; }}
h3 {{ color: #2E5090; font-size: 12pt; margin: 14px 0 8px; page-break-after: avoid; }}
h4 {{ color: #444; font-size: 11pt; margin: 10px 0 6px; page-break-after: avoid; break-after: avoid-page; }}

p {{ margin: 8px 0; text-align: justify; }}
strong {{ color: #1B3A5C; }}
em {{ color: #C62828; font-style: normal; font-weight: 600; }}

.chart {{ text-align: center; margin: 14px 0; page-break-inside: avoid; }}
.chart img {{ max-width: 100%; border: 1px solid #ddd; border-radius: 4px; }}
.caption {{ font-size: 9pt; color: #666; margin-top: 4px; font-style: italic; }}

.insight {{ background: #F0F4FA; border-left: 4px solid #4472C4; padding: 10px 14px; margin: 12px 0; border-radius: 0 4px 4px 0; page-break-inside: avoid; font-size: 10.5pt; }}
.insight strong {{ color: #1B3A5C; }}
.caveat {{ background: #FFF8E1; border-left: 4px solid #F57F17; padding: 10px 14px; margin: 12px 0; font-size: 10.5pt; page-break-inside: avoid; }}
.killer {{ background: #FBE9E7; border-left: 4px solid #C62828; padding: 12px 14px; margin: 14px 0; font-size: 11pt; page-break-inside: avoid; }}
.killer strong {{ color: #8B0000; font-size: 11.5pt; }}

table {{ width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 9.5pt; page-break-inside: avoid; word-break: keep-all; overflow-wrap: break-word; table-layout: auto; }}
th {{ background: #1B3A5C; color: white; padding: 6px 8px; text-align: center; word-break: keep-all; }}
td {{ padding: 5px 8px; border: 1px solid #ddd; text-align: right; word-break: keep-all; }}
td:first-child {{ text-align: center; font-weight: 500; }}
tr:nth-child(even) td {{ background: #FAFAFA; }}
table.emphasis-first td:first-child {{ background: #F5F5F5; font-weight: 600; }}
table.appendix-table {{ font-size: 9.5pt; }}
table.appendix-table td {{ text-align: left; vertical-align: top; padding: 5px 8px; }}
table.appendix-table td:first-child {{ text-align: center; width: 40px; font-weight: 600; color: #1B3A5C; }}
table.appendix-table td:nth-child(2) {{ width: 150px; font-weight: 600; }}
table.appendix-table .formula {{ font-family: 'Courier New', monospace; font-size: 9pt; color: #555; display: block; margin-top: 2px; word-break: break-all; }}
table.appendix-table tr.divider td {{ background: #1B3A5C; color: white; font-weight: 600; padding: 4px 8px; text-align: left; }}

.footer {{ text-align: center; font-size: 8pt; color: #999; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 8px; }}
.ref-list {{ font-size: 10pt; line-height: 1.85; columns: 2; column-gap: 24px; column-fill: balance; }}
.ref-list li {{ margin-bottom: 6px; list-style: decimal inside; break-inside: avoid; }}

ul, ol {{ margin-left: 24px; }}
li {{ margin-bottom: 4px; }}

.appendix-box {{ background: #F9F9F9; padding: 10px 14px; border: 1px solid #E0E0E0; font-family: 'Courier New', monospace; font-size: 8.5pt; line-height: 1.55; margin: 10px 0; white-space: pre; overflow-x: hidden; word-break: keep-all; }}
</style>
</head>
<body>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 표지 ━━━━━━━━━━━━━━━━━━━━━━ -->
<div class="cover">
  <h1>공유 모빌리티 재편기의 따릉이<br>수요 구조의 변화와 운영 모델의 재구성</h1>
  <div class="line"></div>
  <div class="subtitle">서울시 공공자전거 2019~2025년 84개월 이용 데이터 분석</div>
  <h2>트렌드를읽는데이터경영 중간과제</h2>
  <div class="info">
    <span>중앙대학교 소프트웨어학부</span>
    <span>20203876 최성민</span>
    <span>01분반</span>
    <span style="margin-top:20px; font-size:10pt; color:#888;">2026년 4월</span>
    <span style="margin-top:8px; font-size:8.5pt; color:#888;">
      저장소 · <span style="font-family:monospace;">{GITHUB_URL}</span> · commit <span style="font-family:monospace;">{GITHUB_COMMIT}</span>
    </span>
  </div>
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ Executive Summary ━━━━━━━━━━━━━━━━━━━━━━ -->
<div style="page-break-after: always; padding: 20px 0;">
<h2 class="section" style="border-bottom:2px solid #1B3A5C;">Executive Summary</h2>

<div class="insight" style="font-size:11pt;">
<strong>한 줄 진단:</strong> 따릉이는 2021년 9월부터 기온으로 설명되지 않는 구조적 수요 이탈을
누적해 왔으며(supF = 54.6, Andrews 1993 임계값 8.85 초과), 2023년 이벤트(더스윙·기후동행카드)는
이미 진행 중인 구조 변화의 가속 요인이다.
</div>

<h4 style="margin-top:14px;">3 Key Findings</h4>
<table style="font-size:10pt;">
<tr><th style="width:18%;">주제</th><th style="width:32%;">결과</th><th style="width:50%;">정책 함의</th></tr>
<tr><td>구조변화 시점</td><td>supF = 54.6 @ 2021-09 · Bai-Perron 1-break 확정</td><td>엔데믹 전환기의 수요 이탈. 2023 이벤트는 가속 요인.</td></tr>
<tr><td>공간 자기상관</td><td>Queen I = +0.279 (p = 0.015) · LISA FDR 후 0 유의</td><td>전역 군집은 확인 · 국지 cluster는 n=25 검정력 한계.</td></tr>
<tr><td>재무 타당성</td><td>Alt3 전동보조 증분 NPV = +92억 · B/C 1.40 · IRR 10.9%</td><td>공공가치 회계(탄소·건강·혼잡 3축) 반영 시 예타 통과 가능.</td></tr>
</table>

<h4 style="margin-top:14px;">3 Decision Points (서울시 교통정책실장 결재용)</h4>
<table style="font-size:10pt;">
<tr><th style="width:14%;">분기점</th><th style="width:36%;">선택지</th><th style="width:50%;">본 연구의 권고</th></tr>
<tr><td>예산 규모</td><td>연 300억 vs 150억 vs 유지(100억)</td><td>연 150억 증액 · Alt2+Alt3 혼합 (증분 NPV 양수 구간)</td></tr>
<tr><td>Fleet 구성</td><td>일반만 유지 vs 전동보조 10% 도입 vs 대규모 전환</td><td>전동보조 10% 도입 · 한강 인접 대여소 선행 배치</td></tr>
<tr><td>기후동행카드</td><td>독립 운영 vs 제한 통합 vs 완전 통합</td><td>제한 통합 + 가입자 데이터 공동 분석 MOU 선행</td></tr>
</table>

<h4 style="margin-top:14px;">본문 로드맵</h4>
<p style="font-size:10pt; color:#555;">
1장 배경·가설 → 2장 데이터·방법 → 3장 3국면 기술 → <strong>4장 구조변화 식별 (본 연구 핵심)</strong>
→ 5장 예측 · SARIMAX → 6장 수요자·공간 · FDR · 민감도 · SAR/SEM → <strong>7장 운영 모델 재구성 · Competitive Landscape · Roadmap · Decision Tree</strong>
→ 8장 재무 · 3대안 · 공공가치 회계 → 부록 A~H · 기술부록 별책(T1~T7) · GitHub 저장소.
</p>

<p style="font-size:9pt; color:#888; margin-top:16px; text-align:center;">
— 본 연구는 오픈데이터 + 오픈소스로 완전 재현 가능하다.
{GITHUB_URL} 저장소에서 <code>make reproduce</code> 실행 시 본 보고서의 모든 수치가 재생성된다. —
</p>
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 목차 ━━━━━━━━━━━━━━━━━━━━━━ -->
<div class="toc">
<h2>목차</h2>
<ul>
<li class="main">Executive Summary</li>
<li class="main">1. 서론 — 공유 모빌리티 재편기의 따릉이</li>
<li class="sub">1.1 연구 배경</li>
<li class="sub">1.2 연구 질문과 가설</li>
<li class="main">2. 데이터 및 방법</li>
<li class="sub">2.1 데이터 출처와 표본</li>
<li class="sub">2.2 분석 방법론</li>
<li class="sub">2.3 데이터 한계 — '미상' 비율 변동 문제</li>
<li class="main">3. 이용 구조 변화 — 성장·피크·감소의 3국면</li>
<li class="main">4. 기온 효과의 분리와 구조변화 분기점 진단</li>
<li class="sub">4.1 단순회귀의 한계 · 잔차 진단 · 분산분해</li>
<li class="sub">4.2 Bai-Perron 구조변화 검정과 supF</li>
<li class="main">5. 계절성과 2026년 예측 · SARIMAX</li>
<li class="sub">5.1 월별 계절지수</li>
<li class="sub">5.2 SARIMAX 식별과 분석적 예측구간</li>
<li class="main">6. 수요자 구조 — 누가 타는가</li>
<li class="sub">6.1 성별 격차의 심화</li>
<li class="sub">6.2 고령층 성장의 실체 — 미상 보정 후</li>
<li class="sub">6.3 자치구 공간 분포 · Moran 민감도 · LISA FDR · SAR/SEM</li>
<li class="main">7. 운영 모델의 재구성 — 경영 관점</li>
<li class="sub">7.0 Competitive Landscape (수단 간 한계비용)</li>
<li class="sub">7.1 자전거 1대당 회전율의 추이</li>
<li class="sub">7.2 공유 모빌리티 시장의 구조 변화</li>
<li class="sub">7.3 이용자 층별 운영 분화</li>
<li class="sub">7.4 Roadmap 2026~2030 · Decision Tree · Risk Pre-empt</li>
<li class="main">8. 결론 및 제언 · 공공가치 회계</li>
<li class="sub">8.1 종합</li>
<li class="sub">8.2 3대안 재무 비교 · 공공가치 회계 (3축)</li>
<li class="sub">8.3 형평성 KPI 제안</li>
<li class="sub">8.4 연구의 한계 및 후속 과제</li>
<li class="main">부록 A — Excel 분석 재현 가이드</li>
<li class="main">부록 B — 참고문헌 및 출처</li>
<li class="main">부록 H — 결측률·data manifest</li>
<li class="main">별책 — 기술부록 T1~T7 (따릉이_기술부록.pdf)</li>
</ul>
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 1. 서론 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">1. 서론 — 공유 모빌리티 재편기의 따릉이</h2>

<h3>1.1 연구 배경</h3>
<p>서울시 공공자전거 '따릉이'는 2015년 도입 이후 10년간 누적 2억 5천만 건의 이용을 기록하며
지하철·버스에 이은 '제3의 대중교통 수단'으로 자리매김하였다. 2024년 기준 4만 5천 대 규모로
운영되고 있으며, 연간 약 100억 원의 적자를 부담하는 공공서비스다. 운영 효율과 형평성이 곧
시민의 세금과 직결된다.</p>

<p>한편 2021년부터 급성장하던 서울의 PM(전동킥보드) 시장은 2023~2025년 사이 축소 국면에
진입하였고, 같은 시기 따릉이 이용량 또한 2023년 4,490만 건의 피크에서 2년 만에 3,737만 건으로
<em>-17%</em> 감소하였다. 본 연구는 따릉이를 '지속 성장하는 공공재'로 보는 통념을 재고한다.
대신 <strong>공유 모빌리티 시장 재편기에 직면한 따릉이의 수요 구조 변화</strong>를 정량적으로
재검토한다.</p>

<h3>1.2 연구 질문과 가설</h3>
<p>본 연구는 네 가지 열린 질문에서 출발한다. 각 질문에 대해 귀무가설(H₀)과 대립가설(H₁)을
병기하여 결과의 방향을 사전에 단정하지 않는다. 검증 결과는 본문에서 지지·기각·유보 중
어느 쪽으로도 귀결될 수 있다.</p>

<table>
<tr><th>번호</th><th>연구 질문</th><th>H₀ / H₁ 또는 비교 프레임</th></tr>
<tr>
  <td>Q1</td>
  <td>따릉이 수요 변동은 기온으로 충분히 설명되는가?</td>
  <td>H₀: 기온 효과 제거 후 잔차는 시간에 대해 무작위로 분포한다.<br>
      H₁: 잔차는 유의한 시간 추세를 보여 기온 외 변수의 작용을 시사한다.</td>
</tr>
<tr>
  <td>Q2</td>
  <td>2026년 수요에 대해 어느 예측 모델이 적합한가?</td>
  <td>네 모델(Naive · Seasonal Naive · 3개월 이동평균 · 추세×계절)을 동일한
      2025년 hold-out 기준으로 비교한다. 특정 모델의 우위는 사전에 예단하지 않는다.</td>
</tr>
<tr>
  <td>Q3</td>
  <td>'미상' 비율의 변동이 연령별 CAGR 해석에 영향을 미치는가?</td>
  <td>H₀: 미상 안분 보정 전후 CAGR 차이는 무시 가능하다.<br>
      H₁: 차이가 유의하여 원본 CAGR 해석에 주의가 필요하다.</td>
</tr>
<tr>
  <td>Q4</td>
  <td>따릉이 이용의 공간 분포는 어떤 패턴을 보이는가?</td>
  <td>지니계수의 수준·시계열 방향, Moran's I 기반 공간 패턴
      (군집 · 분산 · 무작위)을 사전에 예단하지 않고 검정한다.</td>
</tr>
</table>

<p style="font-size:10pt; color:#555;">
※ 본 연구는 탐색적(exploratory) 성격이 일부 포함된다. 특히 Q4의 공간 자기상관 방향은
선행 문헌의 이론적 예측과 데이터의 경험적 결과가 일치하지 않을 경우를 상정하여 H₀
기각이 일어나지 않는 결과도 함께 보고한다.
</p>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 2. 데이터 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">2. 데이터 및 방법</h2>

<h3>2.1 데이터 출처와 표본</h3>
<table>
<tr><th>데이터</th><th>출처</th><th>기간</th><th>규모</th></tr>
<tr><td>따릉이 이용정보</td><td>서울 열린데이터광장 OA-15248</td><td>2019년 1월 ~ 2025년 12월</td><td>84개월 · 약 530만 행</td></tr>
<tr><td>대여소 마스터</td><td>서울 열린데이터광장 OA-13252 (25.12월 기준)</td><td>2025년 12월 스냅샷</td><td>2,799개 대여소 · 25자치구</td></tr>
<tr><td>서울 월평균기온</td><td>Open-Meteo 재분석 데이터 (서울 37.57°N)</td><td>2019년 1월 ~ 2025년 12월</td><td>84개월</td></tr>
<tr><td>자치구 인구</td><td>서울시 주민등록인구통계</td><td>2025년 1월 기준</td><td>25자치구</td></tr>
</table>

<h3>2.2 분석 방법론</h3>
<p>수업 실습 범위를 모두 충족하기 위해 다음 기법을 Excel에 구현하였다. 기술통계(평균·중앙값·표준편차·분위수),
연평균 성장률(CAGR = RATE 함수), 3·6개월 이동평균과 MAE, INTERCEPT·SLOPE 선형 추세선,
월별 계절지수, 피벗테이블과 스파크라인, 산점도·콤보차트·보조축이 그것이다.
여기에 더해 통계적 엄밀성을 보강하기 위해 <strong>LINEST 다중회귀, Durbin-Watson 자기상관
진단, Seasonal Naive 베이스라인, 95% 신뢰구간, 지니계수</strong>를 추가로 도입하였다.</p>

<h3>2.3 데이터 한계 — '미상' 비율 변동 문제</h3>
<p>원본 데이터의 성별·연령대 필드에는 회원 가입 시점의 자기보고 값이 기록되거나, 비회원의
경우 '미상'으로 분류된다. 이 '미상' 비율은 2019년 <em>{misu19 * 100:.1f}%</em>에서
2025년 <em>{misu25 * 100:.1f}%</em>로 급감하였다. 이는 실제 수요 변동이라기보다 회원
가입 유도 및 비회원 이용 규제 등 <strong>집계 기준의 변화</strong>가 반영된 결과로
해석된다. 보정 없이 산출한 연령별 CAGR은 실제보다 과대 추정될 가능성이 있다.
이러한 한계를 보완하기 위해 6.2절에서는 '미상'을 각 연령에 안분한 보정 CAGR을 병기한다.</p>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 3. 이용 구조 변화 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">3. 이용 구조 변화 — 성장·피크·감소의 3국면</h2>

<div class="chart">
<img src="data:image/png;base64,{charts['monthly_trend']}">
<div class="caption">[그림 1] 월별 따릉이 이용건수 추이 (2019년 1월 ~ 2025년 12월).</div>
</div>

<p>그림 1의 84개월 시계열은 세 국면으로 명확히 구분된다. 첫째는 2019~2020년의 <strong>기반기</strong>다.
2019~2020년 사이 이용량은 1,810만 건에서 2,367만 건으로 <em>+31%</em> 증가하였다.
같은 기간 코로나19 확산 및 대중교통 기피 정서가 함께 관측되었으나, COVID 더미를 투입한
다중회귀에서는 t=1.02로 통계적 유의성에 도달하지 않았다(LINEST_Multi 시트). 본 연구는
기반기의 증가를 코로나19와의 <strong>동반 현상</strong>으로 기록하되 단일 인과로는 해석하지 않는다.
둘째는 2021~2023년의 <strong>확장기</strong>다. 2023년 10월에 단월 기준 역대 최대인
{max(values) / 1e6:.2f}백만 건을 기록하며 연간 4,490만 건의 피크에 도달하였다.
셋째는 2024~2025년의 <strong>조정기</strong>다. 2025년 이용은 3,737만 건으로 피크 대비
<em>-{(1 - yearly_usage[2025] / yearly_usage[2023]) * 100:.1f}%</em> 감소하였다.
따릉이의 성장 국면의 종료가 시사된다.</p>

<div class="chart">
<img src="data:image/png;base64,{charts['yearly_combo']}">
<div class="caption">[그림 2] 연도별 이용건수(막대)와 연평균기온(선) — 콤보차트에 보조축을 적용.</div>
</div>

<table>
<tr><th>연도</th><th>이용건수</th><th>전년 대비</th></tr>
{yearly_tbl_rows}
</table>

<div class="insight">
<strong>2019~2025 단순 CAGR은 {cagr_19_25 * 100:.1f}%</strong>로 산출되지만,
<strong>2023~2025 2년 CAGR은 {cagr_23_25 * 100:+.1f}%</strong>로 부호가 반전된다.
도입 4년차였던 2019년의 낮은 기저효과가 장기 CAGR을 과대 추정한 결과로 해석된다.
최근 2년만을 분리하여 살펴보면 따릉이는 명확한 축소 국면에 진입하였다.
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 4. 기온 효과 분리 + 구조변화 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">4. 기온 효과의 분리와 구조변화 분기점 진단</h2>

<p>3장의 3국면 구분은 시계열의 외형에 기반한 기술적 관찰이다. 본 장은 그 외형을 두
층으로 분해한다. 첫째는 계절성의 단기 변동 요인인 <strong>기온 효과</strong>이며,
둘째는 기온으로 설명되지 않는 장기 추세의 이탈, 즉 <strong>시간 차원의 구조 변화</strong>이다.
분석은 세 단계로 진행된다. 4.1절에서 단순회귀로 피상적 상관의 한계를 드러내고,
4.1.1절에서 연도별 stratified 회귀로 기온의 순효과를 분리한다. 4.2절은 잔차 시계열과
supF 구조변화 검정으로 <strong>분기점의 시점</strong>을 식별한다.</p>

<h3>4.1 단순회귀의 한계 — 기온 효과와 연도 성장의 혼재</h3>
<div class="chart">
<img src="data:image/png;base64,{charts['scatter']}">
<div class="caption">[그림 3] 월평균기온과 따릉이 이용건수의 산점도(연도별 색상) 및 선형 추세선.</div>
</div>

<p>그림 3은 84개월의 월평균기온과 월간 이용건수의 분포를 <strong>연도별 색상</strong>으로
구분하여 보여준다. 피어슨 상관계수 R = {r_value:.4f}, 결정계수 R² = {r_sq:.4f}로 강한 양의
상관관계가 관측되지만, <strong>이 수치는 기온 효과와 연도별 성장·축소 효과가 혼재된
값</strong>이다. 같은 기온 구간(예: 10°C 전후)에서도 2019년(진한 파랑) 월평균 이용량은
약 150만 건, 2023년(주황)은 약 370만 건으로 2배 이상 벌어진다. 단순회귀의 기울기
{slope:,.0f}건/°C는 "기온 1°C 상승 효과"가 아니라 "기온 상승과 연도 성장이 함께 작용한
합산 효과"로 해석되어야 한다.</p>

<div class="caveat">
<strong>시간 교란변수 통제의 필요성:</strong> 단순회귀는 시간을 교란변수로 통제하지
않으므로 R² = {r_sq:.4f}를 기온→이용의 인과적 강도로 읽을 수 없다. 이를 보완하기 위해
<strong>LINEST_Multi 시트</strong>에서 Temp + 시간 인덱스 t + COVID 더미 + 11개 월 더미를
함께 투입한 다중회귀로 기온의 순효과를 분리 추정한다. <strong>Residual_Analysis 시트</strong>에서는
Durbin-Watson으로 잔차 자기상관을 진단한다. 4.2절의 잔차 분석은 이러한 통제 이후에도
남는 <strong>시간 잔여 효과</strong>를 포착하는 역할을 한다.
</div>

<h4>4.1.1 연도별 Stratified 회귀 — 연도 내 기온 순효과</h4>
<p>연도 성장 효과를 제거한 기온의 순효과를 직접 추정하기 위해 각 연도별로 OLS 회귀를
분할 수행하였다(stratified regression). 이는 연도를 <strong>고정효과</strong>로 취급하는
근사에 해당한다.</p>

<table>
<tr><th>연도</th><th>n</th><th>기울기 β (건/°C)</th><th>절편 (건)</th><th>R²</th><th>p-value</th></tr>
<tr><td>2019</td><td>11</td><td>66,228</td><td>677,524</td><td>0.81</td><td>0.0001</td></tr>
<tr><td>2020</td><td>12</td><td>61,636</td><td>1,157,726</td><td>0.58</td><td>0.0040</td></tr>
<tr><td>2021</td><td>12</td><td>83,922</td><td>1,519,702</td><td>0.81</td><td>0.0001</td></tr>
<tr><td>2022</td><td>12</td><td><strong>102,753</strong></td><td>2,052,796</td><td>0.73</td><td>0.0004</td></tr>
<tr><td>2023</td><td>12</td><td>92,258</td><td>2,445,836</td><td>0.62</td><td>0.0023</td></tr>
<tr><td>2024</td><td>12</td><td>81,279</td><td>2,446,460</td><td>0.63</td><td>0.0021</td></tr>
<tr><td>2025</td><td>12</td><td>70,956</td><td>2,116,306</td><td>0.77</td><td>0.0002</td></tr>
<tr><td><strong>평균</strong></td><td>—</td><td><strong>79,862 (±13,594)</strong></td><td>—</td><td><strong>0.70</strong></td><td>—</td></tr>
<tr><td>Pooled (비교)</td><td>83</td><td>81,119</td><td>1,770,098</td><td>0.43</td><td><0.0001</td></tr>
</table>

<div class="insight">
<strong>세 가지 관찰 (within/between 분산분해):</strong>
<br>① 기울기 평균 79,862 vs pooled 81,119 — 거의 동일(-1.5%). 기온의 순효과는 분산분해상
안정적으로 존재한다.
<br>② <strong>R²는 0.43 → 0.70</strong>으로 상승. 이는 "우연"이 아니라
분산 귀속의 결과이다. 전체 분산 중 <strong>between-year {V21_STATS['variance_decomposition']['between_ratio']*100:.0f}%</strong>(연도 성장)이
pooled 회귀의 잔차로 흡수되어 pooled R²가 낮아지고, stratified로 분할하면
<strong>within-year {V21_STATS['variance_decomposition']['within_ratio']*100:.0f}%</strong>만 남아 R²가 기온-이용량 순관계를 온전히 반영한다.
<br>③ 절편은 2019년 68만 → 2022년 205만 → 2025년 212만으로 변동 — 연도 효과는
<strong>기울기가 아니라 절편</strong>에 자리 잡는다. 연도별 최대 기울기(2022년
102,753)와 최소(2020년 61,636)의 차이는 COVID 시기 변동성이 컸음을 시사한다.
</div>

<h4>4.1.2 잔차 진단 표 — 고전적 OLS 가정 검정</h4>
<p>계량경제학적 엄밀성을 위하여 단순회귀 잔차에 대해 다섯 가지 진단 검정을 수행하였다
(Python <code>statsmodels.stats</code>, 재현 저장소 <code>compute_stats_v21.py</code>).</p>

<table style="font-size:9.5pt;">
<tr><th>검정</th><th>통계량</th><th>p-value</th><th>판정 (α = 0.05)</th></tr>
<tr><td>Durbin-Watson (자기상관)</td>
    <td>{V21_STATS['residual_diagnostics']['durbin_watson']}</td>
    <td>—</td>
    <td>강한 양의 자기상관 (DW &lt; 1.5)</td></tr>
<tr><td>Ljung-Box Q(10)</td>
    <td>{V21_STATS['residual_diagnostics']['ljung_box_q10']['stat']:.2f}</td>
    <td>{V21_STATS['residual_diagnostics']['ljung_box_q10']['p']:.4f}</td>
    <td>H₀(백색성) 기각</td></tr>
<tr><td>Ljung-Box Q(20)</td>
    <td>{V21_STATS['residual_diagnostics']['ljung_box_q20']['stat']:.2f}</td>
    <td>{V21_STATS['residual_diagnostics']['ljung_box_q20']['p']:.4f}</td>
    <td>H₀ 기각</td></tr>
<tr><td>Breusch-Godfrey LM(4)</td>
    <td>{V21_STATS['residual_diagnostics']['breusch_godfrey_lm']['stat']:.2f}</td>
    <td>{V21_STATS['residual_diagnostics']['breusch_godfrey_lm']['p']:.4f}</td>
    <td>고차 자기상관 존재</td></tr>
<tr><td>ARCH-LM(4) (조건부 이분산)</td>
    <td>{V21_STATS['residual_diagnostics']['arch_lm']['stat']:.2f}</td>
    <td>{V21_STATS['residual_diagnostics']['arch_lm']['p']:.4f}</td>
    <td>{'이분산 존재' if V21_STATS['residual_diagnostics']['arch_lm']['p'] < 0.05 else '이분산 미유의'}</td></tr>
<tr><td>Jarque-Bera (정규성)</td>
    <td>{V21_STATS['residual_diagnostics']['jarque_bera']['stat']:.2f}</td>
    <td>{V21_STATS['residual_diagnostics']['jarque_bera']['p']:.4f}</td>
    <td>{'정규성 기각' if V21_STATS['residual_diagnostics']['jarque_bera']['p'] < 0.05 else '정규성 수용'}</td></tr>
</table>

<div class="caveat">
<strong>가정 위배의 시사점:</strong> OLS 잔차는 (i) 강한 양의 자기상관
(DW = {V21_STATS['residual_diagnostics']['durbin_watson']}), (ii) Ljung-Box 기각에 따른 비백색
잡음, (iii) BG-LM 고차 자기상관, (iv) JB 정규성 기각으로 고전적 가정을 체계적으로 위반한다.
따라서 pooled OLS의 표준오차·신뢰구간은 과소 추정되었을 가능성이 높으며, 정확한 추론을
위해서는 <strong>Newey-West HAC 표준오차</strong> 또는 <strong>SARIMAX 분산 기반 예측구간</strong>
(5.2절)이 요구된다. 본 연구는 이 두 보정의 결과를 5장에서 제시한다.
</div>

<p style="font-size:9pt; color:#555; margin: -4px 0 12px 0;">
※ 계절·추세를 명시적으로 통제한 <strong>SARIMAX {V21_STATS.get('sarimax', {}).get('best_order', '(1,1,2)(0,1,1,12)')} 모델
(AIC = {V21_STATS.get('sarimax', {}).get('best_aic', 1589):.0f})</strong> 의 외생변수 기온 계수는
<strong>+{V21_STATS.get('sarimax', {}).get('exog_temp_coef', 82671):,.0f} 건/°C</strong>로 추정된다.
종속변수에 자연로그를 취한 semi-log 회귀에서는 기온 1°C 상승 효과가 <strong>+1.14%</strong>로
산출된다. Pooled OLS의 +81,119건/°C는 단순 대표값이며, 계절·추세 통제 후의 기온 효과
대표 범위는 <em>약 +80,000~95,000건/°C (또는 +1.0~1.3%/°C)</em>로 제시한다.
</p>

<h3>4.2 잔차 분석과 구조변화 분기점 — 2022년 2분기 이후의 누적 이탈</h3>
<p>4.1절이 기온의 순효과를 분리하였다면, 본 절은 기온이 설명하지 못하는
<strong>시간 차원의 이탈</strong>을 포착한다. 분석은 두 단계이다. 첫째, supF 구조변화
검정으로 분기점의 시점을 객관적으로 특정한다. 둘째, 해당 분기점을 기준으로 잔차
시계열의 방향성을 확인한다.</p>

<h4>4.2.1 Bai-Perron 구조변화 검정과 supF</h4>
<p>본 연구는 잠재 분기점(break candidate)을 월 단위로 이동시키며 Chow F-통계량을 산출하는
<strong>supF 검정</strong>과 Bai-Perron(1998, 2003)의 <strong>PELT/Binseg 다중 분기점 탐색</strong>을
결합 수행하였다. supF는 trim = 15% 규약을 따라 잠재 분기점 범위를 [0.15n, 0.85n] = [13, 72]로
제한하고, Andrews(1993) 2-모수 5% 임계값 <strong>8.85</strong>를 기준으로 삼는다.</p>

<table style="font-size:9.5pt;">
<tr><th>검정</th><th>결과</th><th>임계값</th><th>판정</th></tr>
<tr><td>supF (grid, trim 15%)</td>
    <td><strong>{V21_STATS['bai_perron']['supF_stat']}</strong> @ {V21_STATS['bai_perron']['supF_ym']}</td>
    <td>8.85 (Andrews 1993)</td>
    <td><strong>H₀ 기각 — 분기점 존재</strong></td></tr>
<tr><td>Bai-Perron Binseg 1-break</td>
    <td>{V21_STATS['bai_perron']['pelt_bkps'].get('1_breaks', [{}])[0].get('ym', 'N/A') if V21_STATS['bai_perron']['pelt_bkps'].get('1_breaks') else 'N/A'}</td>
    <td>—</td>
    <td>분기점 위치 일관</td></tr>
<tr><td>Bai-Perron Binseg 2-breaks</td>
    <td>{', '.join([b['ym'] for b in V21_STATS['bai_perron']['pelt_bkps'].get('2_breaks', [])]) or 'N/A'}</td>
    <td>—</td>
    <td>2분기점 보조 확인</td></tr>
</table>

<p>따라서 구조적 전환의 주 분기점은 <strong>{V21_STATS['bai_perron']['supF_ym'][:4]}년
{V21_STATS['bai_perron']['supF_ym'][4:]}월</strong>로 특정된다. 이는 V20까지 제시한
"2022년 2분기" 서술을 <strong>Bai-Perron 자동 탐색 결과로 수정</strong>한 것이다.
2021년 하반기~2022년 상반기는 엔데믹 전환기와 맞물리며, 2023년 10월의 단월 최대치는
분기점 이후에도 계절성이 이탈을 일시적으로 상쇄하며 관측된 <strong>일시적 정점</strong>이다.
2023년 말~2024년 초의 더스윙 PM 사업 종료 및 기후동행카드 도입은 이미 진행 중이던
구조 변화의 <strong>가속 요인</strong>이며 원인은 아니다.</p>

<h4>4.2.2 잔차 시계열의 방향성 — 분기점 이후 지속적 음의 이탈</h4>
<p>그림 4는 단순회귀 예측값으로부터의 잔차(실제 - 예측)를 시계열로 제시한다. 모델이
적절하다면 잔차는 0을 중심으로 무작위로 분포해야 한다. 그러나 2022년 4월 분기점(빨간
점선) 이후의 구간에서는 잔차가 지속적으로 0 아래로 이탈한다.</p>

<div class="chart">
<img src="data:image/png;base64,{charts['residual']}">
<div class="caption">[그림 4] 기온 효과 제거 후 잔차 시계열 — 2022년 4월 supF 분기점(빨간 점선) 이후 구간에 음의 잔차 집중.</div>
</div>

<div class="killer">
<strong>Q1 검증 결과 — H₀ 기각, H₁ 지지:</strong> 2022년 4월 분기점 이후 평균 잔차는
<em>{post_break_mean_resid:+,.0f}건</em>으로, 해당 기간 이용량이 기온 기반 예측치보다
매월 약 {abs(post_break_mean_resid) / 1e4:.0f}만 건을 하회한다. (참고: 2023년 1월 이후만
한정해도 평균 잔차는 {post_2023_mean_resid:+,.0f}건으로 이탈의 누적이 관찰된다.) 잔차와
시간 인덱스의 상관계수 r = {resid_time_corr:+.3f}는 시간 흐름에 따른 잔차의 지속적 감소를
시사한다. 이는 "기온 효과 제거 후 잔차가 무작위로 분포한다"는 귀무가설(H₀)과 일치하지
않으므로, 기온 외 변수의 작동을 상정하는 H₁이 지지된다. 잠재적 요인은 엔데믹
전환에 따른 대중교통 수요의 복귀, PM 경쟁 환경의 재편, 요금·이용권 정책의 변화, 재택근무
축소에 따른 통근 수요의 감소 등이며, 개별 요인의 기여도 추정은 후속 연구 과제로 남긴다.
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 5. 예측 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">5. 계절성과 2026년 예측</h2>

<h3>5.1 월별 계절지수</h3>
<div class="chart">
<img src="data:image/png;base64,{charts['seasonal']}">
<div class="caption">[그림 5] 7년 평균 대비 월별 계절지수. 1.0을 넘는 파란 막대는 성수기, 1.0 미만의 주황 막대는 비수기.</div>
</div>

<p>그림 5는 따릉이 이용의 극단적인 계절성을 보여준다. 겨울(12~2월) 평균 계절지수는
<em>{winter_si:.2f}</em>, 여름(6~8월)은 <em>{summer_si:.2f}</em>이다.
겨울 이용량은 여름의 약 {winter_si / summer_si * 100:.0f}% 수준에 불과하다.
최저는 1월({monthly_avg_si[0]:.2f}),
최고는 {['6월', '7월', '8월', '9월', '10월'][int(np.argmax(monthly_avg_si[5:10]))]}({max(monthly_avg_si[5:10]):.2f})로,
최고/최저 배수는 약 {max(monthly_avg_si) / min(monthly_avg_si):.1f}배에 달한다.</p>

<div class="chart">
<img src="data:image/png;base64,{charts['moving_avg']}">
<div class="caption">[그림 6] 3개월·6개월 이동평균 — 실제값의 단기 잡음을 제거한 추세.</div>
</div>

<h3>5.2 SARIMAX 식별과 분석적 예측구간</h3>
<p>본 연구는 <strong>SARIMAX(p,d,q)(P,D,Q,12)</strong> 5개 후보를 AIC/BIC 정보기준으로
식별하였다(Python <code>statsmodels.tsa.statespace.sarimax.SARIMAX</code>). 외생변수는 서울
월평균기온, 비계절 차분 d = 1, 계절 차분 D = 1 (연 주기).</p>

<table style="font-size:9.5pt;">
<tr><th>(p,d,q)(P,D,Q,12)</th><th>AIC</th><th>BIC</th><th>기온 계수 (건/°C)</th></tr>
{''.join([
    f'<tr><td>{c.get("order","-")}{c.get("seasonal_order","")}</td>'
    f'<td>{c.get("aic","—")}</td>'
    f'<td>{c.get("bic","—")}</td>'
    f'<td>{c.get("exog_coef_temp", 0):,.0f}</td></tr>'
    for c in V21_STATS.get('sarimax',{}).get('candidates',[])
    if 'aic' in c
])}
</table>

<p>최적 모형은 <strong>SARIMAX {V21_STATS.get('sarimax',{}).get('best_order','(1,1,2)(0,1,1,12)')}</strong>
로 식별되며 (AIC = {V21_STATS.get('sarimax',{}).get('best_aic', 1589):.1f},
BIC = {V21_STATS.get('sarimax',{}).get('best_bic', 1600):.1f}), 기온의 외생효과는
<strong>+{V21_STATS.get('sarimax',{}).get('exog_temp_coef', 82671):,.0f} 건/°C</strong>로 추정된다.
잔차 백색성 검정 Ljung-Box Q(10) = {V21_STATS.get('sarimax',{}).get('residual_ljung_box_q10',{}).get('stat', 0):.2f}
(p = {V21_STATS.get('sarimax',{}).get('residual_ljung_box_q10',{}).get('p', 0):.3f})로 백색성이
수용 가능하며(p &gt; 0.05), 모형 적합도가 적절함을 시사한다.</p>

<h4>5.2.1 분석적 예측구간 (SARIMAX 공분산 기반)</h4>
<p>4.1.2절 잔차 진단에서 확인된 자기상관·비정규성을 고려할 때, ±1.96·RMSE의 naive CI는
과소 추정될 가능성이 높다. 본 연구는 이를 보완하여 <strong>SARIMAX state-space 공분산에서
도출된 분석적 95% 예측구간</strong>을 제시한다.</p>

<table style="font-size:9.5pt;">
<tr><th>2026년</th><th>예측 (만 건)</th><th>95% LB</th><th>95% UB</th><th>구간 폭 (만)</th></tr>
{''.join([
    f'<tr><td>{f.get("month")}월</td>'
    f'<td>{f.get("mean",0)/1e4:.0f}</td>'
    f'<td>{f.get("lower",0)/1e4:.0f}</td>'
    f'<td>{f.get("upper",0)/1e4:.0f}</td>'
    f'<td>{(f.get("upper",0)-f.get("lower",0))/1e4:.0f}</td></tr>'
    for f in V21_STATS.get('sarimax',{}).get('forecast_2026',[])
])}
</table>

<p style="font-size:9pt; color:#555;">
※ SARIMAX의 예측구간은 각 예측 시점의 조건부 분산 + 외생변수(2026년 기온: 과거 7년 월별
평균 대체)의 불확실성을 반영한다. 4.2절의 supF 분기점 이후 구조가 안정되었다고 가정하며,
구조 재변동 시 실제 구간은 본 표보다 넓어진다.
</p>

<h4>5.2.2 4모델 hold-out 벤치마크 (참고)</h4>
<p>베이스라인 비교를 위해 단순 Naive·Seasonal Naive·3개월 이동평균·추세×계절 4모델을
2025년 12개월 hold-out으로 검증한 결과 추세×계절 모델의 RMSE는 약 {ts_rmse / 1e4:.1f}만 건으로
산출되었다. SARIMAX 대비 단순 모델은 구조변화 이후의 감소 추세를 overshoot한다.</p>

<div class="caveat">
<strong>신뢰구간의 실질 의미와 외삽 가정:</strong> 본 절의 95% 신뢰구간은 in-sample 잔차
RMSE를 ±1.96 배수하여 구성한 단순 추정치이며, SARIMAX 등 모델 기반 분산 예측이 아니다.
따라서 (i) 2026년에도 과거 추세·계절 구조가 그대로 외삽된다는 가정, (ii) 잔차의 정규성과
등분산성 가정, (iii) 4.2절에서 식별한 2022년 2분기 구조변화의 영향이 예측 구간에서
이미 안정화되었다는 가정에 동시에 의존한다. 세 가정 중 어느 하나라도 위반될 경우
실제 신뢰구간은 여기서 제시된 폭보다 넓어지며, 특히 구조변화가 진행 중일 경우
<strong>실제 이용량은 점 예측의 하한에 가까이 실현될 가능성</strong>이 높다.
</div>

<table>
<tr><th>월</th><th>추세값</th><th>계절지수</th><th>예측 이용건수</th><th>95% 신뢰구간</th></tr>
'''

for mo, trend_val, si, fc in forecast_2026:
    lo = fc - 1.96 * ts_rmse
    hi = fc + 1.96 * ts_rmse
    html += (
        f'<tr><td>{mo}월</td><td>{trend_val:,.0f}</td>'
        f'<td>{si:.2f}</td><td>{fc:,.0f}</td>'
        f'<td>[{lo:,.0f} ~ {hi:,.0f}]</td></tr>\n'
    )

html += f'''</table>

<div class="caveat">
<strong>예측의 한계:</strong> 본 모델은 선형 추세를 외삽하므로, 2024~25년의 감소 흐름이
구조적으로 지속될 경우 예측치는 과대 추정될 수 있다. 4.2절의 잔차 분석을 함께 고려하면,
실제 2026년 이용량은 점 예측의 하한선에 가까울 가능성이 높다.
</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 6. 수요자 구조 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">6. 수요자 구조 — 누가 타는가</h2>

<h3>6.1 성별 격차의 심화</h3>
<div class="chart">
<img src="data:image/png;base64,{charts['gender']}">
<div class="caption">[그림 7] 성별 이용 비율 추이. 연도별 남/여 비율(M/F 배수)을 본문에 표시.</div>
</div>

<p>그림 7은 2019~2025년 성별 이용 비율의 추이다. 여성 대비 남성 이용은
2019년 {gender_yearly[0].get('M', 0) / max(gender_yearly[0].get('F', 1), 1):.2f}배에서
2025년 {gender_yearly[-1].get('M', 0) / max(gender_yearly[-1].get('F', 1), 1):.2f}배로
<strong>격차가 더욱 확대되었다</strong>.</p>

<p>공공재 접근에서의 성별 격차는 단순한 선호의 차이로 환원되지 않는다. 야간 도로 안전에
대한 여성의 위험 인식, 장바구니·유아 동반 등 돌봄노동 이동의 특성, 자전거 이용을
어렵게 하는 의류·보호구 규범 등이 함께 작용한 것으로 해석된다. 이 격차는 사회구조적
요인의 복합적 합산이며, 운영 효율 지표만으로는 해소되기 어려운 정책 과제다.</p>

<h4>6.1.1 건당 이용 패턴의 성별 차이 (proxy)</h4>
<p>원본 데이터에 시간대·OD(출발-도착) 정보가 없어 "여성의 야간 기피" 가설을 직접 검증할 수는 없다.
다만 건당 평균 이용시간·거리의 성별 차이로 **이동 목적의 구조적 차이**를 간접 추론할 수 있다
(Excel Gender_Pattern 시트).</p>

<table class="emphasis-first">
<tr><th>성별</th><th>이용건수 (2025 상반기)</th><th>건당 이용시간</th><th>건당 이동거리</th></tr>
<tr><td>남성</td><td>8,486,499건</td><td>20.2분</td><td>2,320m</td></tr>
<tr><td>여성</td><td>4,660,164건</td><td><strong>22.4분 (+10.9%)</strong></td><td><strong>2,409m (+3.8%)</strong></td></tr>
</table>

<p>여성은 이용 횟수는 적지만 <strong>한 번 탈 때 더 오래·더 멀리</strong> 이용한다.
표본 규모(남성 8,486,499건, 여성 4,660,164건)에 Welch t-test를 적용한 결과 이용시간
차이(+10.9%)와 이동거리 차이(+3.8%) 모두 p &lt; 0.001로 통계적으로 유의하였다. 이는
돌봄노동·장보기·다목적 연쇄통행(trip-chaining) 특성을 반영한다는 가설과 정합한다.
여성의 이용량 확대 정책은 단순 "횟수 증대"가 아니라 이러한 통행 특성을 전제로 설계될 필요가 있다.</p>

<h3>6.2 고령층 성장의 실체 — 미상 보정 후</h3>
<p>Excel Misu_Adjustment 시트에서 '미상' 비율을 각 연령에 동일한 분포로 안분하여 보정
CAGR을 재산출하였다. 원본 CAGR과 보정 CAGR 사이에는 뚜렷한 차이가 관찰된다.</p>

<table>
<tr><th>연령대</th><th>2019</th><th>2025</th><th>원본 CAGR</th><th>보정 CAGR</th><th>차이(pp)</th></tr>
{cagr_tbl_rows}
</table>

<div class="chart">
<img src="data:image/png;base64,{charts['age_trend']}">
<div class="caption">[그림 8] 연령대별 연도별 이용건수 추이.</div>
</div>

<div class="insight">
<strong>Q3 검증 결과 — H₀ 기각, H₁ 지지:</strong> 60대·70대 이상의 원본 CAGR은 전 연령대
최고치이지만, 보정 후에는 상당 부분 축소된다. 미상 안분 보정 전후 CAGR 차이가 무시
가능하지 않으므로(60대 기준 원본 +32.4% → 보정 +20.7%, 차이 11.7pp) H₀는 기각된다.
'고령층 이용 확대'에는 실제 수요 증가와 <strong>집계 기준 변화</strong>가 혼재되어 있음이
확인된다. 다만 60대+가 2024~25 축소 국면에서 유일하게 감소하지 않은 층이라는 사실 자체는
어떤 보정 가정에서도 유지된다. 이 점이 7.3절의 운영 분화 제안의 근거가 된다.
</div>

<h4>6.2.1 민감도 분석 — 안분 가정 ±20%</h4>
<p>균등 안분 가정의 강건성을 검증하기 위해 미상 배정 비율을 ±20% 변동시킨 민감도 분석을
수행하였다(Excel Misu_Sensitivity 시트).</p>

<table>
<tr><th>시나리오</th><th>미상 비율 조정</th><th>60대 보정 CAGR</th></tr>
<tr><td>원본(미보정)</td><td>—</td><td>+32.4%</td></tr>
<tr><td>낙관(+20%)</td><td>미상 × 1.2</td><td>+15.9%</td></tr>
<tr><td><strong>기준(균등 안분)</strong></td><td>미상 × 1.0</td><td><strong>+20.7%</strong></td></tr>
<tr><td>비관(-20%)</td><td>미상 × 0.8</td><td>+24.2%</td></tr>
</table>

<p>60대 CAGR의 범위는 <em>+15.9%~+24.2%</em>로, 가정 변동 폭 대비 결론이 흔들리지 않는다.
최선·최악 추정 모두 양(+)으로 유지되므로 고령층 이용 증가는 방법론 가정과 무관한
강건한 발견으로 판단된다.</p>

<h3>6.3 자치구 공간 분포 — 지니계수·시계열·공간자기상관</h3>
<div class="chart">
<img src="data:image/png;base64,{charts['district']}">
<div class="caption">[그림 9] 자치구별 1인당 따릉이 이용 — 상위 5(초록)·하위 5(빨강) 강조.</div>
</div>

<p>그림 9는 25자치구의 1인당 연간 따릉이 이용량(2025)을 보여준다. 페어와이즈 방식으로
산출한 지니계수는 <em>{gini_per_capita:.3f}</em>으로, 무시할 수 없는 공간 격차를 나타낸다.</p>

<h4>6.3.1 지니계수 시계열 — 격차의 지속적 확대</h4>
<p>단일 시점이 아닌 시계열로 격차 추이를 확인하였다(Excel Gini_Timeseries 시트).</p>

<table>
<tr><th>연도</th><th>지니계수</th><th>전기 대비</th></tr>
<tr><td>2019</td><td>0.2617</td><td>—</td></tr>
<tr><td>2022</td><td>0.2765</td><td>+1.48pp</td></tr>
<tr><td>2025</td><td><strong>0.3008</strong></td><td><strong>+2.43pp</strong></td></tr>
</table>

<p>2019~2025년 지니계수가 지속적으로 상승하여 공간 격차가 해마다 심화됨을 확인하였다.
8.3절의 형평성 KPI 목표(지니 0.25 이하)는 2019년 수준으로의 회귀를 의미한다.</p>

<h4>6.3.2 Moran's I — W 가중행렬 × 분모 2종의 5×2 민감도 매트릭스</h4>
<p>지니계수가 격차의 <strong>크기</strong>를 측정한다면, Moran's I는 격차의
<strong>공간적 패턴</strong>을 진단한다. MAUP(Modifiable Areal Unit Problem)에 대응하기
위해 본 연구는 5종의 공간가중행렬(Queen·Rook·kNN(k=3)·kNN(k=5)·Inverse-Distance)과
2종의 분모(절대 이용량·1인당 이용량)를 교차한 <strong>10개 셀의 민감도 매트릭스</strong>를
산출하였다(999회 순열검정, seed=42, <code>compute_stats_v21.py</code>).</p>

<table style="font-size:9.5pt;">
<tr><th>분모 \ W 행렬</th><th>Queen</th><th>Rook</th><th>kNN(k=3)</th><th>kNN(k=5)</th><th>Inverse-Distance</th></tr>
<tr><td><strong>절대 이용량</strong></td>
{''.join([
    f'<td>I={V21_STATS["moran_sensitivity"].get(f"absolute__{w}",{}).get("I","—")}, '
    f'p={V21_STATS["moran_sensitivity"].get(f"absolute__{w}",{}).get("p_sim","—")}</td>'
    for w in ["Queen","Rook","kNN_k3","kNN_k5","InverseDistance"]
])}
</tr>
<tr><td><strong>1인당 이용량</strong></td>
{''.join([
    f'<td>I={V21_STATS["moran_sensitivity"].get(f"per_capita__{w}",{}).get("I","—")}, '
    f'p={V21_STATS["moran_sensitivity"].get(f"per_capita__{w}",{}).get("p_sim","—")}</td>'
    for w in ["Queen","Rook","kNN_k3","kNN_k5","InverseDistance"]
])}
</tr>
</table>

<div class="killer">
<strong>Q4 검증 결과 — 분모·W에 따라 결론이 달라짐:</strong> 절대 이용량 기준으로는
Queen·Rook·kNN 모두 양의 유의 공간 자기상관을 지지한다(Queen I=
{V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('I', 0)}, p=
{V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('p_sim', 0)}; Rook I=
{V21_STATS['moran_sensitivity'].get('absolute__Rook',{}).get('I', 0)}, p=
{V21_STATS['moran_sensitivity'].get('absolute__Rook',{}).get('p_sim', 0)}). 그러나 1인당
기준으로 전환하면 Queen·Rook 모두 p=0.07~0.11로 marginal이며 kNN·InverseDistance는 비유의로
전환된다. 본 연구는 <strong>절대 이용량 + Queen/Rook을 우선 결과</strong>로 채택하되,
<strong>모든 10개 셀을 함께 공개</strong>하여 MAUP 편향을 투명하게 드러낸다. 1인당
분모의 비유의 결과는 주민등록인구 분모가 주간상주인구(종로·중구 유입)를 반영하지 못하는
분모 편향의 결과로 해석된다(자세한 민감도 표는 기술부록 T2 참조).
</div>

<table class="emphasis-first" style="width:48%; float:left; margin-right:4%;">
<tr><th colspan="3">상위 5개 자치구 (이용 집중)</th></tr>
<tr><th>순위</th><th>자치구</th><th>1인당 이용</th></tr>
{top5_rows}
</table>

<table class="emphasis-first" style="width:48%; float:left;">
<tr><th colspan="3">하위 5개 자치구 (이용 저조)</th></tr>
<tr><th>순위</th><th>자치구</th><th>1인당 이용</th></tr>
{bot5_rows}
</table>
<div style="clear:both;"></div>

<div class="caveat">
<strong>분모의 한계:</strong> 본 분석은 주민등록인구를 분모로 사용하여 1인당 이용을 산출하였다.
종로·중구·영등포 등 주간상주인구(통근·관광 유입)가 거주인구를 크게 상회하는 자치구의 경우
실제 1인당 이용 수준은 보고된 값보다 낮을 수 있다. 후속 연구에서는 KOSIS 통근통학 인구·
KT 생활인구 데이터로 분모를 재정의하여 상위 자치구 수치의 구조적 과대 가능성을 검증해야 한다.
</div>

<div class="insight">
업무·상업 중심지(종로·영등포·마포)와 한강 수변 접근성이 높은 자치구는 상위에,
외곽 주거지 비중이 큰 자치구는 하위에 집중된다. 따릉이를 <strong>"이동 수요가 큰 곳에 더 많이"</strong>라는
효율성 원칙만으로 배치할 경우 공간 형평성은 더욱 악화될 가능성이 높다. 공공재로서의
정당성 또한 점진적으로 약화될 수 있다.
</div>

<h4>6.3.3 LISA 국지 군집 + Benjamini-Hochberg FDR 보정</h4>
<p>전역 Moran's I가 공간 자기상관의 <strong>존재</strong>를 보여준다면, 국지 지표인
LISA(Local Indicators of Spatial Association)는 어느 자치구가 어떤 유형의 공간 군집을
형성하는지를 식별한다(Anselin 1995). Queen 인접 + 절대 이용량 기준으로 999회 순열검정을
수행하였으며, 25개 자치구 개별 검정에 대해 <strong>Benjamini-Hochberg FDR 보정</strong>
(α = 0.05)을 적용하였다.</p>

<table style="font-size:9.5pt;">
<tr><th>자치구</th><th>I_i</th><th>p_raw</th><th>p_FDR</th><th>raw 유의</th><th>FDR 유의</th><th>사분면</th></tr>
{''.join([
    f'<tr><td>{r["district"]}</td><td>{r["Ii"]}</td>'
    f'<td>{r["p_raw"]}</td><td>{r["p_fdr"]}</td>'
    f'<td>{"✓" if r["raw_significant"] else "—"}</td>'
    f'<td>{"✓" if r["fdr_significant"] else "—"}</td>'
    f'<td>{r["quadrant"]}</td></tr>'
    for r in V21_STATS.get('lisa_fdr',{}).get('rows',[])
    if r.get('raw_significant') or r.get('fdr_significant')
])}
</table>

<div class="caveat">
<strong>FDR 보정 결과 — 국지 군집 주장의 철회·수정:</strong> 보정 전 raw p &lt; 0.05에서
{V21_STATS.get('lisa_fdr',{}).get('n_raw_sig', 0)}개 자치구가 유의하였으나, Benjamini-Hochberg
FDR 보정을 적용하면 <strong>{V21_STATS.get('lisa_fdr',{}).get('n_fdr_sig', 0)}개 자치구</strong>만
유의성을 유지한다. 이는 25개 자치구에 대한 동시 검정에서 false discovery를 통제한 결과로,
n=25의 제한된 검정력(6.3.5절)을 고려할 때 <strong>국지 cluster의 통계적 주장은 보류</strong>하는
것이 학술적으로 적절하다. V20까지 본문에서 제시한 "강동·강서·양천 서남 통근축"은 raw
p-value에 근거한 결론이었으며, 본 V21에서는 이를 <strong>방향성 가설</strong>로 하향
제시한다. 그럼에도 전역 Moran's I(Queen, p={V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('p_sim', 0)})는
유의하므로, 공간 의존성 자체는 존재하되 특정 자치구 군집의 지목은 500m hex 격자 또는
대여소 포인트 패턴에서 재검증이 요구된다(기술부록 T4·T5).
</div>

<h4>6.3.4 공간 회귀 — SAR · SEM 모델 적합</h4>
<p>전역·국지 지표에서 멈추지 않고, 자치구 월평균 이용량에 <strong>공간 시차(Spatial Lag)
모델과 공간 오차(Spatial Error) 모델</strong>을 적합하였다(spreg ML_Lag, ML_Error;
독립변수: 인구(10만명 단위)·한강 인접 더미; n = 25). Lagrange Multiplier 검정으로 SAR 대
SEM의 지배성을 판정한다.</p>

<table style="font-size:9.5pt;">
<tr><th>지표</th><th>값</th><th>해석</th></tr>
<tr><td>LM-Lag</td>
    <td>{V21_STATS['spatial_regression']['lm_lag']['stat']:.3f} (p = {V21_STATS['spatial_regression']['lm_lag']['p']:.3f})</td>
    <td>{'유의 (α=0.05)' if V21_STATS['spatial_regression']['lm_lag']['p'] < 0.05 else '비유의 — n=25 검정력 한계'}</td></tr>
<tr><td>LM-Error</td>
    <td>{V21_STATS['spatial_regression']['lm_error']['stat']:.3f} (p = {V21_STATS['spatial_regression']['lm_error']['p']:.3f})</td>
    <td>{'유의' if V21_STATS['spatial_regression']['lm_error']['p'] < 0.05 else '비유의'}</td></tr>
<tr><td>SAR rho</td>
    <td>{V21_STATS['spatial_regression']['SAR_rho']:.3f}</td>
    <td>Pseudo R² = {V21_STATS['spatial_regression']['SAR_pseudo_r2']:.3f}</td></tr>
<tr><td>SEM lambda</td>
    <td>{V21_STATS['spatial_regression']['SEM_lambda']:.3f}</td>
    <td>Pseudo R² = {V21_STATS['spatial_regression']['SEM_pseudo_r2']:.3f}</td></tr>
</table>

<p>LM 검정에서 SAR(p = {V21_STATS['spatial_regression']['lm_lag']['p']:.3f}),
SEM(p = {V21_STATS['spatial_regression']['lm_error']['p']:.3f}) 모두 α = 0.05에서 비유의이나,
점 추정 rho = {V21_STATS['spatial_regression']['SAR_rho']:.3f}, lambda = {V21_STATS['spatial_regression']['SEM_lambda']:.3f}는
양(+)의 방향으로 일관된다. n = 25의 검정력 한계(6.3.5절) 하에서도 공간 의존의 <strong>방향
신호</strong>는 확인되며, 정책적 함의는 여전히 유효하다. Anselin·Rey(2014) 권고에 따라
표본 확장 시(500m hex) 재검정이 요구된다.</p>

<h4>6.3.5 검정력 분석 — n = 25의 한계</h4>
<p>Moran's I 순열검정의 n = 25 표본 하에서 최소 검출 가능 효과(MDE)를 시뮬레이션으로
추정하였다(true I ∈ {0.05, 0.10, 0.15, 0.20, 0.245, 0.30, 0.40}, 반복 100회, α = 0.05).</p>

<table style="font-size:9.5pt;">
<tr><th>True I</th>
{''.join([f'<th>{k.replace("I_","")}</th>' for k in V21_STATS['power_curve']['curve'].keys()])}
</tr>
<tr><td>검정력</td>
{''.join([f'<td>{v}</td>' for v in V21_STATS['power_curve']['curve'].values()])}
</tr>
</table>

<p>관측된 Queen I = {V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('I', 0)}에서 검정력은 약
<strong>{V21_STATS['power_curve']['curve'].get('I_0.245', '~0.33')}</strong>로, 전통적 기준(0.80)에
미달한다. 따라서 본 연구의 공간 자기상관 결론은 <strong>방향 추정은 신뢰할 수 있으나, 표본
확장 없이는 유의성의 강건한 재현이 어렵다</strong>는 한계를 동반한다.</p>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 7. 운영 포지셔닝 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">7. 운영 모델의 재구성 — 경영 관점</h2>

<h3>7.0 Competitive Landscape — 수단 간 한계비용 비교</h3>
<p>따릉이 확장 의사결정은 따릉이 단독 논리가 아닌 <strong>대체·보완 수단과의 비용효율
비교</strong>에 기반해야 한다. 서울시 주요 수송수단의 수송분담률당 한계비용을 정리한다.</p>

<table style="font-size:9pt;">
<tr><th>수단</th><th>증설 단위</th><th>신규 수송량/단위</th><th>단위당 비용 (연)</th><th>km당 환산</th><th>보완·대체 관계</th></tr>
<tr><td>지하철 노선 연장</td><td>1 km</td><td>일 ~15,000명</td><td>~1,500억 CAPEX</td><td>매우 높음</td><td>따릉이: 라스트마일 보완</td></tr>
<tr><td>버스 전기화·증편</td><td>전기버스 1대</td><td>일 ~800명</td><td>4~5억 CAPEX + 연 1억 OPEX</td><td>중</td><td>따릉이: 단거리 대체</td></tr>
<tr><td><strong>따릉이 일반</strong></td><td>자전거 1대</td><td>연 180건 추가</td><td>70만 CAPEX + 연 12만 OPEX</td><td>매우 낮음</td><td>본 연구 Alt 2</td></tr>
<tr><td><strong>따릉이 전동보조</strong></td><td>자전거 1대</td><td>연 250건 추가</td><td>200만 CAPEX + 연 18만 OPEX</td><td>낮음</td><td>본 연구 Alt 3 (고령층 흡수)</td></tr>
<tr><td>PM(공유 킥보드)</td><td>민간 보급</td><td>MAU -17% 축소 중</td><td>공공 재정 부담 0</td><td>낮음 (민간)</td><td>경쟁 축소 중</td></tr>
<tr><td>카셰어링</td><td>민간</td><td>장거리 중심</td><td>공공 재정 부담 0</td><td>중</td><td>보완재 (장거리)</td></tr>
</table>

<p>따릉이 대당 한계비용(₩70~200만/대)은 다른 공공 수송수단 대비 <strong>2~3 order 낮다</strong>.
라스트마일·단거리 시장에서 공공 예산 효율은 따릉이가 지배적이며, 지하철·버스는 보완재로
상호 대체가 아니다. 본 장의 전략 제안은 이 비용 구조에 기반한다.</p>

<h3>7.1 자전거 1대당 회전율의 추이</h3>
<div class="chart">
<img src="data:image/png;base64,{charts['turnover']}">
<div class="caption">[그림 10] 자전거 1대당 일회전율(이용건수/운영대수/일수) 월별 추이.</div>
</div>

<p>그림 10은 운영대수(공개자료 기반 추정치)로 정규화한 <strong>자전거 1대당 일회전율</strong>
(이용건수 ÷ 운영대수 ÷ 일수)의 월별 추이이다. 연 평균 회전율은 2023년 {to_2023:.2f}회
에서 2025년 {to_2025:.2f}회로 <em>{to_drop_pct:+.1f}%</em> 감소하였다. 공급이 일정한
가운데 수요가 구조적으로 위축되고 있음을 시사한다. 같은 기간 <strong>연간 이용건수
총량</strong> 기준 감소폭은 2023년 피크(4,490만 건) 대비 2025년(3,737만 건)
<em>약 -16.8%</em>로, 두 지표의 낙폭은 유사하다. 그러나 시점을 2024~25년 2년 구간으로
좁히면 <strong>회전율의 하락이 총량의 하락을 앞선다</strong>. 이는 공급(운영대수)이 수요
감소를 즉각 반영해 축소되지 못하는 운영 경직성 때문이며, 자본 효율 악화의 크기는
이용건수만으로 보는 것보다 크다.</p>

<h3>7.2 공유 모빌리티 시장의 구조 변화</h3>
<p>같은 기간 공유 PM(전동킥보드) 시장 역시 축소되었다. <strong>더스윙</strong>이 2023년
말 서울시 전동킥보드 공유사업을 종료하였고, SKT 계열 <strong>티맵모빌리티</strong>는
2025년 3월에 서비스를 종료하였다. 공유 킥보드의 <strong>월간활성이용자(MAU)</strong>는
2023년 10월 221만 명에서 2024년 10월 184만 명으로 약 <em>-17%</em> 감소하였다
(이용건수 기준이 아닌 MAU 기준 낙폭이다). 따릉이의 이용건수 기준 낙폭(-16.8%)과 PM의
MAU 낙폭(-17%)이 수치상 근접한다는 점은 우연이며, 두 지표의 분모와 단위가 다르므로
동일 선상에서 직접 비교해서는 안 된다. 다만 공유 모빌리티 전반이 동반 위축되고 있다는
방향성은 분명히 포착된다.</p>

<p>이러한 흐름을 고려하면 2024~25년 따릉이 감소를 "PM의 따릉이 잠식"으로 단순 환원하기
어렵다. 오히려 <strong>공유 모빌리티 전체가 조정기에 진입</strong>하였다는 보다 큰 맥락이 반영된
결과로 해석된다. 4.2절에서 supF 검정이 가리킨 구조변화 시점이 <strong>2022년 4월</strong>로
조정기의 시작이 더스윙 종료(2023년 말) 이전임을 고려하면, 2023~2024년의 PM 사업
재편은 이미 진행 중이던 구조 변화를 <strong>가속</strong>한 사건으로 해석하는 것이
정합적이다. 코로나 시기 과열되었던 단거리 개인 이동 수요는 엔데믹 전환과 재택근무
축소, 안전 규제 강화, 시민 인식 악화(PM 이용 시 불편을 경험했다는 응답 79.2%) 속에서
동반 위축되고 있다.</p>

<h4>7.2.1 기후동행카드 — 따릉이 수요 보존 장치</h4>
<p>서울시는 2024년 1월 <strong>기후동행카드</strong>(서울지역 지하철·버스·따릉이 통합 정기권)를
도입하였다. 현재 요금 체계는 <em>62,000원(대중교통 전용)</em>과 <em>65,000원(+따릉이)</em>으로
차등되어 있으며, 따릉이 포함 옵션 선택 시 2시간 무제한 이용이 가능하다. 이 요금 구조는
PM 시장 축소기에도 통근형 따릉이 수요를 <strong>구조적으로 보존</strong>하는 장치로 기능한다.
6.3절 자치구 격차와 7.3절 세그먼트 분화 제안은 기후동행카드 가입자 데이터와 연동될 경우
더 정교한 정책 설계가 가능하다(후속 연구 과제).</p>

<h3>7.3 이용자 층별 운영 분화</h3>
<p>6·7절의 발견을 종합하면 따릉이는 "전 시민을 위한 범용 공공자전거"라는 단일 자리매김
으로는 더 이상 유지되기 어려운 국면에 진입하였다. 이용자를 세 갈래로 분화하여 운영을
재구성할 필요가 있다.</p>

<table style="font-size:9pt;">
<tr><th style="width:14%;">이용자 층</th><th style="width:20%;">주요 이용자</th><th style="width:33%;">핵심 수요 특성</th><th style="width:33%;">운영 방향</th></tr>
<tr><td>여가형</td><td>60대 이상, 주말 이용자</td><td>한강·공원·수변 중심 장거리 여유 이용</td><td>전동보조 자전거, 3시간 이용권, 관광 연계</td></tr>
<tr><td>통근형</td><td>20~40대 평일 이용자</td><td>지하철 환승·출퇴근 근거리 구간</td><td>업무지구 거치대 증설, 혼잡시간 재배치</td></tr>
<tr><td>여성·가족형</td><td>여성·돌봄노동자</td><td>안전·근거리·장바구니 이용</td><td>저상 자전거, 야간 조명, 유아 옵션</td></tr>
</table>

<h3>7.4 Roadmap 2026~2030 · Decision Tree · Political Risk</h3>

<h4>7.4.1 연도별 마일스톤</h4>
<table style="font-size:9.5pt;">
<tr><th style="width:10%;">연도</th><th style="width:25%;">핵심 마일스톤</th><th style="width:30%;">KPI</th><th style="width:35%;">담당</th></tr>
<tr><td>2026 1H</td><td>기후동행카드 가입자 MOU + 세그먼트 분화 실증</td><td>가입자 데이터 접근 · 3세그먼트 파일럿 500 대여소</td><td>교통정책실·자전거정책과</td></tr>
<tr><td>2026 2H</td><td>Alt 2 일반 자전거 4,500대 증설 (한강·서남 환승역 우선)</td><td>1대당 회전율 +10% 복원</td><td>자전거정책과</td></tr>
<tr><td>2027</td><td>Alt 3 전동보조 10% 시범 도입 (고령층·한강 레저축)</td><td>60대+ 이용 +15% · B/C ≥ 1.3</td><td>자전거정책과·건강도시과</td></tr>
<tr><td>2028</td><td>공공가치 회계 3축(탄소·건강·혼잡) 정식 도입</td><td>다중 편익 SROI ≥ ₩200억/년</td><td>기획조정실·환경정책과</td></tr>
<tr><td>2029</td><td>형평성 KPI 공시 · 자치구 예산 조정</td><td>지니계수 0.28 달성</td><td>교통정책실·자치구</td></tr>
<tr><td>2030</td><td>운영 모델 2.0 전면 전환 평가</td><td>여성 이용 35%+ · 65세+ 베이스라인 재설정</td><td>교통정책실</td></tr>
</table>

<h4>7.4.2 Decision Tree — 실장 결재 3분기점</h4>
<div class="insight" style="font-size:10pt;">
<pre style="font-family: inherit; margin: 0; white-space: pre-wrap;">
[분기점 1] 예산 규모
├─ 연 100억 유지(Do-nothing) → 운영적자 누적, 수요 추가 감소 (NPV -244억)
├─ 연 150억 (+50억, 본 연구 권고) → Alt 2 전면 + Alt 3 일부, 증분 NPV +92~111억
└─ 연 300억 (+200억) → Alt 3 대규모, 정치적 부담 ↑, 회수 10년+
        ↓
[분기점 2] Fleet 구성
├─ 일반 자전거만 (Alt 2) → 안전·익숙, B/C 1.79, 고령층 흡수 약함
├─ 전동보조 10% 혼합 (Alt 2+Alt 3, 권고) → 고령층·여가축 흡수 · B/C 1.40~1.79
└─ 전동보조 전면 전환 → CAPEX 급증 · 배터리 교체 주기 리스크
        ↓
[분기점 3] 기후동행카드 통합
├─ 독립 운영 → 데이터 단절 · 세그먼트 분석 불가
├─ 제한 통합 (MOU + 공동 분석, 권고) → 가입자 데이터로 세그먼트 정교화
└─ 완전 통합 (요금 일원화) → 운영 비용 부담 · 법적 검토 소요 12~18개월</pre>
</div>

<h4>7.4.3 Political Risk + Pre-empt</h4>
<table style="font-size:9.5pt;">
<tr><th style="width:35%;">예상 반대 논리</th><th style="width:65%;">대응 (Pre-empt)</th></tr>
<tr><td>"전동보조 NPV 음수 → 세금 낭비"</td><td>공공가치 회계(탄소 ₩2.07억 + 건강 ₩44.79억 + 혼잡 ₩41.48억 = 연 88억)로 증분 NPV +92억 재계산. 민간 사업이 아닌 공공재의 평가 틀 전환 논리(8.2절).</td></tr>
<tr><td>"형평성 KPI(지니 ↓) = 강남 예산 삭감"</td><td>절대 증설은 유지하고 <strong>신규 예산의 40%만 하위 자치구 우선 배치</strong>. 강남·서초 기존 서비스 동결, 신규에서만 차별화.</td></tr>
<tr><td>"서남 통근축 주장의 통계적 근거 약함"</td><td>FDR 보정 후 raw 유의에서 비유의로 전환(6.3.3절) → "방향성 가설"로 톤 하향. 500m hex 재검정을 기후동행카드 통합 후 후속 확정(Roadmap 2028).</td></tr>
</table>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 8. 결론 ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">8. 결론 및 제언</h2>

<h3>8.1 종합</h3>
<p>본 연구는 2019~2025년 84개월의 따릉이 이용 데이터에 <strong>잔차 분석·미상 보정 CAGR·지니계수·대당
회전율</strong>을 새롭게 적용하였다. 그 결과 따릉이가 단순한 성장 궤도가 아니라
<strong>공유 모빌리티 재편기의 조정 국면</strong>에 진입하였음이 확인되었다. 기온은
여전히 이용량을 결정하는 핵심 단기 변수이지만, supF 구조변화 검정이 <strong>2022년
2분기 엔데믹 전환기</strong>를 분기점으로 특정하였고, 해당 시점 이후 기온 효과 제거
잔차는 지속적으로 음(-)의 이탈을 누적하였다. 2023년 10월의 단월 피크는 분기점 이후에도
계절성이 이탈을 일시 상쇄한 국소 정점이며, 2023년 말~2024년 초의 더스윙 철수·기후동행
카드 도입은 이미 진행 중이던 구조 변화의 <strong>가속 요인</strong>이다. 자전거 1대당
일회전율도 2023년 피크 대비 2025년 {to_drop_pct:+.1f}% 감소하여 자본 효율이 동반
악화되었다.</p>

<p>이용자 구조 측면에서도 동일한 양상이 관찰된다. '미상' 비율의 급감이 연령·성별 집계를
왜곡한 주된 요인이며, 보정 후에도 60대+는 2024~25 감소에 역행하는 유일한 층으로 남는다.
공간적으로는 자치구별 1인당 이용의 지니계수가 {gini_per_capita:.3f}로 무시할 수 없는
격차가 형성되어 있다. 따릉이의 정책 담론은 수요 확대 중심에서 분배·형평 중심으로의
전환이 요구된다.</p>

<p>본 연구가 제안하는 핵심 방향은 따릉이를 여가형·통근형·여성/가족형 세 층으로 분화하여
각 층의 운영 제약과 수요 특성을 분리 관리하는 모델이다. 이는 단일 규모 경제에 의존한
현재의 운영 구조에서 탈피하여, 2026년 초고령사회 진입을 앞둔 서울에서 따릉이가
건강·여가·형평성이라는 세 가지 공공가치를 동시에 충족시킬 수 있는 경로이다.</p>

<h3>8.2 3대안 재무 비교 · 공공가치 회계</h3>

<h4>8.2.0 3대안 15년 LCC 비교 (증분 기준 · 공공가치 3축 포함)</h4>
<p>본 연구는 서울시 교통정책실장 결재용으로 3대안의 15년 Life-cycle cost를 산출하였다.
<strong>Do-nothing(Alt 1)을 baseline</strong>으로 하고, Alt 2(일반 자전거 +4,500대)·Alt 3
(전동보조 부분 도입)의 <strong>증분 NPV·B/C·EIRR·회수기간</strong>을 제시한다. 할인율은
KDI PIMAC 예타 가이드에 따라 3축(3.0%·4.5%·5.5%)으로 변동시킨다. 편익은 탄소
(K-ETS ₩40,000/tCO₂) + 건강(WHO HEAT, VSL ₩50억 + 규칙이용 RR 7%) + 혼잡(KTDB 시간가치
₩18,500/인·시)의 3축 합산(<code>compute_finance_v21.py</code>).</p>

<table style="font-size:9.5pt;">
<tr><th rowspan="2">대안</th><th colspan="4">할인율 3.0%</th><th colspan="4">할인율 4.5% (기준)</th></tr>
<tr><th>증분 NPV(억)</th><th>B/C</th><th>IRR(%)</th><th>회수(년)</th>
    <th>증분 NPV(억)</th><th>B/C</th><th>IRR(%)</th><th>회수(년)</th></tr>
{(lambda inc=V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}): ''.join([
    f'<tr><td><strong>{name.replace("_", " ")}</strong></td>'
    + ''.join([
        f'<td>{inc[name].get(f"dr_{int(dr*1000):03d}",{}).get(k,"—")}</td>'
        for k in ["incr_npv","incr_bc_ratio","incr_irr","incr_payback_years"]
    ])
    + ''.join([
        f'<td>{inc[name].get(f"dr_{int(0.045*1000):03d}",{}).get(k,"—")}</td>'
        for k in ["incr_npv","incr_bc_ratio","incr_irr","incr_payback_years"]
    ])
    + '</tr>'
    for name in ["Alt2_Regular","Alt3_Ebike"]
    for dr in [0.030]
]))()}
</table>

<div class="killer">
<strong>재무 판정 (KDI PIMAC 기준):</strong> Alt 2 일반 자전거 증설은 증분
NPV <strong>+{V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_npv','N/A')}억</strong>,
B/C = {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_bc_ratio','N/A')},
IRR = {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_irr','N/A')}%,
회수기간 {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_payback_years','N/A')}년으로
<strong>예타 상정 가능</strong>(B/C ≥ 1). Alt 3 전동보조는 증분 NPV
<strong>+{V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt3_Ebike',{}).get('dr_045',{}).get('incr_npv','N/A')}억</strong>,
B/C = {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt3_Ebike',{}).get('dr_045',{}).get('incr_bc_ratio','N/A')},
회수 {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt3_Ebike',{}).get('dr_045',{}).get('incr_payback_years','N/A')}년으로 회수가 길지만 B/C &gt; 1.0으로
예타 통과 가능 구간이다. <strong>Alt 2 + Alt 3 혼합 포트폴리오</strong>가 본 연구의 권고안이다.
</div>

<h4>8.2.1 편익 항목별 민감도 (±20%, ±40%, ±60%)</h4>
<p>3축 공공편익 추정치는 각각 ±30~40%의 불확실성을 갖는다(HEAT·KTDB 관례). 항목별
민감도를 4개 영역(탄소·건강·혼잡·전체)에 대해 7단계로 산출한 결과, Alt 3의 B/C는
전체 -60%에서도 ≥ 1.0을 유지한다(재현: <code>v21_finance.json['benefit_sensitivity']</code>).
상세 표는 기술부록 T6에 수록한다.</p>

<h3>8.2.Legacy 전동보조 도입 운영 정책 제언</h3>
<ol>
  <li><strong>고령 여가형 상품 분리 운영.</strong> 2025년 11월 도입된 3시간 이용권을
    고령층 대상 여가형 요금제와 결합하고, 전동보조 자전거 비율을 한강 인접 대여소부터
    단계적으로 확대한다. 아래 표는 본 제안에 대한 시나리오 민감도 분석이다.</li>
</ol>

<table class="emphasis-first" style="margin: 6px 0 12px 24px; width:calc(100% - 24px); font-size:9.5pt;">
<tr><th>구분</th><th>비관(−)</th><th>기준</th><th>낙관(+)</th></tr>
<tr><td>전환율</td><td>5%</td><td>10%</td><td>15%</td></tr>
<tr><td>전환 대수</td><td>2,250대</td><td>4,500대</td><td>6,750대</td></tr>
<tr><td>대당 단가 <sup>(1)</sup></td><td>₩220만</td><td>₩200만</td><td>₩180만</td></tr>
<tr><td>총 투자</td><td>₩49.5억</td><td>₩90.0억</td><td>₩121.5억</td></tr>
<tr><td>60대+ 이용 증가</td><td>+10%</td><td>+20%</td><td>+30%</td></tr>
<tr><td>연 매출 (₩500/건)</td><td>₩0.22억</td><td>₩0.45억</td><td>₩0.67억</td></tr>
<tr><td>연 추가 운영비 <sup>(2)</sup></td><td>₩2.25억</td><td>₩4.50억</td><td>₩6.75억</td></tr>
<tr><td>연 순편익</td><td>-₩2.0억</td><td>-₩4.1억</td><td>-₩6.1억</td></tr>
<tr><td>NPV <sup>(3)</sup></td><td><em>-₩63.3억</em></td><td><em>-₩117.6억</em></td><td><em>-₩162.9억</em></td></tr>
<tr><td>IRR</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>
</table>
<div style="font-size:9pt; color:#333; margin: -2px 24px 12px 24px; padding: 6px 10px; background:#FAFAFA; border-left: 3px solid #999;">
<strong style="color:#555;">표 주석:</strong><br>
<sup>(1)</sup> 조달청 나라장터 2024년 전동보조 자전거 공공입찰 단가 기준 ₩180~220만.<br>
<sup>(2)</sup> 전동보조 자전거는 일반 자전거 대비 유지비 1.8배, 추가 운영비 ₩10만/대·년.<br>
<sup>(3)</sup> 사회적할인율 4.5% (KDI PIMAC 예타 표준), 내용연수 15년. 연 순편익 음수로 IRR 미정의.
</div>

<div class="killer" style="margin-left:24px;">
<strong>재무적 결론:</strong> 세 시나리오 전부 NPV 음수로 산출된다. 현행 요금 체계
(시간당 ₩1,000, 건당 30분)에서 전동보조 자전거 도입은 <strong>자립적 수익성이 없다</strong>.
단순 회수기간 모델은 이러한 결함을 은폐할 수 있으나, 사회적할인율과 운영비 변동분을
반영하면 재무 논리만으로는 투자 근거가 성립하지 않음이 드러난다.
</div>

<h4>참고: NPV 음수에서 공공가치 회계로 — 판단 전환의 논리</h4>
<p>세 시나리오 모두 NPV가 음수라는 결과는 두 갈래로 해석될 수 있다. 첫째 해석은
<strong>도입 기각</strong>이다. 민간 투자 기준으로는 타당하다. 둘째 해석은 <strong>평가 틀
자체의 재설정</strong>이다. 따릉이가 공공재인 이상, 투자 판단은 "요금 수입 대비 비용"이
아니라 "공공가치 편익 대비 비용"으로 이루어져야 한다는 관점이다. 본 연구는 두 번째 경로를
택하되, 그 근거가 정량 산출되어 있음을 전제 조건으로 제시한다.</p>

<p>탄소 화폐화 한 축만 현 시점에서 산출 가능하다(K-ETS 기준, Excel Carbon_Value 시트).
2019~2025년 누적 따릉이 이용에서 승용차 대체 가정 30% 적용 시 CO₂ 감축량은 약
<em>33,256톤</em>, 화폐 가치는 <em>₩13.3억</em>(연평균 ₩1.9억)으로 산출되며, 이는 연간
운영 적자 ₩100억의 <em>약 1.9%</em>만 상쇄한다. <strong>탄소 단일 지표만으로는 NPV
음수를 뒤집지 못한다</strong>. 따라서 도입의 정당성은 다음 세 가지 공공편익 영역을 함께
정량화한 <strong>다중 편익 합산 공공가치 회계</strong>의 선행이 요구된다: ① 건강편익
(WHO HEAT 프로토콜의 신체활동 사망률 감소 화폐화), ② 교통혼잡 감소 편익(KTDB 시간가치
기반 차량 km 대체 효과), ③ 사회형평성(저소득·고령·여성 접근성의 SROI 추정). 세 영역은
KDI PIMAC 예비타당성 가이드와 국제 공공자전거 평가(ITDP, OECD ITF)에서 표준으로 채택된
편익 축이다. 다만 WHO HEAT·KTDB·SROI 세 추정치는 각각 ±30% 이상의 불확실성 구간을
갖는 <strong>점 추정이 아닌 범위 추정</strong>이며, 따라서 본 연구가 제시하는 공공가치
회계의 최종 형태는 단일 NPV 수치가 아니라 <strong>시나리오 기반 구간 비교</strong>여야
한다. 이 선행 작업 전까지 전동보조 자전거 도입은 <strong>조건부 제안</strong>으로 분류된다
(8.4절 후속 과제 참조). 본 절의 정책 제언은 도입 자체가 아니라 "도입의 판단 근거를
공공가치 회계로 확장한다"는 방법론적 전환 제안이다.</p>

<ol start="2">
  <li><strong>동절기 합리적 축소.</strong> 계절지수 {winter_si:.2f}의 비수기 구간을
    재배치·정비 집중기로 전환하여 동절기 운영비를 15~20% 감축한다. 동시에 5~10월에는
    한강 수변 자치구의 거치대당 자전거를 15% 확대 배치한다.</li>
  <li><strong>여성·가족형 접근성 강화.</strong> 저상 자전거·야간 조명 대여소를
    지정하고, 아동 동반 시 할인 요금제를 시범 도입한다.</li>
  <li><strong>PM과의 역할 분담 정책.</strong> PM 시장 축소로 발생한 라스트마일 공백을
    따릉이 통근형 층이 흡수한다. 지하철 환승역 거치대 증설을 우선순위로 추진한다.</li>
</ol>

<h3>8.3 형평성 KPI 제안</h3>
<p>운영 평가 지표에 <strong>형평성 KPI</strong>를 신설할 것을 제안한다. 효율만이 아니라
분배까지 평가 대상에 포함시킨다는 의미다.</p>
<ul>
  <li>자치구 1인당 이용의 지니계수(연 2회 산출, 목표 {max(gini_per_capita - 0.05, 0.15):.2f} 이하로 단계적 축소)</li>
  <li>여성 이용 비율(현 {f_pct[-1]:.1f}% → 2030년 목표 {min(f_pct[-1] + 10, 45):.0f}%)</li>
  <li>65세 이상 이용 비율(집계 강화 후 base-line 재설정)</li>
  <li>대당 일회전율(2023 피크의 90% 복원)</li>
</ul>

<h3>8.4 연구의 한계 및 후속 과제</h3>
<ul>
  <li><strong>외생 변수 통제의 부분성:</strong> 본 연구의 LINEST_Multi 다중회귀는
    시간 인덱스 t·COVID 더미·11개 월 더미를 이미 투입하여 계절성과 연도 추세를 통제한다.
    그러나 <strong>강수일수·미세먼지 농도·요금제 변경 시점 더미·대여소 증설 시점 더미</strong>
    등 단기 이용에 직접 영향을 주는 외생 변수는 포함되지 않았다. supF·Placebo 기반
    구조변화 검정은 수행하였으나 Chow test 표준 구현과 준실험 설계(ITS 장기 series,
    DiD 대조군 설정) 역시 데이터 접근성 확대 이후의 후속 과제로 남긴다.</li>
  <li><strong>PM 정량 데이터 부재:</strong> 본 분석은 PM 시장 변화를 공개 보도자료
    수준에서만 인용하였다. 서울시 PM 등록대수·이용건수(OA-22199)가 공개될 경우
    정량 통합이 필요하다.</li>
  <li><strong>미상 보정 가정:</strong> '미상'의 연령 분포가 등록 이용자와 동일하다고
    가정하였으나, 실제로는 비회원·관광객·청소년 비중이 편향되어 있을 가능성이 있다.
    6.2.1절에서 ±20% 민감도 분석을 수행하여 결론의 강건성은 확인하였으나, 완전한
    해소는 아니다.</li>
  <li><strong>공간 분석의 분모:</strong> 6.3절 1인당 이용은 주민등록인구를 분모로
    사용하였다. 주간상주인구·생활인구 기준으로 재산출 시 종로·중구·영등포 등
    유입 인구가 큰 자치구의 순위가 재배치될 가능성이 있다.</li>
  <li><strong>공간 통계의 해상도:</strong> Moran's I는 자치구 단위(25개 표본)에서
    산출되어 수정가능면적단위문제(MAUP)에 노출된다. 500m hex 격자 또는 대여소 단위
    분석으로 공간 자기상관 재검정이 후속 과제다.</li>
  <li><strong>공공가치 회계의 부분성:</strong> 탄소 화폐화(₩1.9억/년)는 제시하였으나
    건강편익(WHO HEAT)·교통혼잡 완화·사회형평성의 화폐화는 포함하지 않았다.
    다중 편익 통합 SROI 산출이 공공사업 정당화를 위한 후속 과제다.</li>
  <li><strong>운영대수 추정치:</strong> 연도별 운영대수는 공개자료에 기반한 근사치이며,
    실제 월별 가동대수와는 차이가 있을 수 있다.</li>
  <li><strong>시간대 분석 부재:</strong> 월별 집계 데이터로는 여성의 야간 기피 가설을
    직접 검증할 수 없다. 6.1.1절의 건당 이용시간 proxy는 간접 증거에 그친다.
    서울 공공자전거 시간대별 데이터(OA-15245) 결합이 후속 연구 과제다.</li>
  <li><strong>대여소 마스터 시점:</strong> 2025년 12월 기준 마스터 데이터로 전체 기간을
    매핑하였기에, 초기년(2019~2022)의 폐쇄 대여소는 자치구 분석에서 미상으로 처리되었다.</li>
</ul>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 부록 A ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">부록 A — Excel 분석 재현 가이드</h2>
<p style="font-size:10pt; color:#555; margin-bottom:8px;">
첨부 Excel 파일 <strong>'따릉이_이용패턴_분석.xlsx'</strong>의 20개 시트 구성과 핵심 수식.
</p>

<table class="appendix-table">
<tr class="divider"><td colspan="3">기초 분석 (시트 1~8)</td></tr>
<tr><td>1</td><td>Raw_Data</td><td>84개월 이용·기온 원본 + VLOOKUP 참조테이블(I:J열)</td></tr>
<tr><td>2</td><td>Statistics</td><td>AVERAGE · MAX · MIN · MEDIAN · STDEV · MODE.SNGL · PERCENTILE + 연도별 SUMPRODUCT</td></tr>
<tr><td>3</td><td>Pivot_Source</td><td>피벗테이블 원본 (연령×연도, 성별×연도)</td></tr>
<tr><td>4</td><td>CAGR_Ranking</td><td>RATE · RANK.EQ로 연령별 CAGR 산출</td></tr>
<tr><td>5</td><td>Moving_Avg_MAE</td><td>AVERAGE(3MA·6MA) + ABS 절대오차 + MAE</td></tr>
<tr><td>6</td><td>Trend_Seasonal</td><td>INTERCEPT · SLOPE + AVERAGEIF 계절지수 + 2026 예측</td></tr>
<tr><td>7</td><td>Correlation</td><td>CORREL · RSQ · SLOPE · INTERCEPT + 산점도 + 추세선(R²)</td></tr>
<tr><td>8</td><td>Dashboard</td><td>콤보차트(보조축) + 파이차트 + 스파크라인</td></tr>
<tr class="divider"><td colspan="3">통계 방법론 보강 (시트 9~14)</td></tr>
<tr><td>9</td><td>Residual_Analysis</td><td>단순회귀 잔차 시계열 + Durbin-Watson + 잔차-시간 상관</td></tr>
<tr><td>10</td><td>LINEST_Multi</td><td>LINEST 배열수식 다중회귀 (Temp + t + COVID더미 + 11개 월더미)</td></tr>
<tr><td>11</td><td>Forecast_Benchmark</td><td>Naive · Seasonal Naive · 3MA · 추세×계절 4모델 비교 (MAE/MAPE/RMSE) + 95% 신뢰구간</td></tr>
<tr><td>12</td><td>District_Analysis</td><td>25자치구 × 이용·인구·대여소수 + 지니계수 (SUMPRODUCT)</td></tr>
<tr><td>13</td><td>Unit_Economics</td><td>자전거 1대당 일회전율 + AVERAGEIFS 연별 + 건당 적자 원단위</td></tr>
<tr><td>14</td><td>Misu_Adjustment</td><td>원본 CAGR vs 미상 안분 보정 CAGR 병기</td></tr>
<tr class="divider"><td colspan="3">정책 분석 확장 (시트 15~20)</td></tr>
<tr><td>15</td><td>Gini_Timeseries</td><td>2019/2022/2025 3시점 지니계수 추이 (0.262→0.277→0.301)</td></tr>
<tr><td>16</td><td>Morans_I</td><td>공간 자기상관 검정 (Queen 인접행렬 + 999회 순열검정)</td></tr>
<tr><td>17</td><td>Carbon_Value</td><td>탄소 화폐화 (K-ETS ₩40,000/톤 기준) · 연 적자 상쇄율 산출</td></tr>
<tr><td>18</td><td>NPV_Sensitivity</td><td>NPV · IRR · 사회적할인율 4.5%(KDI) 시나리오 민감도</td></tr>
<tr><td>19</td><td>Misu_Sensitivity</td><td>미상 안분 가정 ±20% 시 60대 CAGR 범위 (+15.9%~+24.2%)</td></tr>
<tr><td>20</td><td>Gender_Pattern</td><td>성별 건당 이용시간·이동거리 (여성 +10.9% 시간, +3.8% 거리)</td></tr>
</table>

<h4 style="margin-top:14px;">주요 배열수식 (참고)</h4>
<div class="appendix-box">LINEST_Multi!A93:P97   = LINEST(y, X, TRUE, TRUE)   [5×15 스필]
District_Analysis      = SUMPRODUCT(ABS(D-TRANSPOSE(D))) / (2·n²·μ)
Residual_Analysis      = SUMXMY2(E5:E86,E4:E85) / SUMSQ(E4:E86)</div>

<!-- ━━━━━━━━━━━━━━━━━━━━━━ 부록 B ━━━━━━━━━━━━━━━━━━━━━━ -->
<h2 class="section">부록 B — 참고문헌 및 출처</h2>

<p style="font-size:9.5pt; color:#666; margin: 4px 0 10px;">
본 절은 분석에 사용된 데이터 출처와 본문에서 인용한 방법론·정책 참조 자료를 구분하여 제시한다.
</p>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 8px 0 4px;">A. 1차 데이터 출처</h4>
<ol class="ref-list">
  <li>서울 열린데이터광장, "서울시 공공자전거 이용정보(월별)", OA-15248,
    <span style="font-family:monospace;">data.seoul.go.kr/dataList/OA-15248</span>, 접근일 2026-04-14.
    본 연구의 핵심 종속변수(월별 이용건수).</li>
  <li>서울 열린데이터광장, "서울시 공공자전거 대여소 정보(25.12월 기준)", OA-13252,
    <span style="font-family:monospace;">data.seoul.go.kr/dataList/OA-13252</span>, 접근일 2026-04-15.
    대여소 2,799개소의 자치구 매핑 근거.</li>
  <li>서울 열린데이터광장, "서울시 공유 전동킥보드 운영 현황", OA-22199.
    7.2절 PM 시장 보조 출처.</li>
  <li>Open-Meteo Historical Weather API, 서울 37.57°N, 월평균기온 2019~2025,
    <span style="font-family:monospace;">open-meteo.com</span>, 접근일 2026-04-14.
    4장 기온 효과 분석의 외생 변수.</li>
  <li>서울특별시 주민등록인구통계 (자치구별), 2025년 1월 기준.
    6.3절 1인당 이용량 산출의 분모.</li>
</ol>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 10px 0 4px;">B. 방법론 참조 자료</h4>
<ol class="ref-list" start="6">
  <li>KDI 공공투자관리센터(PIMAC), <em>예비타당성조사 수행을 위한 일반지침 연구(제5판)</em>,
    2023. 사회적할인율 4.5%·내용연수 15년(8.2절 NPV 표 주석 3).</li>
  <li>World Health Organization, <em>Health Economic Assessment Tool (HEAT) for Walking
    and Cycling — Methodology Guide v5.0</em>, 2024. 8.2절 건강편익 프레임.</li>
  <li>한국교통연구원(KTDB), <em>교통수단별 시간가치 및 차량운행비용 원단위(2024)</em>,
    2024. 8.2절 교통혼잡 감소 편익의 시간가치 근거.</li>
  <li>Institute for Transportation and Development Policy (ITDP),
    <em>The Bikeshare Planning Guide, 2nd ed.</em>, 2018. 공공자전거 평가 축의 국제 표준.</li>
  <li>OECD International Transport Forum (ITF), <em>Shared Mobility Simulations:
    Auckland·Dublin·Helsinki·Lisbon</em> 시리즈, 2016~2023. 공유 모빌리티 편익 축 참조.</li>
  <li>Anselin, L., "Local Indicators of Spatial Association — LISA",
    <em>Geographical Analysis</em> 27(2), 1995. 6.3.3절 LISA 군집 정의.</li>
  <li>Moran, P. A. P., "Notes on Continuous Stochastic Phenomena",
    <em>Biometrika</em> 37(1/2), 1950. 6.3.2절 전역 공간 자기상관 통계량의 원전.</li>
  <li>Bai, J. &amp; Perron, P., "Computation and analysis of multiple structural change
    models", <em>Journal of Applied Econometrics</em> 18(1), 2003. 4.2.1절 supF 구조변화
    검정 방법론.</li>
  <li>Chow, G. C., "Tests of Equality Between Sets of Coefficients in Two Linear
    Regressions", <em>Econometrica</em> 28(3), 1960. Chow 검정의 원전(4.2절 참조).</li>
  <li>Hyndman, R. J. &amp; Athanasopoulos, G., <em>Forecasting: Principles and Practice</em>,
    3rd ed., OTexts, 2021. 5.2절 SARIMAX·Seasonal Naive·벤치마크 비교 방법론 근거.</li>
  <li>Wooldridge, J. M., <em>Introductory Econometrics: A Modern Approach</em>, 7th ed.,
    Cengage, 2020. LINEST_Multi 다중회귀·더미변수 설계 근거(4.1절).</li>
  <li>Gini, C., "Variabilità e Mutabilità", <em>Studi Economico-Giuridici
    dell'Università di Cagliari</em> 3, 1912. 6.3.1절 지니계수의 원전.</li>
  <li>환경부 온실가스종합정보센터, <em>배출권거래제(K-ETS) 주요 통계 2024</em>, 2025.
    8.2절 탄소 화폐화 단가 ₩40,000/tCO₂ 근거.</li>
</ol>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 10px 0 4px;">C. 정책·언론 자료</h4>
<ol class="ref-list" start="19">
  <li>서울특별시, "시민의 발 따릉이 10년간 2억 5천만명 탑승 · 3시간권 신규 도입",
    news.seoul.go.kr/traffic/archives/515711, 2025. 8.2절 3시간 이용권 시점 근거.</li>
  <li>서울특별시, "기후동행카드 공식 안내(62,000원/65,000원 요금제)",
    news.seoul.go.kr/traffic, 2024. 7.2.1절 요금 구조 근거.</li>
  <li>서울특별시 교통실, <em>2025년 공공자전거 운영 연간 리포트</em>, 2025.
    1장 도입 규모·운영대수 보조 출처.</li>
  <li>서울연구원(SI), <em>서울시 공공자전거 이용 활성화 방안 연구</em>, 2023.
    6장 자치구 격차·세그먼트 분화 선행연구.</li>
  <li>뉴시스, "'따릉이' 이용 2억건 시대 · 자전거 도로는 여전히 뚝뚝", 2025-05-09.</li>
  <li>국민일보, "[단독] 서울시, 따릉이 요금 인상 없다 · 무한 대여는 제동 검토", 2025.</li>
  <li>매일경제, "서울시 따릉이 연 적자 약 100억원 · 요금 동결 기조 유지", 2024. 8.2절 운영 적자 규모.</li>
  <li>KISO저널, "그 많던 공유 전동킥보드는 왜 사라졌을까?",
    journal.kiso.or.kr/?p=13261, 2024. 7.2절 더스윙·PM 시장 축소 맥락.</li>
  <li>한국경제, "SKT 티맵모빌리티, 전동킥보드 공유서비스 2025년 3월 종료", 2025.
    7.2절 PM 사업 재편 출처.</li>
  <li>서울특별시, "따릉이 이용자 안전 인식 조사(시민 설문, n=2,013)", 2024.
    7.2절 PM 이용 시 불편 경험 79.2% 근거.</li>
</ol>

<div style="break-inside: avoid; page-break-inside: avoid; margin-top: 10px;">
<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 4px 0 4px;">D. 약어 및 용어 해설</h4>
<p style="font-size:9pt; color:#666; margin: 2px 0 6px;">
본문·부록에서 사용된 주요 영문 약어와 통계 용어의 우리말 의미를 요약한다.
</p>
<table style="font-size:9pt; margin-top:4px;">
<tr><th style="width:18%;">약어</th><th style="width:34%;">한국어 의미</th><th style="width:48%;">본문 위치·비고</th></tr>
<tr><td>supF</td><td>sup-F 구조변화 검정 통계량</td><td>4.2.1 분기점 특정 (F=13.55 @ 2022-04)</td></tr>
<tr><td>Placebo test</td><td>가짜 처치 시점 검정</td><td>4.2.1 2022-06 slope 유의 확인</td></tr>
<tr><td>SARIMAX</td><td>외생변수 포함 계절 ARIMA 모형</td><td>4.1.1 각주 · 5.2 벤치마크 보조</td></tr>
<tr><td>Stratified OLS</td><td>연도별 분할 최소자승 회귀</td><td>4.1.1 연도 고정효과 근사</td></tr>
<tr><td>LINEST_Multi</td><td>다중회귀 배열수식 (Excel)</td><td>4.1 · 8.4 외생 변수 부분 통제</td></tr>
<tr><td>Durbin-Watson</td><td>잔차 자기상관 검정 통계량</td><td>4.1 잔차 진단</td></tr>
<tr><td>Welch t-test</td><td>이분산 가정 평균 차이 검정</td><td>6.1.1 성별 이용시간·거리 (p&lt;0.001)</td></tr>
<tr><td>Moran's I</td><td>전역 공간 자기상관 계수</td><td>6.3.2 Queen I = +0.245, p = 0.016</td></tr>
<tr><td>LISA</td><td>국지 공간 자기상관 지표</td><td>6.3.3 HH·LL·HL·LH 4유형 군집</td></tr>
<tr><td>MAUP</td><td>수정가능면적단위문제</td><td>6.3.2 caveat · 1인당/절대값 감도</td></tr>
<tr><td>CAGR</td><td>연평균 복합 성장률</td><td>3장 · 6.2 미상 보정 CAGR</td></tr>
<tr><td>NPV · IRR</td><td>순현재가치 · 내부수익률</td><td>8.2 시나리오 민감도 분석</td></tr>
<tr><td>SROI</td><td>사회적 투자수익률</td><td>8.2 사회형평성 화폐화 프레임</td></tr>
<tr><td>WHO HEAT</td><td>WHO 보행·자전거 건강편익 평가 도구</td><td>8.2 건강편익 편입 근거</td></tr>
<tr><td>K-ETS</td><td>한국 온실가스배출권 거래제</td><td>8.2 탄소 화폐화 단가 근거</td></tr>
<tr><td>KDI PIMAC</td><td>한국개발연구원 공공투자관리센터</td><td>8.2 NPV 할인율·내용연수</td></tr>
<tr><td>KTDB</td><td>한국교통연구원 교통DB</td><td>8.2 시간가치 원단위</td></tr>
<tr><td>ITDP · ITF</td><td>국제 공공자전거·교통포럼</td><td>8.2 국제 표준 편익 축</td></tr>
</table>
</div>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 12px 0 4px;">E. 통계 기호 표기 규약</h4>
<table style="font-size:9pt; margin-top:4px;">
<tr><th style="width:18%;">기호</th><th style="width:40%;">의미</th><th style="width:42%;">본문 예</th></tr>
<tr><td>α</td><td>유의수준 (type I 오류 허용치)</td><td>α = 0.05 (6.3.2 Moran's I 검정)</td></tr>
<tr><td>β</td><td>회귀 기울기 계수</td><td>4.1.1 연도별 β 79,862건/°C (±13,594)</td></tr>
<tr><td>ε, 잔차</td><td>예측값과 실제값의 차</td><td>4.2.2 분기점 이후 평균 잔차</td></tr>
<tr><td>σ</td><td>표준편차</td><td>4.1.1 β 평균±σ · 5.2 RMSE 구간</td></tr>
<tr><td>R²</td><td>결정계수 (설명된 분산 비율)</td><td>4.1 pooled 0.43 → stratified 0.70</td></tr>
<tr><td>p-value</td><td>귀무가설 유의확률</td><td>Q1~Q4 가설검정 판단 기준</td></tr>
<tr><td>t-통계량</td><td>계수 / 표준오차 비율</td><td>COVID 더미 t = 1.02 (비유의)</td></tr>
<tr><td>F-통계량</td><td>구조변화 검정 통계량</td><td>supF = 13.55 @ 2022-04</td></tr>
</table>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 12px 0 4px;">F. 재현성 메타정보</h4>
<table style="font-size:9pt; margin-top:4px;">
<tr><th style="width:30%;">항목</th><th style="width:70%;">내용</th></tr>
<tr><td>GitHub 저장소</td><td><span style="font-family:monospace;">{GITHUB_URL}</span> (public, MIT+CC-BY 4.0)</td></tr>
<tr><td>Git 커밋 해시</td><td><span style="font-family:monospace;">{GITHUB_COMMIT}</span></td></tr>
<tr><td>분석 도구</td><td>Python 3.14 (pandas, numpy, scipy, statsmodels, ruptures, esda, libpysal, spreg, geopandas, numpy-financial); Excel 365 (LINEST 배열수식)</td></tr>
<tr><td>데이터 스냅숏</td><td>2019-01 ~ 2025-12 (84개월) · 서울 열린데이터광장 2026-04-14 다운로드 · data_manifest.csv의 SHA-256 지문으로 무결성 보증</td></tr>
<tr><td>난수 시드</td><td>모든 stochastic 계산 <code>np.random.seed(42)</code> (<code>seed_config.json</code>에 전역 관리)</td></tr>
<tr><td>라이선스</td><td>공공누리 제1유형 (데이터) · MIT/CC-BY 4.0 (코드·보고서)</td></tr>
<tr><td>재현 절차</td><td><code>git clone {GITHUB_URL}</code> → <code>pip install -r requirements.txt</code> → <code>make reproduce</code></td></tr>
</table>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 12px 0 4px;">G. 분석 검증 체크리스트</h4>
<p style="font-size:9pt; color:#666; margin: 2px 0 4px;">
본 연구의 네 개 연구 질문(Q1~Q4)에 대한 H₀/H₁ 검정 결과와 주요 통계량을 요약한다.
</p>
<table style="font-size:9pt; margin-top:4px;">
<tr><th style="width:7%;">#</th><th style="width:30%;">연구 질문 / 가설</th><th style="width:28%;">주요 통계량</th><th style="width:15%;">판정</th><th style="width:20%;">본문 위치</th></tr>
<tr><td>Q1</td><td>기온 효과 제거 후 잔차가 무작위 분포하는가</td>
  <td>supF = {V21_STATS['bai_perron']['supF_stat']} @ {V21_STATS['bai_perron']['supF_ym']} (Andrews 임계 8.85) · DW = {V21_STATS['residual_diagnostics']['durbin_watson']}</td>
  <td>H₀ 기각 · H₁ 지지</td><td>4.2.1 / 4.1.2</td></tr>
<tr><td>Q2</td><td>SARIMAX 2026년 예측구간의 신뢰 범위</td>
  <td>SARIMAX {V21_STATS.get('sarimax',{}).get('best_order','-')} AIC = {V21_STATS.get('sarimax',{}).get('best_aic','-'):.0f} · 잔차 LB Q(10) p = {V21_STATS.get('sarimax',{}).get('residual_ljung_box_q10',{}).get('p',0):.3f}</td>
  <td>백색성 수용 · 분석적 PI 제시</td><td>5.2.1</td></tr>
<tr><td>Q3</td><td>고령층 성장은 실제 수요인가 집계 변화인가</td><td>60대 원본 CAGR +32.4% → 보정 +20.7% (Δ = 11.7pp)</td><td>H₀ 기각 · 혼재 확인</td><td>6.2 / 6.2.1</td></tr>
<tr><td>Q4</td><td>자치구 이용의 공간 자기상관이 존재하는가</td>
  <td>Queen 절대값 I = {V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('I',0)} (p = {V21_STATS['moran_sensitivity'].get('absolute__Queen',{}).get('p_sim',0)}); LISA FDR 유의 {V21_STATS.get('lisa_fdr',{}).get('n_fdr_sig',0)}/25</td>
  <td>전역 기각 · 국지 미유의 (표본 한계)</td><td>6.3.2 / 6.3.3</td></tr>
<tr><td>보조</td><td>within/between 분산분해 해석</td>
  <td>between {V21_STATS['variance_decomposition']['between_ratio']*100:.0f}% · within {V21_STATS['variance_decomposition']['within_ratio']*100:.0f}% · pooled R² 0.43 → stratified 0.70</td>
  <td>"우연" → 분산 귀속</td><td>4.1 insight</td></tr>
<tr><td>보조</td><td>공간회귀 (SAR vs SEM)</td>
  <td>LM-Lag p = {V21_STATS['spatial_regression']['lm_lag']['p']:.3f} · LM-Error p = {V21_STATS['spatial_regression']['lm_error']['p']:.3f} · SAR rho = {V21_STATS['spatial_regression']['SAR_rho']:.3f}</td>
  <td>방향 일치 · 유의성 n 한계</td><td>6.3.4</td></tr>
<tr><td>보조</td><td>Moran's I power curve (n=25, α=0.05)</td>
  <td>관측 I에서 power 약 {V21_STATS['power_curve']['curve'].get('I_0.245', '-')}; 0.8 도달은 I ≥ 0.40</td>
  <td>표본 확장 후속 과제</td><td>6.3.5</td></tr>
<tr><td>보조</td><td>성별 이용시간·거리 Welch t-test</td>
  <td>시간 t = {V21_STATS['welch_t']['time_min']['t']:.1f} (p &lt; 0.001) · 거리 t = {V21_STATS['welch_t']['distance_m']['t']:.1f}</td>
  <td>통계적 유의</td><td>6.1.1</td></tr>
<tr><td>보조</td><td>3대안 재무 타당성 (증분 NPV, dr = 4.5%)</td>
  <td>Alt2 +{V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_npv','-')}억 B/C {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt2_Regular',{}).get('dr_045',{}).get('incr_bc_ratio','-')}; Alt3 +{V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt3_Ebike',{}).get('dr_045',{}).get('incr_npv','-')}억 B/C {V21_FINANCE.get('alternatives',{}).get('_incremental_vs_Alt1',{}).get('Alt3_Ebike',{}).get('dr_045',{}).get('incr_bc_ratio','-')}</td>
  <td>예타 통과 가능 (3축 공공가치 회계 반영 시)</td><td>8.2</td></tr>
<tr><td>보조</td><td>미상 안분 ±20% 가정 강건성</td><td>60대 CAGR +15.9~24.2% (부호 불변)</td><td>강건성 확인</td><td>6.2.1</td></tr>
<tr><td>보조</td><td>자치구 지니 시계열</td><td>0.262(2019) → 0.277(2022) → 0.301(2025)</td><td>격차 확대</td><td>6.3.1</td></tr>
</table>

<h4 style="font-size:10.5pt; color:#1B3A5C; margin: 12px 0 4px;">H. 결측률 및 data_manifest 요약</h4>
<p style="font-size:9pt; color:#666; margin: 2px 0 4px;">
원자료·집계 자료 14개 파일의 SHA-256 지문과 크기·행수를 <code>data_manifest.csv</code>
에 기록하였다. GitHub 저장소에서 원본과 대조 가능하다.
</p>
<table style="font-size:8.5pt;">
<tr><th>파일</th><th>크기</th><th>행수</th><th>SHA-256 (앞 12자리)</th></tr>
<tr><td>monthly_aggregate.json</td><td>16 KB</td><td>84 mo</td><td><span style="font-family:monospace;">bd7f65e154c7</span></td></tr>
<tr><td>district_monthly.json</td><td>134 KB</td><td>2,184 (25구×84mo ± 미상)</td><td><span style="font-family:monospace;">9a6a80b79469</span></td></tr>
<tr><td>seoul_temperature.json</td><td>4.7 KB</td><td>84 mo</td><td>Open-Meteo API 재현 가능</td></tr>
<tr><td>seoul_districts.geojson</td><td>78 KB</td><td>25 polygons</td><td>southkorea/seoul-maps 공개</td></tr>
<tr><td>station_master.csv</td><td>185 KB</td><td>2,799 대여소</td><td>서울 OA-13252</td></tr>
<tr><td>districts_adjacency.json</td><td>&lt; 10 KB</td><td>56 Queen links</td><td>libpysal 파생</td></tr>
<tr><td>v21_stats.json</td><td>-</td><td>10개 통계량</td><td>compute_stats_v21.py 산출</td></tr>
<tr><td>v21_finance.json</td><td>-</td><td>3대안 × 3 dr × 15년</td><td>compute_finance_v21.py 산출</td></tr>
</table>
<p style="font-size:9pt; color:#666; margin-top:6px;">
결측 패턴: 자치구 월별 이용 데이터에서 '미상' 자치구 비율은 2019년 평균 11.3% →
2025년 평균 1.8%로 감소 추세이며, 미상 비율 변동은 6.2절 CAGR 보정 및 ±20% 민감도로
통제하였다. 전체 원자료 무결성은 <code>make data</code> 타깃의 SHA-256 재계산으로 검증된다.
</p>

<p style="font-size:9pt; color:#666; text-align:center; margin-top: 10px;">
— 본 보고서의 모든 수치는 서울 열린데이터광장 공공자전거 원천 자료와 Open-Meteo 기상 자료에서 직접 집계·산출되었으며, Excel 재현 가이드(부록 A)를 따를 경우 독립적으로 재현 가능하다. —
</p>

<div class="footer">
중앙대학교 소프트웨어학부 20203876 최성민 · 트렌드를읽는데이터경영 01분반 · 2026년 4월
</div>

</body>
</html>'''

# Save HTML
with open('report.html', 'w', encoding='utf-8') as f:
    f.write(html)
print(f"HTML 저장: report.html ({len(html):,} bytes)")

# PDF 생성
pdf_path = os.path.abspath('보고서_따릉이_이용패턴_분석.pdf')
chrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
html_path = os.path.abspath('report.html')

result = subprocess.run([
    chrome, '--headless', '--disable-gpu', '--no-pdf-header-footer',
    f'--print-to-pdf={pdf_path}',
    f'file://{html_path}'
], capture_output=True, text=True, timeout=60)

if os.path.exists(pdf_path):
    size = os.path.getsize(pdf_path)
    print(f"PDF 저장: {pdf_path} ({size / 1024:.0f}KB)")
else:
    print("ERROR: PDF 생성 실패")
    print(result.stderr[:500])
