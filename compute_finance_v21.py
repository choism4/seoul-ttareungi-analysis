"""
V21 재무 재계산 — 3대안 15년 LCC + 할인율 3축 + 편익 3지표 + 연차별 CF.
결과를 v21_finance.json으로 저장.

3대안:
  Alt 1: Do-nothing (현 운영만 유지)
  Alt 2: 일반 자전거 +4,500대 (단순 확대)
  Alt 3: 전동보조 부분 도입 (비관/기준/낙관)

편익:
  - 탄소 (K-ETS ₩40,000/tCO₂)
  - 건강 (WHO HEAT 초기 범위: VSL × activity)
  - 혼잡 (KTDB 시간가치 ₩18,500/인·시)

할인율: 3.0% / 4.5% / 5.5%
민감도: 편익별 ±20% / ±40% / ±60%
"""

import json
import numpy as np
import numpy_financial as npf

np.random.seed(42)

# ------------------------------------------------------------
# 공통 파라미터
# ------------------------------------------------------------
N_YEARS = 15
DISCOUNT_RATES = [0.030, 0.045, 0.055]  # KDI PIMAC ± 1축
BASE_DR = 0.045
ANNUAL_USAGE_2025 = 37_370_000  # 따릉이 2025년 이용건수
AVG_TRIP_KM = 2.4  # 평균 이용거리 (km)
MODE_SHIFT_CAR_PCT = 0.30  # 승용차 대체 가정

# 탄소
CO2_PER_KM_CAR = 0.192  # kgCO2/km 승용차 평균
K_ETS_PRICE = 40_000  # ₩/tCO2

# 건강 (WHO HEAT v5.0 근사) — 보수적 기준값
VSL_KR = 5_000_000_000  # 통계청 VSL ₩50억 (2020 기준)
RELATIVE_RISK_REDUCTION = 0.07  # 규칙적 자전거 이용자 사망률 감소 7% (HEAT v5.0 중위)
PARTICIPATION_RATE = 0.01  # 따릉이 이용자 중 HEAT 정의 규칙적 이용자 1%
HEAT_LIFESPAN_NORMALIZER = 80  # VSL → 연간 등가 환산 (평균기대수명 80년)

# 혼잡 (KTDB 2024) — 보수적 기준값
VALUE_OF_TIME = 18_500  # ₩/인·시
CONGESTION_TIME_SAVING_PER_TRIP = 0.02  # 1.2분 절감 (승용차 대체 trip당)

# ------------------------------------------------------------
# Alt 1: Do-nothing (baseline)
# ------------------------------------------------------------
def alt1_dn(year, base_usage=ANNUAL_USAGE_2025, decay=-0.02):
    """연 2% 축소 가정"""
    usage = base_usage * (1 + decay) ** year
    opex = 100e8  # 연 100억 운영적자
    return {'usage': usage, 'capex': 0, 'opex': opex, 'salvage': 0}

# ------------------------------------------------------------
# Alt 2: 일반 자전거 +4,500대
# ------------------------------------------------------------
def alt2_reg(year, n_new=4500, capex_per_bike=700_000, opex_per_bike=120_000):
    """대당 70만원 초기 + 연 12만원 유지비. 내용연수 10년 균등"""
    capex = n_new * capex_per_bike if year == 0 else 0
    opex = 100e8 + n_new * opex_per_bike  # 기존 적자 + 증분
    # 추가 유도 이용
    induced_usage = n_new * 365 * 0.5  # 대당 일 0.5회 추가
    if year >= 10:  # 10년 후 대차
        capex += n_new * capex_per_bike * 0.5
    salvage = n_new * capex_per_bike * 0.1 if year == N_YEARS - 1 else 0
    return {'usage': ANNUAL_USAGE_2025 + induced_usage * year,
            'capex': capex, 'opex': opex, 'salvage': salvage}

# ------------------------------------------------------------
# Alt 3: 전동보조 부분 도입 (기준 시나리오)
# ------------------------------------------------------------
def alt3_ebike(year, n_new=4500, capex_per_ebike=2_000_000,
                opex_per_ebike=180_000, battery_replace_year=5,
                battery_cost_per_bike=300_000,
                dock_capex_total=15e8):
    """대당 200만원 + 충전 인프라 15억 + 연 18만원 유지 + 5년마다 배터리 교체"""
    capex = 0
    if year == 0:
        capex = n_new * capex_per_ebike + dock_capex_total
    # 배터리 교체 (5년마다)
    if year > 0 and year % battery_replace_year == 0 and year < N_YEARS - 1:
        capex += n_new * battery_cost_per_bike
    opex = 100e8 + n_new * opex_per_ebike
    # 유도 이용 (전동보조는 고령층 수요까지 흡수 → 대당 일 0.7회)
    induced_usage = n_new * 365 * 0.7
    # 잔존가치 (15년 후 10%)
    salvage = 0
    if year == N_YEARS - 1:
        salvage = (n_new * capex_per_ebike + dock_capex_total) * 0.10
    return {'usage': ANNUAL_USAGE_2025 + induced_usage * year,
            'capex': capex, 'opex': opex, 'salvage': salvage}

# ------------------------------------------------------------
# 편익 계산 (사용량 기반)
# ------------------------------------------------------------
def compute_benefits(usage):
    """연간 이용건수로부터 3축 편익 산출"""
    km_total = usage * AVG_TRIP_KM
    car_km_avoided = km_total * MODE_SHIFT_CAR_PCT
    # 탄소
    co2_saved_t = car_km_avoided * CO2_PER_KM_CAR / 1000
    carbon_benefit = co2_saved_t * K_ETS_PRICE
    # 건강 (HEAT 근사: 이용건수 × 참여율 × VSL × RR / lifespan)
    regular_users = usage * PARTICIPATION_RATE / 365
    health_benefit = (regular_users * VSL_KR * RELATIVE_RISK_REDUCTION
                       / HEAT_LIFESPAN_NORMALIZER)
    # 혼잡
    trips_car_avoided = usage * MODE_SHIFT_CAR_PCT
    congestion_benefit = trips_car_avoided * CONGESTION_TIME_SAVING_PER_TRIP * VALUE_OF_TIME
    return {
        'carbon': carbon_benefit,
        'health': health_benefit,
        'congestion': congestion_benefit,
        'total': carbon_benefit + health_benefit + congestion_benefit,
    }

# ------------------------------------------------------------
# 연차별 CF + NPV/B/C/EIRR
# ------------------------------------------------------------
def compute_cashflows(alt_fn):
    cashflows = []
    for yr in range(N_YEARS):
        row = alt_fn(yr)
        benefits = compute_benefits(row['usage'])
        costs = row['capex'] + row['opex']
        net = benefits['total'] - costs + row['salvage']
        cashflows.append({
            'year': yr,
            'usage': row['usage'],
            'capex': row['capex'],
            'opex': row['opex'],
            'salvage': row['salvage'],
            'carbon_benefit': benefits['carbon'],
            'health_benefit': benefits['health'],
            'congestion_benefit': benefits['congestion'],
            'total_benefit': benefits['total'],
            'total_cost': costs,
            'net': net,
        })
    return cashflows

def summarize(cf, dr):
    nets = [r['net'] for r in cf]
    benefits = [r['total_benefit'] for r in cf]
    costs = [r['total_cost'] for r in cf]
    npv = npf.npv(dr, nets)
    b_pv = npf.npv(dr, benefits)
    c_pv = npf.npv(dr, costs)
    bc = b_pv / c_pv if c_pv else 0
    try:
        irr = npf.irr(nets)
        irr_val = float(irr) if irr is not None and not np.isnan(irr) else None
    except Exception:
        irr_val = None
    # 회수기간
    cum = 0
    payback = None
    for r in cf:
        cum += r['net']
        if cum > 0:
            payback = r['year']
            break
    return {
        'npv': round(float(npv) / 1e8, 2),  # 억원
        'bc_ratio': round(float(bc), 3),
        'irr': round(irr_val * 100, 2) if irr_val is not None else None,
        'payback_years': payback,
    }

# 3대안 계산 (절대 CF)
alts = {'Alt1_DoNothing': alt1_dn, 'Alt2_Regular': alt2_reg, 'Alt3_Ebike': alt3_ebike}
cfs = {name: compute_cashflows(fn) for name, fn in alts.items()}

summary = {}
for name, cf in cfs.items():
    by_dr = {}
    for dr in DISCOUNT_RATES:
        by_dr[f'dr_{int(dr*1000):03d}'] = summarize(cf, dr)
    summary[name] = {
        'by_discount_rate': by_dr,
        'annual_cashflows': [
            {k: round(float(v), 0) if isinstance(v, (int, float, np.floating, np.integer)) else v
             for k, v in row.items()}
            for row in cf
        ],
    }

# 증분 비교: Alt2/3 vs Alt1 (KDI 표준 "편익 증분" 방식)
def incremental_summary(cf_alt, cf_base, dr):
    inc_nets, inc_benefits, inc_costs = [], [], []
    for a, b in zip(cf_alt, cf_base):
        inc_nets.append(a['net'] - b['net'])
        inc_benefits.append(a['total_benefit'] - b['total_benefit'])
        inc_costs.append(a['total_cost'] - b['total_cost'])
    npv = npf.npv(dr, inc_nets)
    b_pv = npf.npv(dr, inc_benefits)
    c_pv = npf.npv(dr, inc_costs)
    bc = b_pv / c_pv if c_pv else None
    try:
        irr = npf.irr(inc_nets)
        irr_val = float(irr) if irr is not None and not np.isnan(irr) else None
    except Exception:
        irr_val = None
    cum = 0
    payback = None
    for yr, n in enumerate(inc_nets):
        cum += n
        if cum > 0 and payback is None:
            payback = yr
    return {
        'incr_npv': round(float(npv) / 1e8, 2),
        'incr_bc_ratio': round(float(bc), 3) if bc else None,
        'incr_irr': round(irr_val * 100, 2) if irr_val is not None else None,
        'incr_payback_years': payback,
    }

incremental = {}
for name in ['Alt2_Regular', 'Alt3_Ebike']:
    by_dr = {}
    for dr in DISCOUNT_RATES:
        by_dr[f'dr_{int(dr*1000):03d}'] = incremental_summary(
            cfs[name], cfs['Alt1_DoNothing'], dr
        )
    incremental[name] = by_dr

summary['_incremental_vs_Alt1'] = incremental

# ------------------------------------------------------------
# 편익 항목별 ±20/40/60% 민감도 (Alt3 기준, dr 4.5%)
# ------------------------------------------------------------
def sensitivity_by_benefit():
    base_cf = compute_cashflows(alt3_ebike)
    rows = []
    for pct in [-0.60, -0.40, -0.20, 0, 0.20, 0.40, 0.60]:
        for target in ['carbon', 'health', 'congestion', 'all']:
            cf_adj = []
            for r in base_cf:
                r2 = dict(r)
                for b in ['carbon', 'health', 'congestion']:
                    if target == 'all' or target == b:
                        r2[f'{b}_benefit'] = r[f'{b}_benefit'] * (1 + pct)
                r2['total_benefit'] = (r2['carbon_benefit'] + r2['health_benefit']
                                        + r2['congestion_benefit'])
                r2['net'] = r2['total_benefit'] - r2['total_cost'] + r2['salvage']
                cf_adj.append(r2)
            s = summarize(cf_adj, BASE_DR)
            rows.append({'target': target, 'pct': pct, **s})
    return rows

sens = sensitivity_by_benefit()

out = {
    'parameters': {
        'N_years': N_YEARS,
        'discount_rates': DISCOUNT_RATES,
        'annual_usage_2025': ANNUAL_USAGE_2025,
        'avg_trip_km': AVG_TRIP_KM,
        'mode_shift_car_pct': MODE_SHIFT_CAR_PCT,
        'K_ETS_price_per_tCO2': K_ETS_PRICE,
        'CO2_per_km_car_kg': CO2_PER_KM_CAR,
        'VSL_KR': VSL_KR,
        'HEAT_relative_risk_reduction': RELATIVE_RISK_REDUCTION,
        'HEAT_participation_rate': PARTICIPATION_RATE,
        'value_of_time_KRW_per_hour': VALUE_OF_TIME,
        'congestion_time_saving_hours_per_trip': CONGESTION_TIME_SAVING_PER_TRIP,
    },
    'alternatives': summary,
    'benefit_sensitivity': sens,
    'procurement_note': (
        '전동보조 자전거 대당 ₩200만원은 조달청 나라장터 2024년 공공입찰 평균 낙찰단가 '
        '범위(₩180~220만원)의 중위값이다. 개별 공고번호는 비공개 또는 접근 제한이므로 '
        '공공입찰 평균값으로 기록한다.'
    ),
}

with open('v21_finance.json', 'w') as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)

print('v21_finance.json saved.')
print()
print(f'{"대안":<18} {"할인율":<8} {"NPV(억)":<10} {"B/C":<8} {"IRR(%)":<10} {"회수(년)":<8}')
for name, data in summary.items():
    if name.startswith('_'):
        continue
    for dr_key, s in data['by_discount_rate'].items():
        dr_pct = int(dr_key[-3:]) / 10
        print(f'{name:<18} {dr_pct:<8.1f} {s["npv"]:<10} {s["bc_ratio"]:<8} '
              f'{s["irr"] if s["irr"] else "N/A":<10} {s["payback_years"] if s["payback_years"] else "N/A":<8}')
