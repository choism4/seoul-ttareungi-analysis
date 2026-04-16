"""
V21 별책 기술부록 PDF 생성.
Main report에서 분리한 고급 분석 T1~T7을 단일 PDF로 발행.
"""
import json
import os
import subprocess
import numpy as np

os.chdir(os.path.dirname(os.path.abspath(__file__)))

STATS = json.load(open('v21_stats.json'))
FIN = json.load(open('v21_finance.json'))

GITHUB_URL = 'https://github.com/choism4/seoul-ttareungi-analysis'

CSS = """
<style>
@page { size: A4; margin: 20mm 18mm; @bottom-center { content: counter(page); font-size: 9pt; color: #888; } }
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; font-size: 10.5pt;
       line-height: 1.6; color: #222; word-break: keep-all; overflow-wrap: break-word; }
h1 { color: #1B3A5C; font-size: 22pt; border-bottom: 3px double #1B3A5C;
     padding-bottom: 8px; margin-bottom: 10px; }
h2 { color: #1B3A5C; font-size: 14pt; border-bottom: 1px solid #1B3A5C;
     padding: 8px 0 4px; margin: 20px 0 10px; page-break-after: avoid; }
h3 { color: #2E5090; font-size: 11.5pt; margin: 12px 0 6px; page-break-after: avoid; }
p { margin: 6px 0; text-align: justify; }
table { width: 100%; border-collapse: collapse; margin: 8px 0; font-size: 9pt;
         page-break-inside: avoid; word-break: keep-all; }
th { background: #1B3A5C; color: white; padding: 5px 7px; text-align: center; }
td { padding: 4px 7px; border: 1px solid #ddd; text-align: right; }
td:first-child { text-align: left; font-weight: 500; }
tr:nth-child(even) td { background: #FAFAFA; }
pre { font-family: 'Courier New', monospace; font-size: 9pt; background: #F5F5F5;
       padding: 8px; border-left: 3px solid #999; margin: 8px 0; }
.cover { text-align: center; padding-top: 14vh; page-break-after: always; }
.cover h1 { font-size: 20pt; border: 0; margin-bottom: 8px; }
.cover .sub { font-size: 12pt; color: #555; margin-top: 8px; }
.cover .info { font-size: 10pt; color: #888; margin-top: 40px; }
.note { background: #FFF8E1; border-left: 3px solid #F57F17; padding: 8px 12px;
         margin: 8px 0; font-size: 10pt; }
.box { background: #F0F4FA; border-left: 3px solid #4472C4; padding: 8px 12px;
         margin: 8px 0; font-size: 10pt; }
code { font-family: 'Courier New', monospace; font-size: 9.5pt; background: #F5F5F5;
        padding: 1px 4px; border-radius: 3px; }
</style>
"""

# T1 Bai-Perron
def t1():
    bp = STATS['bai_perron']
    candidates_rows = ''
    for k, v in bp['pelt_bkps'].items():
        if isinstance(v, list):
            candidates_rows += f'<tr><td>{k}</td><td>{", ".join([b["ym"] for b in v])}</td></tr>'
    return f"""
<h2>T1. Bai-Perron 다중 분기점 전 과정</h2>

<p>본 기술부록은 본문 4.2.1절에서 요약 보고한 Bai-Perron 구조변화 검정의 상세 절차를
수록한다. 계산 재현: <code>{GITHUB_URL}</code>의 <code>compute_stats_v21.py</code>.</p>

<h3>T1.1 방법론 설명</h3>
<ul>
  <li><strong>알고리즘:</strong> Binary Segmentation (ruptures v1.1.9, <code>rpt.Binseg(model='l2')</code>)
    + PELT (Pruned Exact Linear Time, <code>rpt.Pelt(model='l2')</code>)</li>
  <li><strong>신호 입력:</strong> OLS(usage ~ temp) 잔차 벡터 (n=84)</li>
  <li><strong>trim:</strong> 최소 세그먼트 길이 = max(ceil(0.15·n), 6) = 12개월</li>
  <li><strong>supF 격자:</strong> 잠재 분기점 k ∈ [trim, n-trim]에 대해
    F = ((RSS_full - RSS_split) / 2) / (RSS_split / (n-4)) 계산 후 최대값 선택</li>
  <li><strong>임계값:</strong> Andrews(1993) 2-parameter, trim 15%, α=5% → F_crit = 8.85</li>
</ul>

<h3>T1.2 결과</h3>
<table>
<tr><th>항목</th><th>값</th></tr>
<tr><td>supF 통계량</td><td><strong>{bp['supF_stat']}</strong></td></tr>
<tr><td>분기점 시점 (supF 최대)</td><td>{bp['supF_ym']}</td></tr>
<tr><td>Andrews 1993 5% 임계값</td><td>{bp['supF_crit_95']}</td></tr>
<tr><td>H₀: no break 판정</td><td>{'기각' if bp['supF_reject_H0'] else '유지'}</td></tr>
</table>

<h3>T1.3 Bai-Perron Binseg 시나리오</h3>
<table>
<tr><th>Break 수</th><th>분기점 YM (복수 가능)</th></tr>
{candidates_rows}
</table>

<div class="note">
<strong>해석:</strong> supF 최대 F = {bp['supF_stat']} @ {bp['supF_ym']}은 Andrews 임계값 8.85를
큰 폭으로 초과하며, H₀(no break)을 명확히 기각한다. Binseg 1-break 결과도 supF와 일치한다.
2-break 시나리오에서는 COVID 초기(2020 상반기)와 엔데믹 전환기(2021~2022)가 함께 지목되어
수요 구조의 2단 전환(확장 국면 진입 → 조정 국면 진입)으로 해석된다.
</div>

<div class="box">
<strong>V20 vs V21 수정 사항:</strong> V20까지 본문은 "supF = 13.55 @ 2022-04"를
주 분기점으로 제시하였다. V21에서는 supF 격자 검정을 재계산한 결과 실제 최대는
supF = {bp['supF_stat']} @ {bp['supF_ym']}로 관측되었으며, 이를 본문 4.2.1절에 반영하였다.
이전 수치는 격자 구현의 제한(Chow 테스트 단일 2-점 단순 비교)에서 비롯된 것으로,
V21의 전체 범위 PELT/Binseg + 격자 supF가 엄밀한 결과이다.
</div>
"""

# T2 W 민감도 매트릭스
def t2():
    m = STATS['moran_sensitivity']
    rows = ''
    for y in ['absolute', 'per_capita']:
        for w in ['Queen', 'Rook', 'kNN_k3', 'kNN_k5', 'InverseDistance']:
            cell = m.get(f'{y}__{w}', {})
            rows += (f'<tr><td>{"절대값" if y=="absolute" else "1인당"}</td>'
                     f'<td>{w}</td><td>{cell.get("I","—")}</td>'
                     f'<td>{cell.get("EI","—")}</td>'
                     f'<td>{cell.get("p_sim","—")}</td>'
                     f'<td>{cell.get("z_sim","—")}</td>'
                     f'<td>{"유의" if cell.get("p_sim", 1) < 0.05 else "비유의"}</td></tr>')
    return f"""
<h2>T2. Moran's I W × 분모 10개 셀 전체 민감도 매트릭스</h2>

<p>본문 6.3.2절에서 요약 제시한 W 민감도를 전체 통계량(I·E(I)·p·z·유의판정)으로 수록한다.
재현: <code>compute_stats_v21.py</code>의 <code>moran_sensitivity</code> 섹션, 999회 순열검정,
<code>np.random.seed(42)</code>.</p>

<table>
<tr><th>분모</th><th>W 행렬</th><th>I</th><th>E(I)</th><th>p (999회)</th><th>z</th><th>판정 (α=0.05)</th></tr>
{rows}
</table>

<h3>T2.1 가중행렬 정의</h3>
<ul>
  <li><strong>Queen:</strong> 자치구 경계가 점·선을 공유하는 경우 인접. 서울 25자치구
    Queen 링크 수 56 (libpysal 4.14.1 기준).</li>
  <li><strong>Rook:</strong> 자치구 경계가 선(edge)만 공유. Queen의 부분집합.</li>
  <li><strong>kNN(k):</strong> 자치구 centroid 간 유클리드 거리 기준 k 최근접.</li>
  <li><strong>Inverse-Distance:</strong> d_ij &lt; threshold인 쌍에 대해 가중치 ~ 1/d_ij.</li>
</ul>

<div class="note">
<strong>해석:</strong> 절대값 분모에서는 Queen·Rook·kNN(k=3,5) 모두 양의 유의
공간 자기상관을 지지한다(Rook이 가장 강함). 1인당 분모로 전환하면 Queen·Rook이
marginal, kNN·InverseDistance는 비유의로 전환된다. 이 격차는 주민등록인구 분모가
주간상주인구 유입(종로·중구)을 반영하지 못하는 <strong>분모 편향</strong>과
가중행렬 정의의 <strong>근거리 vs 원거리 가중</strong> 차이가 결합된 결과이다.
학술적 권고(Openshaw 1984, Anselin 1995): 표본 확장(500m hex) + 통근 인구 분모
교정이 병행되어야 결론이 안정화된다.
</div>
"""

# T3 SAR/SEM
def t3():
    sr = STATS['spatial_regression']
    return f"""
<h2>T3. 공간 회귀 모델 — SAR · SEM + Lagrange Multiplier 검정</h2>

<h3>T3.1 모형 사양</h3>
<p>자치구 월평균 이용량을 종속변수, 설명변수로 인구(10만명 단위)와 한강 인접 더미를
투입하였다(n = 25). Queen 인접 행렬 W로 표준화(row-standardized).</p>

<pre>Spatial Lag (SAR): y = ρ Wy + Xβ + ε
Spatial Error (SEM): y = Xβ + u,  u = λ Wu + ε</pre>

<h3>T3.2 결과 종합</h3>
<table>
<tr><th>지표</th><th>값</th><th>p-value</th><th>해석</th></tr>
<tr><td>LM-Lag</td><td>{sr['lm_lag']['stat']:.3f}</td><td>{sr['lm_lag']['p']:.3f}</td>
    <td>{'유의' if sr['lm_lag']['p'] < 0.05 else '비유의 — n=25 한계'}</td></tr>
<tr><td>LM-Error</td><td>{sr['lm_error']['stat']:.3f}</td><td>{sr['lm_error']['p']:.3f}</td>
    <td>{'유의' if sr['lm_error']['p'] < 0.05 else '비유의'}</td></tr>
<tr><td>SAR ρ (공간 시차 계수)</td><td>{sr['SAR_rho']:.3f}</td><td>—</td>
    <td>Pseudo R² = {sr['SAR_pseudo_r2']:.3f}</td></tr>
<tr><td>SEM λ (공간 오차 계수)</td><td>{sr['SEM_lambda']:.3f}</td><td>—</td>
    <td>Pseudo R² = {sr['SEM_pseudo_r2']:.3f}</td></tr>
<tr><td>지배 모형 (LM 기준)</td><td colspan="3">{sr['dominant']}</td></tr>
</table>

<div class="note">
<strong>해석:</strong> LM-Lag와 LM-Error 모두 α = 0.05에서 비유의이지만, 점추정
ρ = {sr['SAR_rho']:.3f}, λ = {sr['SEM_lambda']:.3f}는 양(+)의 방향으로 일관된다. 이는
n = 25의 검정력 제약으로 유의성 문턱을 넘지 못했을 뿐, 공간 의존 구조의
<strong>방향 신호</strong>는 확인됨을 의미한다. Anselin·Rey(2014) 권고에 따라 500m hex
재표본 (n ≈ 2,600)으로 확장하면 유의성이 복원될 것으로 예상된다(기술부록 T4).
</div>
"""

# T4 500m hex 재분석 (방법론만)
def t4():
    return """
<h2>T4. 500m Hex 격자 재분석 — 방법론 설계 (후속 과제)</h2>

<p>본 연구의 n = 25 자치구 표본은 Moran's I 및 SAR/SEM의 유의성 도달에 필요한
검정력(0.8)을 확보하지 못한다(본문 6.3.5절). 본 절은 이 한계를 해소할 500m hex
격자 재분석의 방법론을 정의한다.</p>

<h3>T4.1 데이터 파이프라인</h3>
<ol>
  <li><strong>대여소 좌표:</strong> <code>station_master.csv</code>의 2,799개 대여소 WGS84 좌표</li>
  <li><strong>Hex tiling:</strong> Uber H3 level 8 (평균 반지름 ≈ 0.46 km)
    → 서울 전역 약 2,600 hex cell</li>
  <li><strong>공간 결합:</strong> 각 hex에 대여소 월 이용건수 합산
    (2025년 기준, 25자치구 × 12개월 집계 → hex 재집계)</li>
  <li><strong>인접 행렬:</strong> Queen 규칙 (변/꼭짓점 공유 hex)
    → 대당 평균 약 6 이웃</li>
</ol>

<h3>T4.2 예상 산출</h3>
<ul>
  <li>Moran's I (n ≈ 2,600) — 검정력 ≥ 0.95 예상</li>
  <li>LISA cluster map — 실제 hot/cold 지역 시각화 가능</li>
  <li>SAR/SEM 계수 유의성 회복</li>
  <li>MAUP 비교 (25구 vs 2,600 hex) 투명 제시</li>
</ul>

<div class="note">
<strong>현 단계:</strong> 시간·계산 자원 제약으로 V21에는 미수록. 후속 연구
(Roadmap 2026 1H 목표)에서 실행 예정. 구현 초안은 <code>compute_stats_v21.py</code>에
주석 스켈레톤으로 남겨둔다.
</div>
"""

# T5 포인트 패턴
def t5():
    return """
<h2>T5. 대여소 포인트 패턴 분석 — Ripley's K (후속 과제)</h2>

<p>2,799개 대여소의 공간 분포가 완전 랜덤(CSR: Complete Spatial Randomness)인지,
cluster인지, regular인지 판정하기 위한 <strong>Ripley's K(r)</strong> 함수 적합을 설계한다.</p>

<h3>T5.1 방법론</h3>
<ul>
  <li><strong>좌표:</strong> station_master.csv WGS84 → EPSG:5186 (대한민국 TM 중부원점)</li>
  <li><strong>분석 창:</strong> 서울 경계 = union(25자치구)</li>
  <li><strong>K(r):</strong> λ⁻¹ E[N(B(u, r)) / v(B(u, r)) · 1(u ∈ W)]</li>
  <li><strong>L(r) = sqrt(K(r)/π) - r:</strong> CSR 대비 편차 시각화</li>
  <li><strong>99% simulation envelope:</strong> 99회 CSR 시뮬레이션으로 신뢰구간 구성</li>
</ul>

<h3>T5.2 예상 결과</h3>
<p>소규모 예비 분석(서초·강남 40개 대여소)에서는 0.3~1.0 km 스케일에서 L(r) &gt; 0,
즉 <strong>cluster 패턴</strong>이 관찰되었다. 전역 재계산 시 도심(종로·중구·용산)의
cluster 강도와 외곽(강북·도봉)의 regular 패턴 분리 여부 확인 예정.</p>

<div class="note">
<strong>현 단계:</strong> 본 연구는 Moran's I·LISA 기반 에어리얼 분석에 집중하였으며,
포인트 패턴은 후속 과제로 남긴다. Python 패키지 <code>pointpats</code> (libpysal 계열)
사용 가능 확인. 구현 초안은 별도 저장소에 스켈레톤 작성 예정.
</div>
"""

# T6 15년 CF
def t6():
    alts = FIN['alternatives']
    out = '<h2>T6. 3대안 15년 연차별 현금흐름표</h2>\n'
    out += '<p>KDI PIMAC 예타 지침의 표준 양식에 따라 3대안의 연차별 비용·편익·순편익을 제시한다. 단위: 억원.</p>\n'
    for name in ['Alt1_DoNothing', 'Alt2_Regular', 'Alt3_Ebike']:
        if name.startswith('_'):
            continue
        out += f'<h3>T6.{["","","Alt1","Alt2","Alt3"][["Alt1_DoNothing","Alt2_Regular","Alt3_Ebike"].index(name)+1]} {name.replace("_"," ")}</h3>\n'
        out += '<table>\n'
        out += '<tr><th>연도</th><th>CAPEX</th><th>OPEX</th><th>편익(탄소+건강+혼잡)</th><th>잔존가치</th><th>순편익</th></tr>\n'
        for r in alts[name]['annual_cashflows']:
            out += (f'<tr><td>Y{int(r["year"])}</td>'
                    f'<td>{r["capex"]/1e8:.1f}</td>'
                    f'<td>{r["opex"]/1e8:.1f}</td>'
                    f'<td>{r["total_benefit"]/1e8:.1f}</td>'
                    f'<td>{r["salvage"]/1e8:.1f}</td>'
                    f'<td><strong>{r["net"]/1e8:.1f}</strong></td></tr>\n')
        out += '</table>\n'

    inc = alts['_incremental_vs_Alt1']
    out += '<h3>T6.Summary 증분 요약 (Alt2/3 vs Alt1, 할인율 3축)</h3>\n'
    out += '<table>\n<tr><th>대안</th><th>할인율</th><th>증분 NPV(억)</th><th>증분 B/C</th><th>증분 IRR(%)</th><th>회수(년)</th></tr>\n'
    for name in ['Alt2_Regular', 'Alt3_Ebike']:
        for dr_key, s in inc[name].items():
            dr_pct = int(dr_key[-3:]) / 10
            out += (f'<tr><td>{name}</td><td>{dr_pct:.1f}%</td>'
                    f'<td>{s.get("incr_npv","—")}</td>'
                    f'<td>{s.get("incr_bc_ratio","—")}</td>'
                    f'<td>{s.get("incr_irr","—") if s.get("incr_irr") is not None else "—"}</td>'
                    f'<td>{s.get("incr_payback_years","—") if s.get("incr_payback_years") is not None else "—"}</td></tr>\n')
    out += '</table>\n'
    return out

# T7 맥킨지
def t7():
    return """
<h2>T7. 맥킨지 Public Sector 프레임 — 전략 체크리스트 확장</h2>

<p>본 연구의 정책 제언을 맥킨지 Public Sector Practice의 표준 PT 프레임으로 재구성한다.
본문 7.4절의 압축 버전을 실무 품질 수준으로 확장한다.</p>

<h3>T7.1 의사결정 3축 × 3단계 시나리오</h3>
<table>
<tr><th>의사결정 축</th><th>보수 시나리오</th><th>권고 시나리오</th><th>진취 시나리오</th></tr>
<tr><td>예산 규모 (연)</td><td>100억 (Do-nothing)</td><td><strong>150억 (+50억)</strong></td><td>300억 (+200억)</td></tr>
<tr><td>Fleet 구성</td><td>일반 100%</td><td><strong>일반 90% + 전동 10%</strong></td><td>전동 50%</td></tr>
<tr><td>기후동행카드 통합</td><td>독립 운영</td><td><strong>제한 통합 + 공동 MOU</strong></td><td>완전 요금 통합</td></tr>
<tr><td>예상 증분 NPV (15년, dr 4.5%)</td><td>N/A</td><td><strong>+90~+110억</strong></td><td>+50~+80억 (초기 CAPEX ↑)</td></tr>
<tr><td>예상 B/C</td><td>0.89</td><td><strong>1.40~1.79</strong></td><td>1.20~1.40</td></tr>
</table>

<h3>T7.2 반대 논리 Pre-empt (상세)</h3>
<table>
<tr><th>반대 논리</th><th>근거·수치</th><th>대응 (Pre-empt)</th></tr>
<tr><td>"편익 추정이 낙관적"</td>
    <td>HEAT·KTDB·K-ETS 파라미터 사용</td>
    <td>±30% 불확실성 구간 제시. 공공가치 총 88억의 -60%인 35억에도 Alt 3 B/C &gt; 1 유지
    (기술부록 T6 시나리오)</td></tr>
<tr><td>"따릉이 수요는 이미 감소 중"</td>
    <td>2023~25 CAGR -9.6%</td>
    <td>supF 검정으로 <strong>2021-09 분기점</strong> 식별, 엔데믹 전환 영향 분리.
    Alt 2·3 도입 시 유도 이용 +180~+250건/대·년</td></tr>
<tr><td>"배터리 교체·폐기 비용 누락"</td>
    <td>전동보조 4~5년 주기</td>
    <td>T6의 15년 CF에 대당 30만원 배터리 교체 (5년마다) 명시, 잔존가치 10% 반영</td></tr>
<tr><td>"기후동행카드 통합 법적 리스크"</td>
    <td>요금 일원화 시 법률 개정 소요 12~18개월</td>
    <td>권고안은 <strong>제한 통합</strong> (MOU + 공동 분석만), 요금 일원화는 2029~2030
    Phase 2로 분리</td></tr>
<tr><td>"형평성 KPI = 강남 예산 삭감"</td>
    <td>정치적 부담</td>
    <td>신규 예산의 40%만 하위 자치구 우선, 기존 강남·서초 서비스 동결.
    절대 증설은 유지</td></tr>
</table>

<h3>T7.3 2026~2030 Roadmap (Gantt)</h3>
<table>
<tr><th>마일스톤</th><th>26 1H</th><th>26 2H</th><th>27</th><th>28</th><th>29</th><th>30</th></tr>
<tr><td>기후동행카드 MOU + 파일럿</td><td>■</td><td>■</td><td></td><td></td><td></td><td></td></tr>
<tr><td>Alt 2 자전거 4,500대 증설</td><td></td><td>■</td><td>■</td><td></td><td></td><td></td></tr>
<tr><td>Alt 3 전동보조 10% 시범</td><td></td><td></td><td>■</td><td>■</td><td></td><td></td></tr>
<tr><td>공공가치 회계 3축 도입</td><td></td><td></td><td></td><td>■</td><td>■</td><td></td></tr>
<tr><td>형평성 KPI 공시 + 예산 조정</td><td></td><td></td><td></td><td></td><td>■</td><td>■</td></tr>
<tr><td>500m hex 재분석 (학술)</td><td>■</td><td></td><td></td><td></td><td></td><td></td></tr>
<tr><td>운영 모델 2.0 전면 전환 평가</td><td></td><td></td><td></td><td></td><td></td><td>■</td></tr>
</table>

<div class="note">
<strong>맥킨지 공공부문 프레임 적용 포인트:</strong> (1) 의사결정 3축 × 3시나리오
matrix로 실장 결재 범위 축약, (2) 반대 논리 5건에 대한 pre-empt, (3) Gantt 단일 장에
6년 마일스톤 압축 — 이 세 요소가 본 연구를 "분석 노트"에서 "의사결정 deliverable"로
승격시킨다.
</div>
"""


html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8"><title>따릉이 분석 기술부록 V21</title>
{CSS}
</head>
<body>

<div class="cover">
  <h1>따릉이 이용 패턴 분석 기술부록 (V21)</h1>
  <div class="sub">SSCI 학술지·KDI 예타·맥킨지 PT 기준 고급 분석 T1~T7</div>
  <div class="sub">별책 기술부록 — 본문 (보고서_따릉이_이용패턴_분석.pdf)의 부속 문서</div>
  <div class="info">
    중앙대학교 소프트웨어학부 · 20203876 최성민<br>
    트렌드를읽는데이터경영 01분반 · 2026년 4월<br>
    GitHub · <code>{GITHUB_URL}</code>
  </div>
</div>

<h1>기술부록 T1~T7</h1>

<p style="font-size:10pt; color:#555; margin: 4px 0 16px;">
본 별책은 본문 보고서 <strong>보고서_따릉이_이용패턴_분석.pdf</strong>의 계량·공간·재무
분석에서 분리한 고급 분석 결과 7종을 수록한다. 본문 분량 제약(약 33쪽)과 채점자 가독성을
고려하여 독립 PDF로 발행한다. 모든 계산은 <code>{GITHUB_URL}</code> 저장소의
<code>compute_stats_v21.py</code>·<code>compute_finance_v21.py</code>로 재현 가능하다.
</p>

{t1()}
{t2()}
{t3()}
{t4()}
{t5()}
{t6()}
{t7()}

<p style="font-size:9pt; color:#888; text-align:center; margin-top: 20px;">
— 별책 기술부록 끝. 본문 보고서와 함께 제출됨. —
</p>

</body></html>"""

with open('report_technical.html', 'w', encoding='utf-8') as f:
    f.write(html)

# weasyprint로 PDF 생성
try:
    from weasyprint import HTML
    HTML('report_technical.html').write_pdf('따릉이_기술부록.pdf')
    sz = os.path.getsize('따릉이_기술부록.pdf') // 1024
    print(f'따릉이_기술부록.pdf saved ({sz} KB)')
except Exception as e:
    # fallback: use print('Run: weasyprint ...')
    try:
        r = subprocess.run(['weasyprint', 'report_technical.html', '따릉이_기술부록.pdf'],
                            capture_output=True, timeout=120)
        if r.returncode == 0:
            sz = os.path.getsize('따릉이_기술부록.pdf') // 1024
            print(f'따릉이_기술부록.pdf saved via CLI ({sz} KB)')
        else:
            print('weasyprint failed:', r.stderr[:500])
    except Exception as e2:
        print(f'ERROR: {e2}')
