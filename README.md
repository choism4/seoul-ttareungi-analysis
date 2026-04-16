# 서울 따릉이 이용 패턴 분석 (2019–2025)

중앙대학교 트렌드를읽는데이터경영 중간과제. 서울시 공공자전거 84개월 데이터에 대한 계량·공간·재무 통합 분석.

## 핵심 결과 (한 줄 요약)

- **구조변화 분기점**: supF = 54.6 @ 2021-09 (Andrews 1993 임계값 8.85 초과)
- **기온 효과 (SARIMAX exog)**: +82,671 건/°C (SARIMAX (1,1,2)(0,1,1,12) AIC=1,589)
- **공간 자기상관**: Queen I=+0.279 (p=0.015), LISA FDR 보정 후 0 자치구 유의
- **증분 NPV(Alt3 전동보조 vs Do-nothing, 할인율 4.5%)**: **+92억** (B/C 1.40, IRR 10.9%, 회수 10년)

## 재현 절차 (3단계)

```bash
# 1. 라이브러리 설치
pip install -r requirements.txt

# 2. 통계·재무 계산
python compute_stats_v21.py      # → v21_stats.json
python compute_finance_v21.py    # → v21_finance.json

# 3. 보고서·부록 생성
python create_report_v2.py                 # → 보고서_따릉이_이용패턴_분석.pdf
python create_technical_appendix.py        # → 따릉이_기술부록.pdf
```

또는 `make reproduce` 한 번에 실행 가능.

## 저장소 구조

```
midterm_data_ttareungi/
├── README.md                    이 파일
├── requirements.txt             Python 의존성
├── Makefile                     파이프라인 타깃
├── seed_config.json             모든 난수 시드
├── CITATION.cff                 인용 정보
├── data_manifest.csv            원자료 SHA-256 지문
├── compute_stats_v21.py         통계 10종 계산
├── compute_finance_v21.py       3대안 15년 재무 계산
├── create_report_v2.py          본문 보고서 생성
├── create_technical_appendix.py 기술부록 PDF 생성
├── monthly_aggregate.json       이용량 집계 자료
├── district_monthly.json        자치구×월 자료
├── seoul_districts.geojson      25자치구 경계
├── districts_adjacency.json     Queen 인접행렬
└── 보고서_따릉이_이용패턴_분석.pdf  본문 (제출본)
```

## 데이터 출처

| 자료 | 출처 | 접근일 |
|---|---|---|
| 따릉이 이용정보 (월별) | 서울 열린데이터광장 OA-15248 | 2026-04-14 |
| 따릉이 대여소 정보 | 서울 열린데이터광장 OA-13252 | 2026-04-15 |
| 서울 월평균기온 | Open-Meteo Historical API (37.57°N) | 2026-04-14 |
| 25자치구 경계 | southkorea/seoul-maps (MIT License) | 2026-04-16 |
| 주민등록인구 (자치구) | 서울시 주민등록인구통계 2025.01 | - |

## 분석 방법 (요약)

1. **계량**: 단순 OLS + 연도 stratified OLS + SARIMAX 5 후보 중 AIC 최소 선정
2. **구조변화**: Bai-Perron PELT/Binseg (ruptures) + supF 격자 검정 (trim 15%, Andrews 1993 임계값)
3. **공간**: Queen/Rook/kNN(k=3,5)/Inverse-Distance 5종 × 절대값·1인당 2종 = 10개 Moran's I 민감도
4. **LISA**: Local Moran 999회 순열 + Benjamini–Hochberg FDR 보정
5. **공간회귀**: OLS → LM-Lag/LM-Error → ML_Lag(SAR)/ML_Error(SEM)
6. **검정력**: n=25에서 MDE 시뮬레이션 (100회 반복, alpha=0.05)
7. **재무**: 3대안 15년 LCC + 할인율 3축 (3.0/4.5/5.5%) + 편익 ±20/40/60% 민감도

## 라이선스

- **데이터**: 공공누리 제1유형 (서울 열린데이터광장)
- **코드·보고서**: MIT / CC-BY 4.0

## 인용

본 보고서 또는 코드를 인용할 경우 `CITATION.cff`를 참고.
