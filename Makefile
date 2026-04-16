PY := python3

.PHONY: all stats finance report appendix reproduce clean

all: reproduce

stats:
	$(PY) compute_stats_v21.py

finance:
	$(PY) compute_finance_v21.py

report: stats finance
	$(PY) create_report_v2.py

appendix:
	$(PY) create_technical_appendix.py

reproduce: stats finance report appendix
	@echo "✓ Pipeline complete. Expected outputs:"
	@echo "  보고서_따릉이_이용패턴_분석.pdf (본문)"
	@echo "  따릉이_기술부록.pdf (별책)"

clean:
	rm -f v21_stats.json v21_finance.json report.html
	@echo "Regenerable artifacts cleaned."
