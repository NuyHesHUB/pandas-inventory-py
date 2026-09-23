# pandas-inventory-py

재고관리대장 엑셀에서 기간 판매량을 집계하고 FIFO로 입고 lot을 역추적해
원가 결과물 엑셀을 생성하는 도구. tkinter 데스크톱 앱(Windows/macOS)으로 배포한다.

## 필수 문서 (작업 전에 읽을 것)

- **[docs/LESSONS.md](docs/LESSONS.md)** — 패키징/tkinter/FIFO에서 실제로 겪은 함정 모음.
  특히 macOS onefile 금지, uv 독립형 Python의 Tcl/Tk 문제는 반복 리스크.
- **[docs/LOGIC-CALIBRATION.md](docs/LOGIC-CALIBRATION.md)** — FIFO 로직 교정 기록.
  사람이 만든 실제 결과물과 대조해 지속 교정하는 운영 문서. **FIFO 로직을 고칠 때는
  반드시 이 문서의 절차를 따르고(사례 기록 → 원인 분류 → 회귀 테스트 → 수정 이력),
  수정 후 문서를 갱신할 것.**

## 명령어

```bash
uv run python -m unittest discover -s tests   # 테스트 (필수: 커밋 전 통과)
uv run python -m sw_exel_py_project --ui      # UI 실행
uv run cost-report-desktop                    # 데스크톱 진입점 (창 유지 + 진행바)
```

배포 빌드는 GitHub Actions "Build desktop apps" 워크플로 (v* 태그 또는 수동 실행).
로컬 PyInstaller 빌드는 README의 Desktop build 절 참조 — **uv Python이 아닌
python.org 빌드로만** 할 것.

## 구조 요약

- `exel_reader.py` — 엑셀 로딩(lru_cache), 판매 집계, FIFO 코어. 순서 민감 정렬은
  전부 `kind="stable"` 유지할 것 (같은 날짜 lot은 원장 기재 순서가 FIFO 순서).
- `cost_report.py` — 결과물 생성 (템플릿 갱신 / 신규 생성). anchor 휴리스틱 주의.
- `app.py` — UI 결과 실행 + 요약 문구 (CLI/데스크톱 공용).
- `ui/main_ui.py` — tkinter UI (claude.ai/design 'v2' 디자인 기반). UI는 pandas를
  임포트하지 않는다 — 기동 속도를 위한 의도적 분리이므로 유지할 것.
- `desktop.py` — 데스크톱 진입점. pandas는 실행 시점 지연 임포트 (기동 속도).

## 주의사항

- 결과물 수치에 영향 주는 변경은 사용자와 대조 검증 없이 배포하지 말 것.
- 커밋 메시지는 한국어, conventional prefix (feat/fix/perf/ci/chore).
- `fifo_calculator.py`, `utils/sort_util.py` 등 죽은 코드가 남아 있음 (삭제 예정 목록은
  대화 이력 참조) — 새 코드에서 참조하지 말 것.
