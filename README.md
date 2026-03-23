<!-- poetry run python -m src.sw_exel_py_project --ui -->

```
Poetry 설치(한 번만):
pip install poetry

의존성 설치:
poetry install

프로그램 실행:
poetry run python -m src.sw_exel_py_project --ui
```
```
excel-automator/
├── app/
│   ├── __init__.py
│   ├── main.py                # entrypoint: python -m app
│   ├── watcher.py             # watchdog 폴더 감시/이벤트 큐
│   ├── pipeline.py            # ingest → validate → transform → export
│   ├── processors/
│   │   ├── __init__.py
│   │   ├── rules.py           # 컬럼 매핑, 수식, 파생 컬럼 규칙
│   │   └── transforms.py      # 구체 변환 로직(집계, 피벗 등)
│   ├── io_utils.py            # 엑셀 읽기/쓰기, 파일 이동
│   ├── validators.py          # 형식/스키마/데이터 검증
│   ├── report.py              # 처리 결과 요약 리포트(csv/md)
│   └── config.py              # 설정 로딩 / 경로 관리
│
├── config/
│   ├── settings.yaml          # 경로, 파일패턴, 출력방식, 언어/로케일
│   └── rules.yaml             # 컬럼 매핑, 검증 룰, 파생컬럼 정의
│
├── dropbox/                   # 사용자가 파일을 넣는 곳
├── output/                    # 결과물 저장
├── errors/                    # 실패 파일과 원인 로그
├── logs/                      # rotate 로그 파일
│
├── tests/
│   ├── test_pipeline.py
│   └── test_validators.py
│
├── scripts/
│   ├── run_watch.bat          # 윈도우용 더블클릭 실행
│   └── run_watch.command      # 맥용 더블클릭 실행(실행권한 부여)
│
├── pyproject.toml             # 의존성/빌드 설정(권장)
├── requirements.txt           # (poetry 미사용 시)
└── README.md
```
