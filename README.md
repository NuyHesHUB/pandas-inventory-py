# pandas-inventory-py

Excel inventory automation utilities.

## Setup

Install Python and project dependencies with uv:

```powershell
uv python install
uv sync
```

## Run

```powershell
uv run python -m sw_exel_py_project --ui
```

## Test

```powershell
uv run python -m unittest
```

## Desktop build (Windows / macOS)

배포용 실행파일은 PyInstaller로 만든다. 크로스 컴파일이 안 되므로 각 OS에서 빌드해야 한다.

주의: uv가 내려받는 독립형 Python은 Tcl/Tk 라이브러리를 PyInstaller가 제대로 담지 못하므로,
로컬 빌드는 python.org 공식 설치본(3.13+)으로 한다.

로컬 빌드 (해당 OS에서, python.org Python 기준):

```powershell
pip install . pyinstaller
# Windows
pyinstaller --noconfirm --clean --onefile --windowed --name CostReport --icon assets/icon.ico desktop_launcher.py
# macOS
pyinstaller --noconfirm --clean --onefile --windowed --name CostReport --icon assets/icon.icns desktop_launcher.py
# 결과물: dist/CostReport.exe (Windows) / dist/CostReport.app (macOS)
```

앱 아이콘은 `assets/sungwoo_symbol.svg`가 원본이다. 아이콘을 바꾸려면 SVG를 교체한 뒤
cairosvg로 512px PNG를 만들고 Pillow로 `assets/icon.ico`(Windows), `assets/icon.icns`(macOS),
`src/sw_exel_py_project/ui/app_icon.py`(타이틀바용 64px base64)를 다시 생성한다.

GitHub Actions 자동 빌드: `v*` 태그를 푸시하거나 Actions 탭에서 `Build desktop apps` 워크플로를
수동 실행하면 Windows/macOS 실행파일이 아티팩트로 올라온다.

데스크톱 실행파일은 콘솔 없이 동작하며, 실행 결과 요약을 메시지박스로 보여준다.
개발 환경에서 같은 흐름을 확인하려면:

```powershell
uv run cost-report-desktop
```
