# 교훈 모음 (Lessons Learned)

이 프로젝트를 진행하며 실제로 겪은 문제와 해결책. 같은 함정을 다시 밟지 않기 위한 기록이다.
새 교훈이 생기면 "증상 → 원인 → 해결 → 교훈" 형식으로 추가한다.

## 배포 / 패키징

### macOS 앱 기동이 10초씩 걸렸다 (2026-09)
- **증상**: PyInstaller로 만든 .app이 맥에서 실행마다 ~10초 걸림.
- **원인**: `--onefile`은 매 실행마다 번들 전체(약 108MB)를 임시 폴더에 풀고,
  macOS가 새로 풀린 라이브러리 수십 개의 코드 서명을 **매번 재검증**한다
  (파일이 매번 새로 생기므로 검증 캐시가 무효).
- **해결**: macOS 빌드만 `--onedir`로 전환 (90c26f2). .app은 어차피 폴더라서
  배포 형태(zip 안의 .app 하나)와 사용자 경험은 동일하다.
- **교훈**: pandas급 대형 의존성을 가진 앱은 macOS에서 onefile을 쓰지 마라.
  Windows는 단일 exe 배포 편의상 onefile 유지가 낫다.

### uv 독립형 Python은 GUI/빌드에 쓰면 안 된다 (2026-09)
- **증상 1**: WSL에서 tkinter 한글이 전부 □(tofu). 폰트를 설치해도 그대로.
- **증상 2**: 그 Python으로 PyInstaller 빌드하면 `libtcl9tk9.0.so` 누락으로 즉사.
- **원인**: uv가 내려받는 python-build-standalone은 자체 번들 Tcl/Tk가
  fontconfig와 연결돼 있지 않고, PyInstaller 후크도 그 배치를 인식하지 못한다.
- **해결**: GUI 실행/PyInstaller 빌드는 **python.org 공식 빌드** 사용.
  CI는 `actions/setup-python`으로 강제 (워크플로 주석 참조).
- **WSL/Ubuntu 로컬 실행** (2026-10): Linux용 python.org 바이너리는 없으므로
  deadsnakes PPA의 `python3.13` + `python3.13-tk`를 설치하고
  `uv venv --clear --python /usr/bin/python3.13` 후 `uv sync`. **`-tk` 패키지를
  빠뜨리기 쉬움** (없으면 tkinter import 자체가 실패). `.venv`를 새로 만들면 uv가
  다시 독립형 3.14를 고르므로 위 명령으로 재생성할 것.
- **교훈**: "tkinter import 성공"과 "tkinter가 제대로 동작"은 별개다.
  폰트가 깨지면 폰트 탓하기 전에 Tk 빌드가 fontconfig를 쓰는지부터 확인
  (`tkfont.Font(family=...).actual()`이 'fixed'로 나오면 Tk 빌드 문제).

### WSL git은 자격 증명이 없어 push가 안 된다 (2026-10)
- **증상**: WSL 쉘에서 `git push` → `fatal: could not read Username for
  'https://github.com': No such device or address`. fetch/ls-remote는 되는데
  (공개 레포라 읽기는 익명 허용) push만 실패한다.
- **원인**: WSL git에 `credential.helper`가 설정돼 있지 않다. Windows 쪽 git에는
  자격 증명 관리자가 붙어 있어서, 같은 레포를 Windows 터미널에서 밀면 성공한다.
  그래서 "분명 아까는 됐는데" 하고 헷갈리기 쉽다.
- **해결 1 (즉시)**: 레포가 `C:` 드라이브에 있으므로 Windows git을 그대로 쓴다.
  ```bash
  "/mnt/c/Program Files/Git/bin/git.exe" -C 'C:\path\to\repo' push origin <ref>
  ```
- **해결 2 (영구)**: WSL git을 Windows 자격 증명 관리자에 연결.
  ```bash
  git config --global credential.helper \
    "/mnt/c/Program\ Files/Git/mingw64/bin/git-credential-manager.exe"
  ```
- **교훈**: push가 실패하면 네트워크·권한·브랜치 보호를 의심하기 전에 **어느 git
  바이너리로 실행했는지**부터 본다. 에러 문구가 "Username을 못 읽는다"면 인증 설정
  문제이지 서버 거부가 아니다.

### 기타 패키징 함정
- PyInstaller는 크로스 컴파일 불가 → OS별 CI 매트릭스로 빌드 (GitHub Actions).
- `upload-artifact`는 실행 권한을 보존하지 않음 → macOS .app은 zip으로 묶어 업로드.
- 미서명 배포물: Windows SmartScreen은 Mark-of-the-Web 때문 → 사내 공유폴더/USB
  배포면 경고 없음. macOS는 `xattr -cr 앱.app` 1회. 근본 해결은 코드 서명(유료).
- pandas 임포트(수백 ms)는 진입점에서 지연 임포트해 창이 먼저 뜨게 한다 (desktop.py).

## tkinter UI

- **macOS 다크 모드**: Tk 기본 글자색이 흰색이 된다. 배경색을 밝게 강제한 위젯은
  `fg`(와 Entry류는 `insertbackground`)를 반드시 명시할 것 (94ca19a).
- **타이틀바 X 버튼**: X로 닫으면 root가 이미 파괴된 상태라 `root.destroy()`가
  TclError를 던진다. mainloop 후 destroy는 try/except로 감쌀 것.
- **드롭다운/팝업**: 창 높이보다 아래로 열리면 잘린다. 열 때 공간을 재서
  위/아래 자동 전환 (main_ui.position_popup).
- **디자인 시안의 타이틀바 목업**(Windows 크롬/맥 신호등)은 재현하지 않는다 —
  실제 앱은 OS 네이티브 타이틀바가 정답.

## FIFO / 판매 원가 로직

원인 분석과 수정 이력은 [LOGIC-CALIBRATION.md](LOGIC-CALIBRATION.md)에 상세 기록.
핵심 교훈만 요약:

- **pandas 정렬은 기본이 불안정 정렬**이다. 같은 날짜 행이 많아지면(동점 16개 이상)
  원장 기재 순서가 조용히 깨진다. FIFO처럼 순서가 의미인 곳은 전부 `kind="stable"`.
- **"마지막 행" 로직은 빈 값을 만난다**. 재고 스냅샷이 마지막 행 하나만 보면
  그 행의 현재고가 비어 있을 때 0으로 오판한다. 마지막 "유효" 값을 찾아라.
- **휴리스틱 anchor는 우연 일치에 당한다**. "수량이 같으면 사람이 확정한 것"이라는
  가정은 매달 같은 양을 사가는 고정 거래처에서 깨졌다. 휴리스틱에는 이차 증거
  (수기 흔적: 비표준 수식/텍스트 단가)를 요구할 것.
- 합계가 맞아도 lot 분할이 틀릴 수 있다. 보정 seed가 데이터 공백을 조용히 메우므로
  "합계 일치"만으로 검증 완료라고 판단하지 말 것.
- **판정은 거래의 증거로 하고, 거래처의 속성으로 하지 마라.** "공급처 상호에 한글이
  있으면 국내 매입"은 상호가 한글인 수입처("화동", `Hefei Top Industry(무역회사)`)를
  통째로 떨어뜨렸고, GLYCERIN 계열은 판매가 잡히는데 결과물에 0행으로 나왔다
  (2026-10-27, aa0877f). 국내 매입 여부는 **입고(수입)단가가 원화냐(₩500 이상)** 로만
  판별한다. 원장 실측상 ₩ 서식 행은 전부 500 이상, 달러·엔화 행은 전부 500 미만이라
  이 기준은 값과 서식이 정확히 일치한다.
- **품목이 결과물에서 통째로 빠지면 판매가 아니라 입고를 의심하라.** 판매 집계는
  멀쩡한데 입고 lot이 0건이면 `[경고] ... cutoff 이전 입고 없음`만 찍히고 품목이
  조용히 사라진다. 제외 휴리스틱이 가장 흔한 원인이다.
