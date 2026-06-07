---
name: package-exe
description: "안전점검 모니터링 시스템을 PyInstaller로 단일 EXE 파일로 패키징하고 배포 패키지를 구성한다. 'EXE 만들어줘', '실행파일 만들어줘', 'PyInstaller', '배포 패키지', 'Phase 6' 요청 시 사용."
---

# Package EXE

PyInstaller로 단일 EXE 파일을 생성하고 배포 패키지를 구성한다.

## 사전 확인 (필수)

코드 전수 점검 후 빌드 시작:

```bash
grep -r "open(" modules/ --include="*.py" | grep -v "BASE_DIR"
# 결과가 있으면 BASE_DIR 패턴으로 수정 후 진행
```

## PyInstaller 빌드 명령

```bash
pyinstaller \
  --onefile \
  --windowed \
  --name 안전점검_모니터링 \
  --add-data "config;config" \
  --icon icon.ico \
  main.py
```

## .spec 파일 (재빌드용)

```python
# 안전점검_모니터링.spec
a = Analysis(
    ['main.py'],
    datas=[('config', 'config')],
    hiddenimports=['PyQt6.sip', 'lxml.etree'],
    ...
)
```

## install.bat

```batch
@echo off
echo playwright 설치 중...
pip install playwright
playwright install chromium
echo.
echo 설치 완료! 이제 안전점검_모니터링.exe를 실행하세요.
pause
```

## 배포 폴더 구조

```
안전점검_모니터링/
├── 안전점검_모니터링.exe
├── install.bat
├── config/
│   ├── sites.xlsx
│   └── keywords.json
├── data/              ← 빈 폴더 (history.json 자동 생성)
└── results/           ← 빈 폴더 (엑셀 자동 저장)
```

## 배포 체크리스트

| 항목 | 확인 방법 |
|------|---------|
| BASE_DIR 경로 전수 확인 | grep으로 절대경로 사용 여부 점검 |
| `--onefile` 빌드 성공 | `dist/` 폴더에 EXE 생성 확인 |
| EXE 단독 실행 | 개발 폴더 없는 환경에서 더블클릭 테스트 |
| config/sites.xlsx 읽기 | EXE 실행 후 사이트 목록 표시 확인 |
| results/ 폴더 자동 생성 | 첫 실행 시 자동 생성 확인 |
| install.bat 동작 | playwright 설치 성공 확인 |
| 바이러스 백신 오탐 | Windows Defender 경고 여부 확인 |

## 흔한 오류 해결

| 오류 | 해결 |
|------|------|
| ModuleNotFoundError: PyQt6.sip | `--hidden-import PyQt6.sip` 추가 |
| lxml 관련 오류 | `--collect-all lxml` 추가 |
| config 파일 못 찾음 | `sys.frozen` 분기 확인, `--add-data` 경로 재확인 |
| 창이 바로 닫힘 | `--windowed` 제거 후 콘솔 오류 확인 |

## 최종 산출물

`dist/안전점검_모니터링.exe` — 이 파일과 `config/` 폴더만 배포하면 됨
