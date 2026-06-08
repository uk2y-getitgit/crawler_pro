# -*- mode: python ; coding: utf-8 -*-
"""
안전점검_모니터링.spec — PyInstaller 빌드 설정 (onedir 방식)

빌드:
    pyinstaller 안전점검_모니터링.spec
또는:
    build.bat

배포 구조 (dist/안전점검_모니터링/):
    안전점검_모니터링.exe       ← 실행 파일
    _internal/                   ← PyInstaller 런타임/라이브러리 (자동 생성)
    config/  sites.xlsx, keywords.json
    data/    (빈 폴더, history.json 등 런타임 생성)
    logs/    (빈 폴더, scheduler.log 등 런타임 생성)
    results/ (빈 폴더, 엑셀 결과 저장)

중요 — BASE_DIR 정합성:
    main.py / 모든 모듈은 frozen일 때
        BASE_DIR = os.path.dirname(sys.executable)
    를 사용한다. onedir 방식에서 sys.executable 은
        dist/안전점검_모니터링/안전점검_모니터링.exe
    이므로 BASE_DIR = dist/안전점검_모니터링/ 가 된다.
    따라서 config/ data/ logs/ results/ 는 _internal 안이 아니라
    EXE 와 같은 폴더(외부)에 있어야 하며, 런타임에 쓰기 가능해야 한다.
    -> 이 폴더들은 datas(번들)로 넣지 않고, 빌드 후 build.bat / make_release.py
       에서 EXE 옆으로 복사한다.
"""

import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

APP_NAME = "안전점검_모니터링"

# ---------------------------------------------------------------------------
# Playwright(심화 크롤링) 번들
#   - 패키지 + node 드라이버: collect_all('playwright')
#   - Chromium 브라우저: 시스템 캐시(ms-playwright/chromium-1117)를
#     _internal/ms-playwright 로 동봉 → 다른 PC에서도 심화 동작.
#     (main.py 가 frozen 시 PLAYWRIGHT_BROWSERS_PATH 를 이 경로로 지정)
# ---------------------------------------------------------------------------
pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

_ms = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")
_browser_datas = []
for _sub in ("chromium-1117", "winldd-1007"):
    _src = os.path.join(_ms, _sub)
    if os.path.isdir(_src):
        _browser_datas.append((_src, os.path.join("ms-playwright", _sub)))


# ---------------------------------------------------------------------------
# datas:
#   여기서는 _internal 안에 들어가도 무방한(읽기 전용) 리소스만 넣는다.
#   config/ 는 사용자가 직접 수정하는 파일이므로 EXE 옆 외부 폴더로 복사한다
#   (datas 에 넣으면 _internal/config 로 들어가 BASE_DIR 경로와 불일치).
#   -> config/, data/, results/, logs/ 는 build.bat 후처리에서 복사.
#   Playwright 패키지/드라이버 + Chromium 브라우저는 여기서 동봉한다.
# ---------------------------------------------------------------------------
datas = pw_datas + _browser_datas


# ---------------------------------------------------------------------------
# hiddenimports:
#   동적/조건부 import 되어 PyInstaller 정적 분석이 놓칠 수 있는 패키지.
#   modules/ 에서 실제로 사용하는 서드파티 패키지를 모두 명시한다.
# ---------------------------------------------------------------------------
hiddenimports = [
    # GUI
    "PyQt6",
    "PyQt6.sip",
    "PyQt6.QtCore",
    "PyQt6.QtGui",
    "PyQt6.QtWidgets",
    # HTML 파싱 (bs4 의 lxml 백엔드는 동적 로드)
    "bs4",
    "lxml",
    "lxml.etree",
    "lxml._elementpath",
    # HTTP
    "requests",
    # 엑셀
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    # AI SDK
    "anthropic",
    "google",
    "google.genai",
    "google.genai.types",
    # 환경/스케줄/로깅
    "dotenv",
    "schedule",
    "loguru",
    # 심화 크롤링(Playwright)
    "playwright",
    "playwright.sync_api",
] + pw_hiddenimports
# 주의: PIL(Pillow)은 본 프로젝트에서 실제로 import 되지 않으며 설치되어 있지
# 않으므로 hiddenimports 에 포함하지 않는다(포함 시 불필요한 경고 발생).


a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=pw_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "matplotlib",
        "numpy",
        "pandas",
        "PIL",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,         # onedir: 바이너리는 COLLECT 로 분리
    name=APP_NAME,                 # -> 안전점검_모니터링.exe
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,                 # --windowed (콘솔 창 숨김)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,                 # -> dist/안전점검_모니터링/
)
