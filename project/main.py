# -*- coding: utf-8 -*-
"""
main.py — 안전점검 모니터링 시스템 Phase 5 (GUI 진입점)

실행:
  python main.py
PyInstaller 패키징(EXE) 시에도 동일하게 동작한다.
"""

import sys
import os

# BASE_DIR 패턴 (EXE/스크립트 양쪽 호환)
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    # 심화 크롤링용 번들 Chromium을 Playwright가 찾도록 경로 지정.
    # (PyInstaller datas → _internal/ms-playwright 에 동봉됨)
    _bundled = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "ms-playwright")
    if os.path.isdir(_bundled):
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", _bundled)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# project/ 를 import 경로에 추가 (modules 패키지 import 보장)
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from PyQt6.QtWidgets import QApplication

from modules.gui_main import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("안전점검 모니터링")
    window = MainWindow(base_dir=BASE_DIR)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
