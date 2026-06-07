# -*- coding: utf-8 -*-
"""
scheduler.py — 안전점검 모니터링 시스템 Phase 5

schedule 라이브러리를 QThread로 감싼 자동 실행 스케줄러.

설계 원칙(Qt 스레드 안전):
  스케줄러 스레드는 직접 CrawlerThread를 start()하지 않는다.
  대신 지정 시각이 되면 `trigger` 시그널만 emit하고,
  메인 스레드(MainWindow)가 이 시그널을 받아 CrawlerThread를 시작한다.
  → QThread/QObject는 자신을 만든 스레드에서 다뤄야 안전하므로,
    실제 크롤링 시작 책임을 메인 스레드로 넘긴다.

설정(.env):
  SCHEDULE_ENABLED=Y|N
  SCHEDULE_TIME=HH:MM
"""

from __future__ import annotations

import os
import logging
from datetime import datetime

from PyQt6.QtCore import QThread, pyqtSignal

try:
    import schedule
    _schedule_available = True
except ImportError:
    _schedule_available = False

try:
    from loguru import logger  # type: ignore
except ImportError:
    logger = logging.getLogger("scheduler")


class SchedulerThread(QThread):
    """지정 시각마다 trigger 시그널을 emit하는 스케줄러 스레드."""

    trigger = pyqtSignal()          # 크롤링 실행 요청 (메인 스레드가 수신)
    status = pyqtSignal(str)        # 상태 메시지(로그용)

    def __init__(self, run_time="09:00", base_dir=None, parent=None):
        super().__init__(parent)
        self.run_time = run_time
        self.base_dir = base_dir
        self._running = False

    def run(self):
        if not _schedule_available:
            self.status.emit("schedule 라이브러리 미설치 — 스케줄러 비활성화")
            return

        self._running = True
        schedule.clear()
        try:
            schedule.every().day.at(self.run_time).do(self._fire)
        except Exception as e:
            self.status.emit(f"스케줄 시각 설정 오류({self.run_time}): {e}")
            self._running = False
            return

        self.status.emit(f"스케줄러 시작 — 매일 {self.run_time} 자동 실행")
        # 30초 간격으로 보류된 작업 확인. msleep로 끊어 자며 stop에 빠르게 반응.
        while self._running:
            try:
                schedule.run_pending()
            except Exception as e:
                self.status.emit(f"스케줄 실행 오류: {e}")
            for _ in range(30):
                if not self._running:
                    break
                self.msleep(1000)

        schedule.clear()
        self.status.emit("스케줄러 중지됨")

    def _fire(self):
        """스케줄 시각 도달 — 로그 기록 후 trigger emit."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg = f"[{now}] 스케줄 트리거 — 자동 크롤링 시작"
        self.status.emit(msg)
        self._write_log(msg)
        self.trigger.emit()

    def _write_log(self, message):
        """실행 이력을 logs/scheduler.log에 기록."""
        try:
            base = self.base_dir or os.getcwd()
            logs_dir = os.path.join(base, "logs")
            os.makedirs(logs_dir, exist_ok=True)
            log_path = os.path.join(logs_dir, "scheduler.log")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(message + "\n")
        except OSError as e:
            logger.warning(f"스케줄 로그 기록 실패: {e}")

    def stop(self):
        """스케줄러 루프 종료 요청."""
        self._running = False
