# -*- coding: utf-8 -*-
"""
crawler_thread.py — 안전점검 모니터링 시스템 Phase 5

Crawler를 QThread로 백그라운드 실행하는 래퍼.
GUI 메인 스레드를 멈추지 않고 크롤링을 수행하며, 진행 상황을 시그널로 전달한다.

주의(시그널 네이밍):
  QThread에는 이미 내장 시그널 `started`, `finished`가 존재한다.
  내장 `finished`(인자 없음)를 사용자 정의 `finished(list, list)`로 덮어쓰면
  스레드 종료·자원정리에 쓰이는 내장 시그널이 가려져 위험하다.
  따라서 완료 시그널은 `result_ready`라는 별도 이름을 사용한다.
  (지침서의 finished(list, list)와 동일한 역할 — 충돌 회피용 개명)
"""

from __future__ import annotations

import logging

from PyQt6.QtCore import QThread, pyqtSignal

try:
    from modules.crawler import Crawler
    from modules.ai_agent import create_post_filter_agent
except ImportError:  # 단독 실행/패키징 호환
    from crawler import Crawler
    from ai_agent import create_post_filter_agent

try:
    from loguru import logger  # type: ignore
except ImportError:
    logger = logging.getLogger("crawler_thread")


class CrawlerThread(QThread):
    # 시그널 정의
    # pyqtSignal(str, dict) — PyQt6에서 dict는 QVariant로 매핑되어 정상 동작한다.
    progress = pyqtSignal(str, dict)        # event_type, data
    result_ready = pyqtSignal(list, list)   # results, errors  (지침서의 finished 역할)
    error = pyqtSignal(str)                 # 치명 오류 메시지

    def __init__(self, sites_path, history_path, site_filter=None,
                 base_dir=None, parent=None):
        super().__init__(parent)
        self.sites_path = sites_path
        self.history_path = history_path
        self.site_filter = site_filter
        self.base_dir = base_dir
        self._crawler = None  # stop()에서 참조
        self.excluded = []    # 제외 게시글(날짜/AI) — 완료 후 GUI가 참조

    # ------------------------------------------------------------------ run
    def run(self):
        """QThread 진입점. 별도 스레드에서 실행된다."""
        try:
            # AI 필터 에이전트 생성 (키 없으면 내부적으로 키워드 폴백 모드)
            ai_agent = create_post_filter_agent()

            crawler = Crawler(
                sites_path=self.sites_path,
                history_path=self.history_path,
                progress_callback=self._on_progress,
                base_dir=self.base_dir,
                ai_agent=ai_agent,
            )
            self._crawler = crawler

            results = crawler.run(site_filter=self.site_filter)
            self.excluded = crawler.excluded
            # 완료 — 결과/오류 전달. emit 후 run()이 반환되면 스레드가 자연 종료된다.
            self.result_ready.emit(results, crawler.errors)
        except Exception as e:  # 치명 오류 — 스레드는 정상 종료시키되 알림
            logger.exception("CrawlerThread 실행 중 치명 오류")
            self.error.emit(str(e))
        # run()이 반환되면 QThread는 finished(내장) 시그널을 내고 종료한다.

    # -------------------------------------------------------------- callback
    def _on_progress(self, event_type, data):
        """Crawler의 progress_callback → Qt 시그널 변환.

        이 콜백은 워커 스레드 컨텍스트에서 호출되지만, pyqtSignal.emit()은
        스레드 안전하며 연결된 슬롯은 수신자(메인 윈도우)의 스레드에서
        큐 연결로 실행된다.
        """
        # data가 dict가 아닐 경우를 대비해 방어적으로 감싼다.
        if not isinstance(data, dict):
            data = {"value": data}
        self.progress.emit(str(event_type), data)

    # ------------------------------------------------------------------ stop
    def stop(self):
        """크롤링 중단 요청. 현재 사이트 처리를 마친 뒤 안전하게 멈춘다.

        - history.json / sites.xlsx 저장은 run() 종료부에서 수행되므로
          중단해도 파일이 손상되지 않는다(사이트 경계에서만 중단).
        """
        if self._crawler is not None:
            self._crawler.request_stop()
