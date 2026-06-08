# -*- coding: utf-8 -*-
"""
playwright_crawler.py — 심화(JS 렌더링) 크롤링 엔진

requests로 수집되지 않는 사이트(타입=심화)를 실제 브라우저(Chromium)로
렌더링하여 HTML을 얻는다. 필요 시 헤더메뉴 클릭·검색버튼 클릭 등
상호작용(actions)을 수행한 뒤의 DOM을 반환한다.

crawler.py는 site_type=='심화'일 때만 이 모듈을 지연 임포트하여 사용한다.
(일반 사이트는 requests 경로를 그대로 사용 → 성능 유지)

actions 스펙 (사이트별 진입 시나리오, 선택):
  [{"type": "click_text",     "value": "입찰공고"},   # 링크/버튼 텍스트 클릭
   {"type": "click_selector", "value": "#searchBtn"}, # CSS 선택자 클릭
   {"type": "fill",  "selector": "#kwd", "value": "안전점검"},
   {"type": "wait",  "value": 2000}]                  # ms 대기
"""
from __future__ import annotations

import logging

try:
    from loguru import logger  # type: ignore
except ImportError:
    logger = logging.getLogger("playwright_crawler")

# Playwright는 무거우므로 import 자체를 지연시킨다.
_PW_AVAILABLE = None  # None=미확인, True/False


def _check_playwright():
    global _PW_AVAILABLE
    if _PW_AVAILABLE is None:
        try:
            import playwright.sync_api  # noqa: F401
            _PW_AVAILABLE = True
        except ImportError:
            _PW_AVAILABLE = False
    return _PW_AVAILABLE


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

DEFAULT_TIMEOUT = 30000   # ms (페이지 이동)
RENDER_WAIT = 2500        # ms (JS 렌더링 대기)


class DeepFetcher:
    """
    Playwright 기반 렌더링 페처. with 문으로 브라우저 수명을 관리한다.

        with DeepFetcher() as df:
            status, html, final_url, err = df.fetch(url, actions=...)

    여러 URL을 한 인스턴스로 순차 처리하면 브라우저를 재사용하여 빠르다.
    (Playwright sync API는 스레드 공유 불가 → 사이트는 순차 처리할 것)
    """

    def __init__(self, headless: bool = True, timeout: int = DEFAULT_TIMEOUT,
                 render_wait: int = RENDER_WAIT):
        self.headless = headless
        self.timeout = timeout
        self.render_wait = render_wait
        self._pw = None
        self._browser = None
        self._ctx = None

    # ---------------------------------------------------------------- lifecycle
    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    def start(self):
        if not _check_playwright():
            raise RuntimeError(
                "playwright 미설치 — pip install playwright && playwright install chromium")
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=self.headless)
        self._ctx = self._browser.new_context(
            user_agent=USER_AGENT,
            ignore_https_errors=True,           # 정부 GPKI 인증서 대응
            viewport={"width": 1400, "height": 900},
        )
        self._ctx.set_default_timeout(self.timeout)

    def close(self):
        for obj, name in ((self._ctx, "context"), (self._browser, "browser")):
            try:
                if obj:
                    obj.close()
            except Exception as e:
                logger.warning(f"{name} close 오류: {e}")
        try:
            if self._pw:
                self._pw.stop()
        except Exception as e:
            logger.warning(f"playwright stop 오류: {e}")
        self._pw = self._browser = self._ctx = None

    # ------------------------------------------------------------------- fetch
    def fetch(self, url: str, actions=None):
        """
        url을 렌더링하고 (status_code, html, final_url, error)를 반환한다.
        actions가 있으면 페이지 로드 후 순서대로 수행한다.
        실패 시 (None, None, url, '사유').
        """
        if self._ctx is None:
            self.start()
        page = self._ctx.new_page()
        status = None
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=self.timeout)
            status = resp.status if resp else None
            # AJAX 목록 로딩(새올 eminwon 등)이 끝나도록 네트워크 안정 대기
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass
            page.wait_for_timeout(self.render_wait)

            if actions:
                self._run_actions(page, actions)

            # 페이지가 내부 네비게이션 중이면 content()가 실패 → 잠시 후 재시도
            html, final = None, url
            for _ in range(3):
                try:
                    html = page.content()
                    final = page.url
                    break
                except Exception:
                    page.wait_for_timeout(1500)
            if html is None:
                html = page.content()  # 마지막 시도(실패 시 예외 → except로)
            return status, html, final, None
        except Exception as e:
            return status, None, url, f"{type(e).__name__}: {str(e)[:120]}"
        finally:
            try:
                page.close()
            except Exception:
                pass

    # ----------------------------------------------------------------- actions
    def _run_actions(self, page, actions):
        """사이트별 진입 시나리오 수행 (헤더메뉴 클릭, 검색 등)."""
        for act in actions:
            atype = act.get("type")
            try:
                if atype == "click_text":
                    page.get_by_role("link", name=act["value"]).first.click(timeout=8000)
                elif atype == "click_button":
                    page.get_by_role("button", name=act["value"]).first.click(timeout=8000)
                elif atype == "click_selector":
                    page.locator(act["value"]).first.click(timeout=8000)
                elif atype == "fill":
                    page.locator(act["selector"]).fill(act.get("value", ""), timeout=8000)
                elif atype == "wait":
                    page.wait_for_timeout(int(act.get("value", 1000)))
                else:
                    logger.warning(f"알 수 없는 action: {atype}")
                    continue
                page.wait_for_timeout(self.render_wait)
            except Exception as e:
                logger.warning(f"action 실패({atype}={act.get('value')}): {type(e).__name__}")


def render_html(url: str, actions=None, headless: bool = True):
    """단발성 렌더링 헬퍼. (status, html, final_url, error)."""
    with DeepFetcher(headless=headless) as df:
        return df.fetch(url, actions=actions)
