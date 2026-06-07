# -*- coding: utf-8 -*-
"""
crawler.py — 안전점검 모니터링 시스템 Phase 2

sites.xlsx에서 활성화=Y인 사이트의 게시글 목록을 수집한다.
requests + BeautifulSoup 기반 (playwright는 Phase 5).
"""

import os
import sys
import json
import time
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

try:
    import openpyxl
except ImportError as e:  # pragma: no cover
    raise ImportError("openpyxl가 필요합니다: pip install openpyxl") from e

# .env 로드 (선택)
try:
    from dotenv import load_dotenv
    _dotenv_available = True
except ImportError:
    _dotenv_available = False

# loguru가 있으면 사용, 없으면 표준 logging
try:
    from loguru import logger  # type: ignore
    _USE_LOGURU = True
except ImportError:
    _USE_LOGURU = False
    logger = logging.getLogger("crawler")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        _h = logging.StreamHandler()
        _h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(_h)


# ---------------------------------------------------------------------------
# BASE_DIR (EXE 호환)
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # modules/crawler.py 기준 -> project/
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# .env 로드
if _dotenv_available:
    _env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(_env_path):
        load_dotenv(_env_path)


def _env_int(name, default):
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


CRAWL_DELAY = _env_int("CRAWL_DELAY", 2)
MAX_PAGES = _env_int("MAX_PAGES", 3)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

TIMEOUT = 10          # 초
RETRY = 1             # 재시도 1회 (총 2회 시도)
FAIL_THRESHOLD = 3    # 연속 실패 임계치 -> 점검필요 표시
HISTORY_MONTHS = 6    # 6개월 지난 항목 삭제


# CSS 선택자 자동 탐지 순서
SELECTOR_STRATEGY = [
    "table.bbs_list tr td a",
    "table tr td a",
    "ul.board_list li a",
    "div.board_list a",
    '[class*="list"] a[href]',
]


def _is_valid_post_link(text, href):
    """제목 링크로 보이는지 휴리스틱 검사."""
    if not href:
        return False
    href = href.strip()
    if href in ("#", "", "javascript:;") or href.lower().startswith("javascript"):
        # javascript:fnView(...) 같은 onclick 기반은 URL 없는 게시글로 별도 처리
        return False
    text = (text or "").strip()
    if len(text) < 2:
        return False
    # 페이지네이션/메뉴성 텍스트 제외
    lowered = text.lower()
    skip_words = ("다음", "이전", "처음", "마지막", "more", "목록", "검색", "로그인", "home")
    if lowered in skip_words or text in ("1", "2", "3", "4", "5"):
        return False
    return True


class Crawler:
    def __init__(self, sites_path, history_path, progress_callback=None,
                 base_dir=None, ai_agent=None):
        self.base_dir = base_dir or BASE_DIR
        self.sites_path = self._abspath(sites_path)
        self.history_path = self._abspath(history_path)
        self.progress_callback = progress_callback or (lambda et, data: None)
        # AI 필터 (PostFilterAgent). None이면 키워드 폴백을 사용한다.
        self.ai_agent = ai_agent

        self.errors = []          # 오류로그용
        self.history = self._load_history()
        # 연속 실패 카운트 추적용 (메모리 내)
        self._fail_counts = {}
        # 중단 플래그 (GUI/스레드에서 stop() 호출 시 set). 사이트 경계에서만
        # 검사하므로 진행 중인 단일 사이트 수집·파일 저장이 중간에 끊기지 않는다.
        self._stop_requested = False

    def request_stop(self):
        """크롤링 중단을 요청한다. 현재 사이트 처리를 마친 뒤 안전하게 종료된다."""
        self._stop_requested = True

    # ------------------------------------------------------------------ utils
    def _abspath(self, path):
        if os.path.isabs(path):
            return path
        return os.path.join(self.base_dir, path)

    def _emit(self, event_type, data):
        try:
            self.progress_callback(event_type, data)
        except Exception as e:  # 콜백 오류가 크롤링을 멈추면 안 됨
            logger.warning(f"progress_callback 오류: {e}")

    # ---------------------------------------------------------------- history
    def _load_history(self):
        try:
            with open(self.history_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("history.json 형식 오류")
            data.setdefault("last_updated", "")
            data.setdefault("items", {})
            return data
        except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
            logger.warning(f"history.json 로드 실패, 새로 생성: {e}")
            return {"last_updated": "", "items": {}}

    def _prune_history(self):
        """6개월 지난 항목 삭제."""
        cutoff = datetime.now() - timedelta(days=HISTORY_MONTHS * 30)
        items = self.history.get("items", {})
        removed = 0
        for key in list(items.keys()):
            date_str = items[key].get("date", "")
            dt = _parse_date(date_str)
            if dt is not None and dt < cutoff:
                del items[key]
                removed += 1
        if removed:
            logger.info(f"history 정리: {removed}개 항목 삭제 (6개월 경과)")

    def _save_history(self):
        self.history["last_updated"] = datetime.now().isoformat()
        os.makedirs(os.path.dirname(self.history_path), exist_ok=True)
        with open(self.history_path, "w", encoding="utf-8") as f:
            json.dump(self.history, f, ensure_ascii=False, indent=2)

    @staticmethod
    def _make_key(title, date, url):
        if url:
            return url
        return f"{title}_{date}"

    # ------------------------------------------------------------------ sites
    def _load_sites(self):
        wb = openpyxl.load_workbook(self.sites_path)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        sites = []
        for idx, row in enumerate(rows[1:], start=2):  # 엑셀 행 번호(2부터)
            if row is None or all(c is None for c in row):
                continue
            site = {
                "row": idx,
                "agency": row[0],
                "site_type": row[1],
                "board_name": row[2],
                "url": row[3],
                "page_param": row[4],
                "active": row[5],
                "note": row[6],
            }
            sites.append(site)
        return wb, ws, header, sites

    # ---------------------------------------------------------------- fetch
    def _fetch(self, url):
        """타임아웃 10초, 재시도 1회. (status_code, html|None, error|None)"""
        headers = {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        }
        last_err = None
        for attempt in range(RETRY + 1):
            try:
                resp = requests.get(url, headers=headers, timeout=TIMEOUT,
                                    verify=True)
                return resp.status_code, resp.text, None
            except requests.exceptions.SSLError as e:
                # 일부 공공기관 인증서 문제 -> verify=False 재시도
                try:
                    resp = requests.get(url, headers=headers, timeout=TIMEOUT,
                                        verify=False)
                    return resp.status_code, resp.text, None
                except Exception as e2:
                    last_err = f"SSL/{e2}"
            except requests.exceptions.RequestException as e:
                last_err = str(e)
            if attempt < RETRY:
                time.sleep(1)
        return None, None, last_err

    # ------------------------------------------------------------- extract
    def _extract_posts(self, html, base_url):
        """CSS 선택자 전략 순서대로 시도. (posts, used_selector)"""
        soup = BeautifulSoup(html, "lxml")
        for selector in SELECTOR_STRATEGY:
            try:
                anchors = soup.select(selector)
            except Exception:
                continue
            posts = []
            seen = set()
            for a in anchors:
                title = a.get_text(strip=True)
                href = a.get("href", "")
                if not _is_valid_post_link(title, href):
                    continue
                full_url = urljoin(base_url, href)
                if full_url in seen:
                    continue
                seen.add(full_url)
                date = _find_date_near(a)
                posts.append({"title": title, "url": full_url, "date": date})
            if posts:
                return posts, selector
        return [], None

    def _board_pages(self, site):
        """1..MAX_PAGES 페이지 URL 생성."""
        base = site["url"]
        param = site.get("page_param")
        urls = [base]
        if param:
            for p in range(2, MAX_PAGES + 1):
                sep = "&" if "?" in base else "?"
                urls.append(f"{base}{sep}{param}={p}")
        return urls

    # ----------------------------------------------------------------- run
    def run(self, site_filter=None):
        """
        site_filter: None이면 활성화=Y 전체, 리스트면 해당 기관명만.
        반환: 결과 dict 리스트.
        """
        self._prune_history()
        wb, ws, header, sites = self._load_sites()

        target = []
        for s in sites:
            if s["active"] != "Y":
                continue
            if site_filter is not None and s["agency"] not in site_filter:
                continue
            target.append(s)

        results = []
        sheet_dirty = False

        stopped = False
        for i, site in enumerate(target):
            # 사이트 경계에서 중단 요청 확인 (진행 중 사이트는 끊지 않음)
            if self._stop_requested:
                stopped = True
                logger.info("중단 요청 감지 — 현재까지 수집한 결과로 마무리합니다.")
                self._emit("stopped", {"completed": i, "total": len(target)})
                break

            agency = site["agency"]
            board_name = site["board_name"] or ""
            url = site["url"]

            self._emit("start_site", {
                "index": i + 1, "total": len(target),
                "agency": agency, "url": url,
            })
            logger.info(f"[{i+1}/{len(target)}] {agency} 수집 시작: {url}")

            if not url:
                self._record_error(agency, url, "게시판URL 없음", 0, "스킵")
                self._emit("error", {"agency": agency, "error": "URL 없음"})
                continue

            site_posts, err = self._crawl_site(site)

            if err is not None:
                # 실패 처리 + 연속 실패 카운트
                cnt = self._fail_counts.get(agency, 0) + 1
                self._fail_counts[agency] = cnt
                self._record_error(agency, url, err, RETRY, "실패")
                self._emit("error", {"agency": agency, "error": err, "fail_count": cnt})
                logger.warning(f"  {agency} 실패: {err} (연속 {cnt}회)")
                if cnt >= FAIL_THRESHOLD:
                    self._mark_inspection_needed(ws, site)
                    sheet_dirty = True
            else:
                self._fail_counts[agency] = 0
                if not site_posts:
                    logger.warning(f"  {agency}: 게시글 0개 — 구조 변경 의심")
                    self._record_error(agency, url, "게시글 0개 (선택자 미매칭)",
                                       RETRY, "경고")
                for post in site_posts:
                    key = self._make_key(post["title"], post["date"], post["url"])
                    is_new = key not in self.history["items"]
                    result = {
                        "agency": agency,
                        "board_name": board_name,
                        "title": post["title"],
                        "date": post["date"],
                        "url": post["url"],
                        "is_new": is_new,
                        "ai_reason": "",
                    }
                    results.append(result)
                    if is_new:
                        self.history["items"][key] = {
                            "title": post["title"],
                            "date": post["date"],
                            "agency": agency,
                        }
                        self._emit("post_found", result)

            # 사이트 간 딜레이 (마지막 사이트 제외).
            # 중단 요청이 들어오면 즉시 빠져나올 수 있도록 1초 단위로 끊어 잔다.
            if i < len(target) - 1:
                for _ in range(CRAWL_DELAY):
                    if self._stop_requested:
                        break
                    time.sleep(1)

        # 점검필요 표시가 있었으면 sites.xlsx 저장
        if sheet_dirty:
            try:
                wb.save(self.sites_path)
                logger.info("sites.xlsx 비고란 '점검필요' 저장 완료")
            except Exception as e:
                logger.warning(f"sites.xlsx 저장 실패: {e}")

        # AI 필터 적용: 수집한 게시글 중 '안전점검 수행기관 지정공고'만 선별
        collected = len(results)
        results = self._apply_ai_filter(results)

        self._save_history()
        self._emit("complete", {
            "total_collected": collected,
            "matched_count": len(results),
            "new_count": sum(1 for r in results if r["is_new"]),
            "error_count": len(self.errors),
            "stopped": stopped,
        })
        logger.info(f"수집 완료: 총 {collected}건 수집, "
                    f"AI 선별 {len(results)}건, "
                    f"신규 {sum(1 for r in results if r['is_new'])}건, "
                    f"오류 {len(self.errors)}건")
        return results

    def _apply_ai_filter(self, results):
        """
        수집된 게시글에 AI 필터를 적용한다.
        - ai_agent가 있으면 PostFilterAgent.filter() 호출
        - 없으면 키워드 폴백 필터(keyword_fallback_filter) 사용
        선별된 게시글만 ai_reason이 채워진 채로 반환된다.
        """
        if not results:
            return results
        try:
            if self.ai_agent is not None:
                filtered = self.ai_agent.filter(results)
            else:
                # 지연 임포트: ai_agent 미사용 환경에서도 crawler 단독 동작.
                # 패키지/단독 실행 양쪽 import 방식 모두 지원.
                try:
                    from modules.ai_agent import keyword_fallback_filter
                except ImportError:
                    from ai_agent import keyword_fallback_filter
                filtered = keyword_fallback_filter(results)
            self._emit("ai_filter_done", {
                "before": len(results), "after": len(filtered),
            })
            return filtered
        except Exception as e:
            # 필터 자체가 실패하면 수집 결과를 그대로 반환(데이터 유실 방지)
            logger.warning(f"AI 필터 적용 실패, 원본 반환: {e}")
            self._emit("error", {"agency": "AI필터", "error": str(e)})
            return results

    def _crawl_site(self, site):
        """단일 사이트 전체 페이지 수집. (posts, error)"""
        all_posts = []
        seen_urls = set()
        for page_url in self._board_pages(site):
            status, html, err = self._fetch(page_url)
            if err is not None:
                # 첫 페이지 실패면 사이트 실패, 2페이지 이후 실패는 무시
                if page_url == site["url"]:
                    return [], err
                else:
                    logger.warning(f"  {site['agency']} 페이지 실패(무시): {err}")
                    break
            if status in (404, 503) or (status is not None and status >= 400):
                if page_url == site["url"]:
                    return [], f"HTTP {status}"
                else:
                    logger.warning(f"  {site['agency']} HTTP {status} (페이지 스킵)")
                    break
            posts, selector = self._extract_posts(html, page_url)
            if selector:
                logger.info(f"  선택자 매칭: '{selector}' -> {len(posts)}건")
            for p in posts:
                if p["url"] in seen_urls:
                    continue
                seen_urls.add(p["url"])
                all_posts.append(p)
        return all_posts, None

    # ----------------------------------------------------- error / sheet ops
    def _record_error(self, agency, url, message, retries, outcome):
        self.errors.append({
            "agency": agency,
            "url": url or "",
            "error": message,
            "retries": retries,
            "outcome": outcome,
        })

    def _mark_inspection_needed(self, ws, site):
        """sites.xlsx 비고(7번째 컬럼, G열)에 '점검필요' 표시."""
        try:
            cell = ws.cell(row=site["row"], column=7)
            existing = cell.value or ""
            if "점검필요" not in str(existing):
                cell.value = (str(existing) + " 점검필요").strip() if existing \
                    else "점검필요"
            logger.warning(f"  {site['agency']} 연속 {FAIL_THRESHOLD}회 실패 "
                           f"-> 비고 '점검필요' 표시")
        except Exception as e:
            logger.warning(f"비고 표시 실패: {e}")


# ---------------------------------------------------------------------------
# 날짜 파싱 헬퍼
# ---------------------------------------------------------------------------
_DATE_FORMATS = [
    "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d",
    "%Y년 %m월 %d일", "%y-%m-%d", "%y.%m.%d",
    "%Y-%m-%d %H:%M", "%Y.%m.%d %H:%M",
]

import re

_DATE_RE = re.compile(
    r"(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})"
)


def _parse_date(date_str):
    if not date_str:
        return None
    date_str = date_str.strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    m = _DATE_RE.search(date_str)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _normalize_date(raw):
    """추출한 날짜 문자열을 YYYY-MM-DD로 정규화. 실패 시 원문 반환."""
    dt = _parse_date(raw)
    if dt:
        return dt.strftime("%Y-%m-%d")
    return raw.strip() if raw else ""


def _find_date_near(anchor):
    """앵커 주변(같은 tr/li, 형제 td 등)에서 날짜 문자열을 찾는다."""
    # 1) 부모 행(tr/li) 전체 텍스트에서 정규식 탐색
    container = anchor.find_parent(["tr", "li"])
    if container is None:
        container = anchor.parent
    if container is not None:
        text = container.get_text(" ", strip=True)
        m = _DATE_RE.search(text)
        if m:
            return _normalize_date(m.group(0))
    # 2) 앵커 자체 텍스트
    m = _DATE_RE.search(anchor.get_text(" ", strip=True))
    if m:
        return _normalize_date(m.group(0))
    return ""
