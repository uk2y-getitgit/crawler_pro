# -*- coding: utf-8 -*-
"""
poc_menu_nav.py — '헤더 메뉴 클릭 → 게시판 접속' 가능성 검증 POC (Playwright)

목적:
  로그인/JS 때문에 직접 URL 접속이 막힌 사이트에서,
  실제 브라우저로 메인페이지를 열고 헤더 내비게이션 메뉴를 따라가
  공개 게시판(공지/공고/입찰)에 로그인 없이 도달 가능한지 확인한다.

requests 정적 파서가 놓치는 JS 메뉴(onclick, 마우스오버 드롭다운)까지
브라우저 렌더링 후 추출되는지를 보여준다.
"""
import sys
from playwright.sync_api import sync_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

BOARD_HINT = ("공고", "고시", "입찰", "공지", "알림", "사업", "조달", "구매", "계약")


def probe(home_url, click_keyword=None):
    print(f"\n{'='*70}\n대상: {home_url}\n{'='*70}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=UA, ignore_https_errors=True,
                                   viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        try:
            page.goto(home_url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2500)  # JS 메뉴 렌더링 대기
        except Exception as e:
            print(f"  [접속실패] {type(e).__name__}: {str(e)[:80]}")
            browser.close(); return

        # 1) 렌더링 후 전체 메뉴 링크 추출 (정적 파서가 못 보는 것 포함)
        links = page.eval_on_selector_all(
            "header a, nav a, .gnb a, .menu a, [class*=nav] a, [class*=menu] a",
            """els => els.map(e => ({
                t: (e.innerText||'').trim().replace(/\\s+/g,' '),
                h: e.href || ''
            })).filter(x => x.t && x.t.length>=2)"""
        )
        # 중복 제거
        seen, menu = set(), []
        for l in links:
            k = (l["t"], l["h"])
            if k in seen:
                continue
            seen.add(k); menu.append(l)

        board_like = [l for l in menu if any(h in l["t"] for h in BOARD_HINT)]
        print(f"  렌더링 후 메뉴 링크: 총 {len(menu)}개 / 게시판성 메뉴 {len(board_like)}개")
        print("  --- 게시판 후보 메뉴 (상위 15) ---")
        for l in board_like[:15]:
            print(f"    · {l['t'][:24]:24}  {l['h'][:60]}")

        # 2) 지정 키워드 메뉴를 실제 클릭 → 게시판 목록 렌더 여부 확인
        if click_keyword:
            print(f"\n  ▶ '{click_keyword}' 메뉴 클릭 시도")
            try:
                target = page.get_by_role("link", name=click_keyword).first
                target.click(timeout=8000)
                page.wait_for_timeout(2500)
                # 게시판 목록(테이블/리스트 행) 존재 확인
                rows = page.eval_on_selector_all(
                    "table tr, .board li, ul.list li, [class*=board] li",
                    "els => els.length")
                dates = page.eval_on_selector_all(
                    "body", "els => (els[0].innerText.match(/20\\d\\d[.\\-]\\d{1,2}[.\\-]\\d{1,2}/g)||[]).length")
                print(f"    → 이동 후 URL: {page.url[:70]}")
                print(f"    → 목록행 {rows}개, 날짜 {dates}개  "
                      f"=> {'게시판 도달 성공(로그인 불필요)' if rows>3 else '목록 미검출(추가 클릭 필요)'}")
            except Exception as e:
                print(f"    클릭 실패: {type(e).__name__}: {str(e)[:70]}")

        browser.close()


if __name__ == "__main__":
    # 케이스1: 한국수력원자력 — 직접URL은 로그인페이지였음. 공개 메인에서 메뉴 탐색
    probe("https://www.khnp.co.kr/main/index.do", click_keyword="공지사항")
    # 케이스2: 대전 교육청 — 정적수집 실패(일반→심화). 브라우저 렌더 확인
    probe("https://www.dje.go.kr/boardCnts/list.do?boardID=52&m=0205&s=dje")
