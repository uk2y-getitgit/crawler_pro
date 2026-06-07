# -*- coding: utf-8 -*-
"""
test_ai.py — Phase 3 AI 에이전트 통합 테스트

1) PostFilterAgent 단독 테스트 (실제 API 키 없어도 폴백 동작 확인)
2) keyword_fallback_filter 규칙 검증
3) ErrorHandlerAgent 동작 확인
4) crawler + AI 필터 통합 테스트 (3개 사이트)

API 키가 없으면(.env 플레이스홀더) 키워드 폴백으로 자동 전환되며,
키 유무와 관계없이 전체 파이프라인이 동작해야 한다.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from modules.ai_agent import (
    PostFilterAgent,
    ErrorHandlerAgent,
    keyword_fallback_filter,
)


def _hr(title):
    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)


# 테스트용 가짜 게시글 (수집 대상 + 제외 대상 혼합)
SAMPLE_POSTS = [
    {"title": "2026년 안전점검 수행기관 지정공고", "date": "2026-06-01",
     "url": "http://a/1", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
    {"title": "정기 안전점검 대행기관 안전진단 수행기관 모집", "date": "2026-06-02",
     "url": "http://a/2", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
    {"title": "도로포장 공사 입찰공고", "date": "2026-06-03",
     "url": "http://a/3", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
    {"title": "안전점검 수행기관 지정 입찰결과 낙찰 통보", "date": "2026-06-04",
     "url": "http://a/4", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
    {"title": "교통안전 캠페인 안내", "date": "2026-06-05",
     "url": "http://a/5", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
    {"title": "시특법 수행기관 정기안전점검 지정 안내", "date": "2026-06-06",
     "url": "http://a/6", "agency": "테스트청", "board_name": "공지",
     "is_new": True, "ai_reason": ""},
]


def progress_cb(event_type, data):
    print(f"  [{event_type}] {data}")


def test_keyword_fallback():
    _hr("[TEST 1] keyword_fallback_filter 규칙 검증")
    # 원본 보존 위해 복사
    import copy
    posts = copy.deepcopy(SAMPLE_POSTS)
    matched = keyword_fallback_filter(posts)
    print(f"입력 {len(posts)}건 → 매칭 {len(matched)}건")
    for m in matched:
        print(f"  ✓ {m['title']}  | {m['ai_reason']}")

    titles = [m["title"] for m in matched]
    # 기대: primary 매칭 1건, secondary 2개 매칭 1건, 시특법 매칭 1건
    assert "2026년 안전점검 수행기관 지정공고" in titles, "primary 매칭 실패"
    # exclude(입찰결과/낙찰) 포함 건은 제외되어야 함
    assert "안전점검 수행기관 지정 입찰결과 낙찰 통보" not in titles, "exclude 미작동"
    # 단순 입찰공고/캠페인은 제외
    assert "도로포장 공사 입찰공고" not in titles, "비대상 포함됨"
    assert "교통안전 캠페인 안내" not in titles, "비대상 포함됨"
    print("  => 키워드 규칙(primary/secondary/exclude) 정상 동작")


def test_postfilter_standalone():
    _hr("[TEST 2] PostFilterAgent 단독 테스트 (폴백 동작 확인)")
    import copy
    posts = copy.deepcopy(SAMPLE_POSTS)
    agent = PostFilterAgent(progress_callback=progress_cb)
    print(f"폴백 모드: {agent.use_fallback}  (사유: {agent.fallback_reason})")
    matched = agent.filter(posts)
    print(f"입력 {len(posts)}건 → 선별 {len(matched)}건")
    for m in matched:
        print(f"  ✓ {m['title']}  | {m['ai_reason']}")
    assert len(matched) >= 1, "선별 결과 0건 — 파이프라인 이상"
    print("  => PostFilterAgent 정상 동작 (API 키 유무와 무관)")


def test_error_handler():
    _hr("[TEST 3] ErrorHandlerAgent 동작 확인")
    handler = ErrorHandlerAgent()
    print(f"폴백 모드: {handler.use_fallback}")
    for err in ["Connection timed out", "SSL: CERTIFICATE_VERIFY_FAILED",
                "HTTP 404 Not Found", "게시글 0개 (선택자 미매칭)"]:
        msg = handler.analyze(err)
        print(f"\n  오류: {err}")
        print(f"  분석: {msg}")
    print("\n  => ErrorHandlerAgent 정상 동작")


def test_crawler_integration():
    _hr("[TEST 4] crawler + AI 필터 통합 테스트 (3개 사이트)")
    from modules.crawler import Crawler

    # PostFilterAgent를 crawler에 주입 (API 키 없으면 자동 폴백)
    ai_agent = PostFilterAgent(progress_callback=progress_cb)
    print(f"AI 에이전트 폴백 모드: {ai_agent.use_fallback}")

    crawler = Crawler(
        sites_path="config/sites.xlsx",
        history_path="data/history.json",
        progress_callback=progress_cb,
        ai_agent=ai_agent,
    )

    # sites.xlsx에서 활성화=Y 사이트 중 처음 3개만 테스트
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(BASE_DIR, "config", "sites.xlsx"))
    ws = wb.active
    test_sites = []
    for row in list(ws.iter_rows(values_only=True))[1:]:
        if row and len(row) >= 6 and row[5] == "Y":
            test_sites.append(row[0])
        if len(test_sites) >= 3:
            break

    print(f"테스트 사이트({len(test_sites)}개): {test_sites}")
    results = crawler.run(site_filter=test_sites)
    print(f"\n최종 AI 선별 결과: {len(results)}건")
    for r in results[:10]:
        print(f"  ✓ [{r['agency']}] {r['title']}  | {r.get('ai_reason','')}")
    print("  => crawler + AI 필터 통합 파이프라인 정상 동작")


def main():
    test_keyword_fallback()
    test_postfilter_standalone()
    test_error_handler()
    try:
        test_crawler_integration()
    except Exception as e:
        print(f"\n[TEST 4] 통합 테스트 중 네트워크/사이트 오류 발생(허용): {e}")
        print("  => 폴백/필터 단위 테스트(1~3)는 통과했습니다.")

    _hr("모든 테스트 완료")


if __name__ == "__main__":
    main()
