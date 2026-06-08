# -*- coding: utf-8 -*-
"""
run_full_crawl.py — 활성(Y) 사이트 전체 크롤링 + Gemini AI 필터 + 엑셀 저장

config/sites.xlsx 의 활성화=Y 사이트만 수집(crawler 내부 필터).
심화(타입=심화) 사이트는 Playwright로 자동 렌더링.
수집 결과를 Gemini PostFilterAgent로 '안전점검 수행기관 지정공고'만 선별 후
results/ 폴더에 엑셀로 저장한다.
"""
import os
import sys
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from modules.crawler import Crawler
from modules.reporter import Reporter
from modules.ai_agent import create_post_filter_agent


def progress_cb(event_type, data):
    # 사이트 시작/완료/오류만 간략 출력
    if event_type == "start_site":
        print(f"[{data.get('index')}/{data.get('total')}] {data.get('agency')} …", flush=True)
    elif event_type == "error":
        print(f"   ! 오류: {data.get('agency')} — {data.get('error')}", flush=True)
    elif event_type == "ai_filter_done":
        print(f"   ▶ AI 필터: {data.get('before')}건 → {data.get('after')}건", flush=True)
    elif event_type == "complete":
        print(f"\n=== 수집 완료 ===\n"
              f"  총수집 {data.get('total_collected')} / "
              f"AI선별 {data.get('matched_count')} / "
              f"신규 {data.get('new_count')} / 오류 {data.get('error_count')}", flush=True)


def main():
    t0 = time.time()
    print("AI 필터 에이전트 초기화…", flush=True)
    agent = create_post_filter_agent()
    print(f"  에이전트: {type(agent).__name__} (폴백={getattr(agent,'use_fallback','?')})\n", flush=True)

    crawler = Crawler(
        sites_path="config/sites.xlsx",
        history_path="data/history.json",
        progress_callback=progress_cb,
        ai_agent=agent,
    )

    # site_filter=None → 활성화=Y 전체
    results = crawler.run(site_filter=None)

    reporter = Reporter(results_dir="results")
    out = reporter.save(results, crawler.errors, excluded=crawler.excluded)

    elapsed = round(time.time() - t0)
    print(f"\n엑셀 저장: {out}")
    print(f"총 소요: {elapsed}초 ({elapsed//60}분 {elapsed%60}초)")
    print(f"AI 선별(통과) {len(results)}건 / 제외 {len(crawler.excluded)}건 / "
          f"오류 {len(crawler.errors)}건")


if __name__ == "__main__":
    main()
