# -*- coding: utf-8 -*-
"""테스트 실행 스크립트: 처음 5개 사이트로 크롤러/리포터 검증."""

import os
import sys

# project/ 를 import 경로에 추가 (어느 위치에서 실행해도 동작)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from modules.crawler import Crawler
from modules.reporter import Reporter


def progress_cb(event_type, data):
    print(f"[{event_type}] {data}")


def main():
    crawler = Crawler(
        sites_path="config/sites.xlsx",
        history_path="data/history.json",
        progress_callback=progress_cb,
    )

    # 처음 5개 사이트만 테스트
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(BASE_DIR, "config", "sites.xlsx"))
    ws = wb.active
    test_sites = []
    for row in list(ws.iter_rows(values_only=True))[1:6]:  # 헤더 제외 5개
        test_sites.append(row[0])  # 기관명

    print(f"테스트 사이트: {test_sites}")
    results = crawler.run(site_filter=test_sites)
    print(f"\n수집 완료: {len(results)}건")

    reporter = Reporter(results_dir="results")
    out = reporter.save(results, crawler.errors)
    print(f"엑셀 저장 완료: {out}")


if __name__ == "__main__":
    main()
