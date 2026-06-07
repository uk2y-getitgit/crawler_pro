# -*- coding: utf-8 -*-
"""
run_board_finder_auto.py — 게시판 자동 수집 + JSON 결과 저장

대화형 확인 없이 전체 66개 사이트를 수집하고
data/board_finder_result.json 에 저장한다.
실패/미수집 사이트는 data/board_finder_failed.json 에 따로 저장.
"""

import os, sys, json
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from modules.board_finder import BoardFinder

LOG_LINES = []

def progress_cb(event_type, data):
    ts = datetime.now().strftime("%H:%M:%S")
    msg = f"[{ts}][{event_type}] {data}"
    print(msg, flush=True)
    LOG_LINES.append(msg)

def main():
    finder = BoardFinder(
        sites_path=os.path.join(BASE_DIR, "config", "sites.xlsx"),
        progress_callback=progress_cb,
    )

    print("=" * 60)
    print(f"게시판 자동 수집 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    results = finder.run_all()

    # ── sites.xlsx 저장 ──
    added = finder.save_to_sites(results)

    # ── 결과 JSON 저장 ──
    os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
    result_path = os.path.join(BASE_DIR, "data", "board_finder_result.json")
    failed_path = os.path.join(BASE_DIR, "data", "board_finder_failed.json")

    # 성공/실패 분류
    success = {}
    failed  = {}
    for agency, boards in results.items():
        if boards:
            success[agency] = [{"name": b.get("name",""), "url": b.get("url","")} for b in boards]
        else:
            failed[agency] = {"boards": [], "note": "게시판 미탐지 — 수동 등록 필요"}

    # errors 속성이 있으면 실패 목록에 추가
    if hasattr(finder, "errors") and finder.errors:
        for e in finder.errors:
            agency = e.get("agency", "알수없음")
            failed.setdefault(agency, {"boards": [], "note": e.get("error", "")})

    with open(result_path, "w", encoding="utf-8") as f:
        json.dump({"updated_at": datetime.now().isoformat(), "results": success}, f,
                  ensure_ascii=False, indent=2)

    with open(failed_path, "w", encoding="utf-8") as f:
        json.dump({"updated_at": datetime.now().isoformat(), "failed": failed}, f,
                  ensure_ascii=False, indent=2)

    # ── 요약 출력 ──
    print("\n" + "=" * 60)
    print(f"수집 완료: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  성공: {len(success)}개 기관")
    print(f"  실패/미탐지: {len(failed)}개 기관")
    total_boards = sum(len(v) for v in success.values())
    print(f"  총 게시판: {total_boards}개")
    print(f"  sites.xlsx 신규 추가 행: {added}개")
    print(f"\n  결과 파일: data/board_finder_result.json")
    print(f"  실패 파일: data/board_finder_failed.json")
    print("=" * 60)

if __name__ == "__main__":
    main()
