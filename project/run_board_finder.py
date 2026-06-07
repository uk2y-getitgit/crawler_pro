# -*- coding: utf-8 -*-
"""
run_board_finder.py — 게시판 등록 도구 실행 스크립트 (Phase 4)

사용법:
  python run_board_finder.py              # 전체 66개 사이트
  python run_board_finder.py 대전광역시   # 단일 기관
  python run_board_finder.py 대전광역시 --yes   # 대화형 확인 없이 자동 저장
"""

import os
import sys

# project/ 를 import 경로에 추가 (어느 위치에서 실행해도 동작)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from modules.board_finder import BoardFinder


def progress_cb(event_type, data):
    print(f"[{event_type}] {data}")


def main():
    args = [a for a in sys.argv[1:]]
    auto_yes = "--yes" in args or "-y" in args
    agencies = [a for a in args if not a.startswith("-")]
    agency_filter = agencies if agencies else None

    finder = BoardFinder(
        sites_path="config/sites.xlsx",
        progress_callback=progress_cb,
    )

    scope = ", ".join(agencies) if agencies else "전체 사이트"
    print(f"게시판 탐색 시작: {scope}")
    results = finder.run_all(agency_filter=agency_filter)

    if not auto_yes:
        results = finder.interactive_review(results)

    added = finder.save_to_sites(results)
    total_boards = sum(len(v) for v in results.values())
    print(f"\n완료: 기관 {len(results)}개, 게시판 {total_boards}개 "
          f"(분리 추가 행 {added}개)")
    print(f"sites.xlsx 갱신: {finder.sites_path}")


if __name__ == "__main__":
    main()
