# -*- coding: utf-8 -*-
"""
make_release.py — 안전점검 모니터링 시스템 배포 패키지 정리 스크립트

build.bat (PyInstaller onedir) 빌드가 끝난 뒤 실행한다.
dist/안전점검_모니터링/ 의 산출물을 날짜 기반 릴리스 폴더로 정리/복사한다.

사용:
    python deploy/make_release.py

생성 결과 (예: 2026-06-05 실행 시):
    안전점검_모니터링_20260605/
    ├── 안전점검_모니터링.exe        (+ _internal/ : PyInstaller 런타임)
    ├── install.bat
    ├── README_배포.txt
    ├── .env                          (있으면 복사, 없으면 .env.example 안내)
    ├── config/
    │   ├── sites.xlsx
    │   └── keywords.json
    ├── data/                         (빈 폴더)
    ├── logs/                         (빈 폴더)
    └── results/                      (빈 폴더)

주의: onedir 방식이므로 EXE 와 _internal/ 폴더, 그리고 config/data/logs/results
      는 반드시 같은 폴더에 함께 있어야 한다 (frozen BASE_DIR = EXE 폴더).
"""

from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime

# 경로 기준: deploy/ 의 상위 = project 루트
DEPLOY_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(DEPLOY_DIR)

APP_NAME = "안전점검_모니터링"
DIST_APP_DIR = os.path.join(PROJECT_DIR, "dist", APP_NAME)


def _log(msg: str) -> None:
    print(f"[make_release] {msg}")


def _copytree(src: str, dst: str) -> None:
    """폴더 전체 복사 (대상이 있으면 병합/덮어쓰기)."""
    shutil.copytree(src, dst, dirs_exist_ok=True)


def main() -> int:
    # 1) 빌드 산출물 확인
    if not os.path.isdir(DIST_APP_DIR):
        _log(f"빌드 산출물을 찾을 수 없습니다: {DIST_APP_DIR}")
        _log("먼저 build.bat 으로 PyInstaller 빌드를 완료하세요.")
        return 1

    # 2) 날짜 기반 릴리스 폴더명 (예: 안전점검_모니터링_20260605)
    stamp = datetime.now().strftime("%Y%m%d")
    release_name = f"{APP_NAME}_{stamp}"
    release_dir = os.path.join(PROJECT_DIR, "dist", release_name)

    if os.path.isdir(release_dir):
        _log(f"기존 릴리스 폴더 삭제: {release_dir}")
        shutil.rmtree(release_dir)
    os.makedirs(release_dir, exist_ok=True)
    _log(f"릴리스 폴더 생성: {release_dir}")

    # 3) EXE + _internal/ (그 외 빌드 산출물) 복사
    #    dist/안전점검_모니터링/ 의 내용을 릴리스 폴더로 그대로 복사.
    #    (config/data/logs/results 는 아래에서 명시적으로 보장)
    _log("EXE 및 런타임(_internal) 복사 중...")
    for entry in os.listdir(DIST_APP_DIR):
        src = os.path.join(DIST_APP_DIR, entry)
        dst = os.path.join(release_dir, entry)
        if os.path.isdir(src):
            _copytree(src, dst)
        else:
            shutil.copy2(src, dst)

    # 4) config/ (sites.xlsx, keywords.json) 보장
    src_config = os.path.join(PROJECT_DIR, "config")
    dst_config = os.path.join(release_dir, "config")
    if os.path.isdir(src_config):
        _copytree(src_config, dst_config)
        _log("config/ 복사 완료 (sites.xlsx, keywords.json)")
    else:
        _log(f"[경고] config 폴더가 없습니다: {src_config}")

    # 필수 config 파일 점검
    for fname in ("sites.xlsx", "keywords.json"):
        if not os.path.exists(os.path.join(dst_config, fname)):
            _log(f"[경고] config/{fname} 가 릴리스에 없습니다. 수동 확인 필요.")

    # 5) 빈 작업 폴더 생성 (data, logs, results)
    for sub in ("data", "logs", "results"):
        os.makedirs(os.path.join(release_dir, sub), exist_ok=True)
    _log("빈 작업 폴더 생성 완료 (data, logs, results)")

    # 6) install.bat (project 루트에서 복사)
    src_install = os.path.join(PROJECT_DIR, "install.bat")
    if os.path.exists(src_install):
        shutil.copy2(src_install, os.path.join(release_dir, "install.bat"))
        _log("install.bat 복사 완료")
    else:
        _log("[경고] install.bat 을 찾을 수 없습니다.")

    # 7) 배포 안내문 (deploy/README_배포.txt) 복사
    src_readme = os.path.join(DEPLOY_DIR, "README_배포.txt")
    if os.path.exists(src_readme):
        shutil.copy2(src_readme, os.path.join(release_dir, "README_배포.txt"))
        _log("README_배포.txt 복사 완료")
    else:
        _log("[경고] deploy/README_배포.txt 를 찾을 수 없습니다.")

    # 8) .env (있으면 복사). 없으면 안내만.
    src_env = os.path.join(PROJECT_DIR, ".env")
    if os.path.exists(src_env):
        shutil.copy2(src_env, os.path.join(release_dir, ".env"))
        _log(".env 복사 완료 (API 키 포함 여부 확인 후 배포하세요!)")
    else:
        _log(".env 가 없습니다. 사용자가 README 안내에 따라 직접 생성합니다.")

    _log("=" * 50)
    _log(f"배포 패키지 완성: {release_dir}")
    _log("이 폴더를 압축(zip)하여 배포하세요.")
    _log("=" * 50)
    return 0


if __name__ == "__main__":
    sys.exit(main())
