---
name: auto-cra-build
description: "안전점검 수행기관 지정공고 자동 모니터링 시스템 개발 오케스트레이터. '시스템 만들어줘', '크롤러 만들어줘', '개발 시작', '다음 단계', '계속 개발', '재실행', '업데이트', '수정해줘' 등 이 프로젝트의 개발·수정·재실행 요청 시 반드시 이 스킬을 사용할 것. Phase 1(환경설정)→Phase 2(크롤러)→Phase 3(AI에이전트)→Phase 4(게시판탐색)→Phase 5(통합+GUI) 순서로 진행하며, 특정 Phase만 재실행 요청도 처리한다."
---

# Auto-CRA Build — 오케스트레이터

66개 관공서 사이트에서 '안전점검 수행기관 지정공고'를 자동 수집하는 PyQt6 + Claude API 시스템을 단계적으로 개발한다.

## Phase 0: 컨텍스트 확인

실행 시작 전 기존 산출물 상태를 확인한다:

1. `project/` 폴더 존재 여부 → 없으면 **초기 실행**
2. 특정 모듈 파일(crawler.py, ai_agent.py 등) 존재 여부 → 있으면 **부분 재실행**
3. 사용자가 "Phase N부터" 또는 특정 모듈 수정 요청 시 → **지정 Phase 실행**

판별 후 사용자에게 실행 계획을 한 줄로 확인한다.

## Phase 1: 환경 구성

**실행 모드: 서브 에이전트**

`setup-agent`를 호출하여:
- Python 가상환경 + requirements.txt 설치
- `.env` 템플릿, `.gitignore` 생성
- `config/sites.xlsx` 재구성 (site_list(전체).xlsx → 게시판URL 컬럼 추가)
- `config/keywords.json` 생성
- 프로젝트 폴더 구조 전체 생성

완료 조건: `check_env.py` 실행 성공

## Phase 2: 기본 크롤러 제작 (5개 사이트 테스트)

**실행 모드: 서브 에이전트 (crawler-agent + qa-agent 병렬)**

`crawler-agent`를 호출하여:
- `modules/crawler.py` 작성 (requests + BeautifulSoup)
- `modules/reporter.py` 작성 (엑셀 저장)
- history.json 신규 판별 로직

크롤러 완성 즉시 `qa-agent`를 호출하여:
- 5개 사이트로 실제 실행 테스트
- 컬럼 shape 정합성, 신규 판별 로직 검증

완료 조건: 5개 사이트 크롤링 성공, 결과 엑셀 생성 확인

## Phase 3: AI 에이전트 통합

**실행 모드: 서브 에이전트**

`ai-integration-agent`를 호출하여:
- `modules/ai_agent.py` 작성
- PostFilterAgent (haiku 모델, 배치 30개) 구현
- ErrorHandlerAgent 구현
- 키워드 폴백 필터 구현

완료 후 `qa-agent` 재호출:
- PostFilterAgent 출력 형식 vs crawler 기대 형식 교차 검증
- 배치 처리 index 오프셋 오류 확인

완료 조건: 5개 사이트 + AI 필터링 통합 테스트 성공

## Phase 4: 게시판 등록 도구

**실행 모드: 서브 에이전트**

`board-finder-dev-agent`를 호출하여:
- `modules/board_finder.py` 작성
- BoardFinderAgent (sonnet 모델) 구현 및 프롬프트 튜닝
- 10개 사이트 테스트 → 프롬프트 개선
- 전체 66개 사이트 게시판 URL 수집

완료 조건: sites.xlsx 게시판URL 컬럼 80% 이상 채워짐

## Phase 5: 통합, GUI, 안정화

**실행 모드: 에이전트 팀 (gui-agent + crawler-agent + qa-agent)**

팀을 구성하여:
1. `gui-agent` — PyQt6 UI 개발 (4가지 스타일 시안 먼저 제시 → 사용자 선택 후 코드 작성)
2. `crawler-agent` — 전체 66개 사이트 확장, playwright 추가, 스케줄러 구현
3. `qa-agent` — E2E 통합 테스트, 오류 시나리오 검증

완료 조건: 전체 파이프라인 실행 성공, GUI 동작 확인

## Phase 6: EXE 패키징

**실행 모드: 서브 에이전트**

`packager-agent`를 호출하여:
- PyInstaller로 단일 EXE 빌드
- 배포 폴더 구조 정리
- install.bat 생성
- 체크리스트 기반 검증

## 데이터 전달 프로토콜

- 중간 산출물은 `_workspace/` 폴더에 저장
- 파일명 컨벤션: `{phase}_{agent}_{artifact}.{ext}`
- 최종 코드는 프로젝트 폴더에 직접 저장

## 에러 핸들링

- Phase 실패 시 1회 재시도 후 재실패 → 사용자에게 상세 오류 보고
- QA 검증 실패 시 해당 Phase 완료 처리하지 않고 수정 후 재검증

## 테스트 시나리오

**정상 흐름**: Phase 1 → check_env.py 성공 → Phase 2 → 5개 사이트 결과 엑셀 생성 → Phase 3 → AI 필터 통합 → Phase 4 → sites.xlsx 채워짐 → Phase 5 → GUI 실행 → Phase 6 → EXE 생성

**에러 흐름**: crawler.py 타임아웃 발생 → logs/crawler.log에 기록 확인 → 해당 사이트 스킵 → 다음 사이트 진행 → 최종 오류로그 시트에 기록
