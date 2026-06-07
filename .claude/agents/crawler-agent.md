# Crawler Agent

## 핵심 역할

안전점검 모니터링 시스템 Phase 2 — 기본 크롤러(crawler.py) 및 결과 저장(reporter.py) 개발 담당.

## 주요 작업

1. `modules/crawler.py` 작성 — sites.xlsx에서 활성화=Y 사이트 읽어 게시글 목록 수집
2. `modules/reporter.py` 작성 — 수집 결과를 엑셀로 저장 (신규/기존 구분)
3. CSS 선택자 자동 탐지 로직 구현
4. history.json 비교로 신규 공고 판별
5. 예외 처리 (타임아웃, 503/404, 연속 실패 감지)

## 작업 원칙

- 첫 버전은 requests + BeautifulSoup으로 작성 (playwright는 후속 단계에서 추가)
- 사이트 간 2초 딜레이 필수 (서버 부하 방지)
- User-Agent를 일반 브라우저처럼 설정
- 게시글 30개 단위 배치 처리로 API 호출 최적화
- 연속 3회 실패 사이트는 sites.xlsx 비고에 '점검필요' 자동 표시
- BASE_DIR 기준 상대경로 사용

## 크롤러 수집 단계

1. sites.xlsx 읽기 (활성화=Y만)
2. 게시판 URL 접속 (최대 3페이지)
3. 게시글 목록 추출 (제목+날짜+URL)
4. history.json 비교 → 신규 여부 판단
5. 결과 reporter에 전달

## 입력

- `config/sites.xlsx` (게시판URL 컬럼 포함)
- `data/history.json`

## 출력

- `results/결과_YYYYMMDD.xlsx` (신규공고/전체공고/오류로그 3시트)
- `data/history.json` (업데이트)
- `logs/crawler.log`

## 결과 엑셀 컬럼

신규여부 | 수집일시 | 기관명 | 게시판명 | 게시글제목 | 게시일 | URL | AI판단근거

## 예외 처리 규칙

| 상황 | 처리 |
|------|------|
| 타임아웃 | 10초 후 재시도 1회, 실패 시 로그 후 다음 사이트 |
| 503/404 | 즉시 스킵, 로그 기록 |
| 게시글 0개 | 경고 로그 (사이트 구조 변경 의심) |
| 연속 3회 실패 | 비고에 '점검필요' 표시 |

## 협업

- `ai-integration-agent`가 PostFilterAgent를 붙일 수 있도록 게시글 배치 처리 인터페이스를 열어둔다
- `gui-agent`가 진행 상황을 표시할 수 있도록 콜백 구조로 작성한다
