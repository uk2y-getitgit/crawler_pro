---
name: build-crawler
description: "안전점검 모니터링 시스템 크롤러(crawler.py) + 결과 저장(reporter.py) 개발. sites.xlsx에서 게시판 URL을 읽어 게시글을 수집하고 엑셀로 저장한다. 'Phase 2', '크롤러 만들어줘', 'crawler 개발', '결과 저장 만들어줘' 요청 시 사용."
---

# Build Crawler

게시판 URL을 순회하며 게시글 목록을 수집하고 엑셀로 저장하는 핵심 크롤러를 개발한다.

## 개발 순서

1. `modules/crawler.py` 먼저 작성 후 단독 테스트
2. `modules/reporter.py` 작성 후 연결
3. 5개 사이트로 통합 테스트

## crawler.py 핵심 구조

```python
class Crawler:
    def __init__(self, sites_path, history_path, progress_callback=None):
        # progress_callback: GUI QThread 시그널 연결용
        ...

    def run(self, site_filter=None):
        # site_filter=None이면 활성화=Y 전체 실행
        # site_filter=['대전광역시'] 이면 해당 사이트만
        ...

    def _fetch_posts(self, board_url, max_pages=3):
        # requests → BeautifulSoup으로 게시글 목록 추출
        # JS 렌더링 필요 시 playwright로 폴백
        ...

    def _extract_posts(self, html, board_url):
        # CSS 선택자 자동 탐지 로직
        # 반환: [{"title": ..., "url": ..., "date": ...}]
        ...
```

## CSS 선택자 자동 탐지 전략

다음 순서로 시도하여 게시글 목록을 추출한다:
1. `table tr td a` — 가장 일반적인 테이블 구조
2. `ul.bbs-list li a`, `div.board-list a` — 리스트 구조
3. `[class*="list"] a`, `[class*="board"] a` — 클래스명 패턴 매칭
4. 앞의 모든 시도 실패 시 → 경고 로그 + 게시글 0개 반환

## reporter.py 엑셀 구조

3개 시트:
- **신규공고**: 🆕 신규, 노란 배경 강조, 상단 배치
- **전체공고**: 누적 수집 전체, 최신순 정렬
- **오류로그**: 접속 실패 사이트, 원인, 재시도 여부

공통 컬럼: `신규여부 | 수집일시 | 기관명 | 게시판명 | 게시글제목 | 게시일 | URL | AI판단근거`

## history.json 구조

```json
{
  "last_updated": "2026-06-05T08:00:00",
  "items": {
    "https://...게시글URL...": {"title": "...", "date": "2026-06-01", "agency": "대전광역시"},
    ...
  }
}
```

URL이 없는 사이트는 `{title}_{date}` 조합을 키로 사용한다.
6개월 지난 항목은 자동 삭제한다.

## 예외 처리 상세

```python
# 타임아웃
response = requests.get(url, timeout=10, headers=HEADERS)
# 실패 시 1회 재시도, 그 후 실패 → 로그 후 continue

# 연속 실패 카운터
if fail_count >= 3:
    update_sites_xlsx_remark(agency, "점검필요")
```

## 진행 콜백 인터페이스

```python
# GUI 연결용 콜백
def progress_callback(event_type, data):
    # event_type: "start_site" | "post_found" | "error" | "complete"
    # data: {"agency": ..., "count": ..., "message": ...}
    pass
```

## 테스트 방법

```bash
python -c "
from modules.crawler import Crawler
c = Crawler('config/sites.xlsx', 'data/history.json')
results = c.run(site_filter=['대전광역시', '논산시'])
print(f'수집: {len(results)}건')
"
```
