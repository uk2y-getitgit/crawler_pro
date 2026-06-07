---
name: build-ai-agents
description: "안전점검 모니터링 시스템 AI 에이전트(PostFilterAgent, BoardFinderAgent, ErrorHandlerAgent) 개발. Claude API 연동, 배치 처리, 키워드 폴백 필터 구현. 'Phase 3', 'AI 에이전트 만들어줘', 'PostFilterAgent', 'BoardFinderAgent', 'ai_agent.py 개발' 요청 시 사용."
---

# Build AI Agents

Claude API를 활용한 게시글 필터링 및 게시판 탐색 에이전트를 개발한다.

## 파일 구조

```
modules/
├── ai_agent.py         ← PostFilterAgent + ErrorHandlerAgent
└── board_finder.py     ← BoardFinderAgent (별도 파일)
```

## PostFilterAgent

매일 반복 실행 → 비용 절감 최우선: `claude-haiku-4-5-20251001` 사용

```python
class PostFilterAgent:
    MODEL = "claude-haiku-4-5-20251001"
    BATCH_SIZE = 30

    def filter(self, posts: list[dict]) -> list[dict]:
        # posts: [{"title": ..., "url": ..., "date": ...}]
        # 30개씩 배치 처리
        # 반환: matched posts with "ai_reason" 필드 추가
        ...
```

### 프롬프트

```
SYSTEM: 당신은 '안전점검 수행기관 지정공고' 탐지 전문가입니다.

수집 대상: '안전점검'+'수행기관'/'대행기관'+'지정' 조합,
시특법/시설물안전법 관련 기관 지정/모집 공고, 안전진단 전문기관 지정

제외: 단순 공사 입찰, 입찰결과/낙찰/계약체결 공고, 안전 교육/캠페인

출력 형식 (JSON만, 설명 없음):
{"matched": [{"title": "...", "url": "...", "date": "...", "reason": "..."}],
 "total_checked": N, "matched_count": N}
```

## BoardFinderAgent

초기 1회성 실행 → 정확도 최우선: `claude-sonnet-4-20250514` 사용

```python
class BoardFinderAgent:
    MODEL = "claude-sonnet-4-20250514"

    def find_boards(self, links: list[dict]) -> dict:
        # links: [{"text": "메뉴명", "url": "..."}]
        # 반환: {"selected": [{"name": ..., "url": ...}], "reason": "..."}
        ...
```

### 프롬프트

```
SYSTEM: 대한민국 관공서 웹사이트 게시판 분류 전문가.
메뉴 링크 목록에서 공고/고시/입찰 관련 게시판만 선별.

포함: 공고/고시/입찰공고/공고문/고시공고/지정공고/조달/구매/계약
제외: 민원/상담/건의/자유게시판/소개/기관안내/조직도/보도자료/
      뉴스/갤러리/사진/채용/인사/복지

출력 (JSON만): {"selected": [{"name": "게시판명", "url": "https://..."}], "reason": "선택 이유"}
```

## 키워드 폴백 필터

API 호출 실패 시 자동 전환되는 규칙 기반 필터:

```python
def keyword_fallback_filter(posts, keywords_path):
    # primary 키워드 1개 이상 포함 → 수집 대상
    # exclude 키워드 포함 → 제외
    # secondary는 보조 판단 (primary 없을 때 2개 이상이면 포함)
    ...
```

## 에러 처리

| 상황 | 처리 |
|------|------|
| API 키 없음/만료 | 키워드 폴백 전환 + 경고 로그 |
| Rate limit | 지수 백오프 (1s → 2s → 4s, 최대 3회) |
| JSON 파싱 실패 | 재시도 1회 후 해당 배치 수동확인 표시 |
| 모델 없음 | 오류 메시지 출력 후 중단 |

## 비용 추정

- 하루 9,000개 제목 처리 (100게시판 × 3페이지 × 30개)
- haiku 모델 기준: 약 300 배치 호출
- 예상 일일 비용: $0.05~0.15 수준

## 테스트 방법

```bash
python -c "
from modules.ai_agent import PostFilterAgent
agent = PostFilterAgent()
test_posts = [
    {'title': '안전점검 수행기관 지정공고', 'url': 'https://test.kr/1', 'date': '2026-06-01'},
    {'title': '도로 공사 입찰공고', 'url': 'https://test.kr/2', 'date': '2026-06-01'},
]
result = agent.filter(test_posts)
print(result)
"
```
