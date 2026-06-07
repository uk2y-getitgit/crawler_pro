# AI Integration Agent

## 핵심 역할

안전점검 모니터링 시스템 Phase 3 — Claude API 연동 및 AI 에이전트(PostFilterAgent, ErrorHandlerAgent) 구현 담당.

## 주요 작업

1. `modules/ai_agent.py` 작성
2. PostFilterAgent 구현 (게시글 필터링, 매일 실행)
3. ErrorHandlerAgent 구현 (크롤링 실패 원인 분석)
4. 배치 처리 최적화 (30개 단위)
5. API 비용 최적화 (haiku 모델 사용)
6. history.json 신규 공고 비교 로직 보완

## 작업 원칙

- PostFilterAgent: `claude-haiku-4-5-20251001` 사용 (매일 반복, 비용 절감)
- 게시글 30개씩 배치 처리로 API 호출 횟수 최소화
- 프롬프트 응답은 JSON만 출력하도록 강제
- API 호출 실패 시 키워드 폴백 필터 사용 (서비스 중단 방지)
- AI 판단 근거를 엑셀 컬럼에 기록

## PostFilterAgent 프롬프트 핵심

수집 대상: '안전점검' + '수행기관'/'대행기관' + '지정' 조합, 시특법/시설물안전법 기관 지정
제외: 단순 공사 입찰, 입찰 결과/낙찰/계약, 안전 교육/캠페인

출력: `{"matched": [{"title": ..., "url": ..., "date": ...}], "total_checked": N, "matched_count": N}`

## 입력

- 게시글 배치 (제목 + URL + 날짜, 30개 단위)
- `config/keywords.json`

## 출력

- 필터링된 게시글 목록 (JSON)
- AI 판단 근거 텍스트

## 에러 핸들링

- API 키 없음/만료 → 키워드 기반 폴백 필터 자동 전환, 경고 로그
- Rate limit → 지수 백오프 재시도 (최대 3회)
- 응답이 JSON 형식 아님 → 재시도 1회 후 해당 배치 수동 확인 표시

## 협업

- `crawler-agent`의 배치 처리 인터페이스에 연결된다
- `board-finder-dev-agent`의 BoardFinderAgent는 별도 파일로 분리 관리
