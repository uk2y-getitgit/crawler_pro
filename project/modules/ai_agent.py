# -*- coding: utf-8 -*-
"""
ai_agent.py — 안전점검 모니터링 시스템

AI 에이전트 (Claude / Gemini 전환 가능).
  - BaseFilterAgent: 공통 인터페이스
  - ClaudeFilterAgent: Claude API (claude-haiku-4-5-20251001)
  - GeminiFilterAgent: Google Gemini API (gemini-2.0-flash)
  - create_post_filter_agent(provider): 팩토리 함수
  - keyword_fallback_filter: 두 API 모두 실패 시 키워드 기반 폴백
  - ErrorHandlerAgent: 크롤링 오류 원인 분석

AI_PROVIDER 설정 (.env):
  AI_PROVIDER=claude   → ClaudeFilterAgent 사용 (기본값)
  AI_PROVIDER=gemini   → GeminiFilterAgent 사용
  API 키 없으면 반대 provider 자동 시도, 둘 다 없으면 키워드 폴백.
"""

from __future__ import annotations

import abc
import json
import logging
import os
import sys
import time

# --------------------------------------------------------------------------- #
# .env 로드
# --------------------------------------------------------------------------- #
try:
    from dotenv import load_dotenv
    _dotenv_available = True
except ImportError:
    _dotenv_available = False

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

if _dotenv_available:
    _env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(_env_path):
        load_dotenv(_env_path)

# --------------------------------------------------------------------------- #
# 선택적 SDK
# --------------------------------------------------------------------------- #
try:
    import anthropic as _anthropic_sdk
    _anthropic_available = True
except ImportError:
    _anthropic_available = False

try:
    from google import genai as _genai
    from google.genai import types as _genai_types
    _gemini_available = True
except ImportError:
    _gemini_available = False

# --------------------------------------------------------------------------- #
# 로거
# --------------------------------------------------------------------------- #
try:
    from loguru import logger  # type: ignore
except ImportError:
    logger = logging.getLogger("ai_agent")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        _h = logging.StreamHandler()
        _h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(_h)

# --------------------------------------------------------------------------- #
# 상수
# --------------------------------------------------------------------------- #
CLAUDE_MODEL   = "claude-haiku-4-5-20251001"
GEMINI_MODEL   = "gemini-2.5-flash"
BATCH_SIZE     = 30
MAX_RETRY      = 3
KEYWORDS_PATH_DEFAULT = "config/keywords.json"

SYSTEM_PROMPT = """당신은 '안전점검 수행기관 지정공고' 탐지 전문가입니다.
아래 게시글 제목 목록에서 해당 공고에 해당하는 것만 선별하세요.

【수집 대상】
- '안전점검' + '수행기관' 또는 '대행기관' + '지정' 조합
- '시특법', '시설물안전법' 관련 기관 지정/모집 공고
- 안전진단 전문기관 지정 공고

【제외】
- 단순 공사 입찰공고 (안전점검 수행기관 선정이 아닌 경우)
- 입찰 결과, 낙찰 통보, 계약 체결 공고
- 안전 교육, 안전 캠페인 관련

【출력 형식】JSON만 출력, 설명 없음:
{"matched": [{"index": 0, "reason": "판단 근거 한 줄"}], "total_checked": N, "matched_count": N}"""


# --------------------------------------------------------------------------- #
# 내부 예외
# --------------------------------------------------------------------------- #
class _APIFatalError(Exception):
    """API 치명 오류 — 폴백으로 전환해야 함."""

class _BatchParseError(Exception):
    """배치 응답 JSON 파싱 실패 — 해당 배치 수동확인."""


# --------------------------------------------------------------------------- #
# 유틸
# --------------------------------------------------------------------------- #
def _is_valid_claude_key(key: str) -> bool:
    if not key:
        return False
    key = key.strip()
    if not key.startswith("sk-ant-"):
        return False
    markers = ("여기에", "입력", "your", "xxx", "placeholder", "...")
    low = key.lower()
    if any(m in key or m in low for m in markers):
        return False
    return len(key) >= 30

def _is_valid_gemini_key(key: str) -> bool:
    if not key:
        return False
    key = key.strip()
    markers = ("여기에", "입력", "your", "xxx", "placeholder", "...")
    low = key.lower()
    if any(m in key or m in low for m in markers):
        return False
    return len(key) >= 20

def _load_keywords(keywords_path: str | None = None) -> dict:
    path = keywords_path or os.path.join(BASE_DIR, KEYWORDS_PATH_DEFAULT)
    if not os.path.isabs(path):
        path = os.path.join(BASE_DIR, path)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "primary":   data.get("primary", []),
            "secondary": data.get("secondary", []),
            "exclude":   data.get("exclude", []),
        }
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        logger.warning(f"keywords.json 로드 실패: {e}")
        return {"primary": [], "secondary": [], "exclude": []}

@staticmethod
def _build_user_message(batch: list[dict]) -> str:
    lines = ["다음 게시글 목록을 분석해주세요:"]
    for i, post in enumerate(batch):
        title = post.get("title", "") or ""
        date  = post.get("date",  "") or ""
        lines.append(f"[{i}] 제목: {title} | 날짜: {date}")
    return "\n".join(lines)

def _parse_json_response(text: str) -> dict:
    if not text:
        raise _BatchParseError("빈 응답")
    cleaned = text.strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    try:
        c = cleaned
        if c.startswith("```"):
            c = c.split("```", 2)
            c = c[1] if len(c) >= 2 else cleaned
            if c.lower().startswith("json"):
                c = c[4:]
            c = c.strip()
        start, end = c.find("{"), c.rfind("}")
        if start != -1 and end > start:
            return json.loads(c[start:end + 1])
    except (json.JSONDecodeError, IndexError):
        pass
    raise _BatchParseError("JSON 파싱 실패(재시도 1회 후)")


# --------------------------------------------------------------------------- #
# 키워드 폴백 필터
# --------------------------------------------------------------------------- #
def keyword_fallback_filter(posts: list[dict], keywords_path: str | None = None) -> list[dict]:
    """
    키워드 기반 폴백 필터.
    - exclude 포함 → 제외
    - primary 1개 이상 → 수집 대상
    - secondary 2개 이상 + primary 없음 → 수집 대상
    """
    kw = _load_keywords(keywords_path)
    matched = []
    for post in posts:
        title = post.get("title", "") or ""
        if any(e and e in title for e in kw["exclude"]):
            continue
        hit_p = [p for p in kw["primary"]   if p and p in title]
        hit_s = [s for s in kw["secondary"] if s and s in title]
        if hit_p:
            post["ai_reason"] = "키워드 매칭: " + ", ".join(hit_p)
            matched.append(post)
        elif len(hit_s) >= 2:
            post["ai_reason"] = "키워드 매칭: " + ", ".join(hit_s)
            matched.append(post)
    return matched


# --------------------------------------------------------------------------- #
# BaseFilterAgent
# --------------------------------------------------------------------------- #
class BaseFilterAgent(abc.ABC):
    """게시글 필터 에이전트 공통 인터페이스."""

    provider_name: str = "base"

    def __init__(self, keywords_path: str | None = None,
                 progress_callback=None):
        self.keywords_path    = keywords_path
        self.progress_callback = progress_callback or (lambda et, data: None)
        self.use_fallback     = False
        self.fallback_reason  = ""

    def _emit(self, event_type: str, data: dict):
        try:
            self.progress_callback(event_type, data)
        except Exception as e:
            logger.warning(f"progress_callback 오류: {e}")

    @abc.abstractmethod
    def _call_api(self, user_msg: str) -> str:
        """API 호출 — 응답 텍스트 반환. 치명 오류 시 _APIFatalError."""

    def _filter_batch(self, batch: list[dict]) -> list[dict]:
        user_msg = _build_user_message(batch)
        text     = self._call_api(user_msg)
        data     = _parse_json_response(text)
        result   = []
        for entry in data.get("matched", []):
            idx = entry.get("index")
            if not isinstance(idx, int) or not (0 <= idx < len(batch)):
                continue
            post = batch[idx]
            reason = entry.get("reason", "")
            post["ai_reason"] = f"AI 판단({self.provider_name}): {reason}" if reason else f"AI 판단({self.provider_name}): 수집 대상"
            result.append(post)
        return result

    def filter(self, posts: list[dict]) -> list[dict]:
        if not posts:
            return []
        if self.use_fallback:
            self._emit("ai_fallback", {"reason": self.fallback_reason, "total": len(posts)})
            return keyword_fallback_filter(posts, self.keywords_path)

        matched_all = []
        total_batches = (len(posts) + BATCH_SIZE - 1) // BATCH_SIZE

        for b in range(total_batches):
            batch = posts[b * BATCH_SIZE:(b + 1) * BATCH_SIZE]
            self._emit("ai_batch_start", {"batch": b + 1, "total_batches": total_batches, "size": len(batch)})
            try:
                matched_all.extend(self._filter_batch(batch))
            except _APIFatalError as e:
                logger.warning(f"[{self.provider_name}] API 치명 오류, 키워드 폴백 전환: {e}")
                self.use_fallback    = True
                self.fallback_reason = str(e)
                self._emit("ai_fallback", {"reason": self.fallback_reason, "total": len(posts)})
                matched_all.extend(keyword_fallback_filter(posts[b * BATCH_SIZE:], self.keywords_path))
                return matched_all
            except _BatchParseError as e:
                # 파싱 실패 시 배치 전체를 포함하면(fail-open) 정밀도가 무너진다.
                # 정밀도 최우선 원칙에 따라, 해당 배치는 키워드 폴백으로만 선별한다.
                logger.warning(f"[{self.provider_name}] 배치 {b+1} 파싱 실패, 키워드 폴백 적용: {e}")
                fb = keyword_fallback_filter(batch, self.keywords_path)
                for post in fb:
                    post["ai_reason"] = (post.get("ai_reason") or "") + " (AI 파싱실패→키워드선별)"
                matched_all.extend(fb)
                self._emit("ai_batch_manual",
                           {"batch": b + 1, "size": len(batch), "kept": len(fb)})

        return matched_all


# --------------------------------------------------------------------------- #
# ClaudeFilterAgent
# --------------------------------------------------------------------------- #
class ClaudeFilterAgent(BaseFilterAgent):
    """Claude API (claude-haiku-4-5-20251001) 기반 필터."""

    provider_name = "claude"

    def __init__(self, api_key: str | None = None, **kwargs):
        super().__init__(**kwargs)
        self.api_key = api_key or os.getenv("ANTHROPIC_API_KEY", "")
        self.client  = None
        self._init_client()

    def _init_client(self):
        if not _anthropic_available:
            self.use_fallback    = True
            self.fallback_reason = "anthropic SDK 미설치"
            logger.warning(f"[ClaudeFilterAgent] {self.fallback_reason}")
            return
        if not _is_valid_claude_key(self.api_key):
            self.use_fallback    = True
            self.fallback_reason = "ANTHROPIC_API_KEY 없음 또는 미설정"
            logger.warning(f"[ClaudeFilterAgent] {self.fallback_reason}")
            return
        try:
            self.client = _anthropic_sdk.Anthropic(api_key=self.api_key)
        except Exception as e:
            self.use_fallback    = True
            self.fallback_reason = f"클라이언트 초기화 실패: {e}"
            logger.warning(f"[ClaudeFilterAgent] {self.fallback_reason}")

    def _call_api(self, user_msg: str) -> str:
        delay, last_err = 1.0, None
        for attempt in range(MAX_RETRY):
            try:
                resp = self.client.messages.create(
                    model=CLAUDE_MODEL,
                    max_tokens=2048,
                    system=SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_msg}],
                )
                for block in resp.content:
                    if getattr(block, "type", None) == "text":
                        return block.text
                return ""
            except _anthropic_sdk.RateLimitError as e:
                last_err = e
                logger.warning(f"  rate limit (재시도 {attempt+1}/{MAX_RETRY}), {delay:.0f}초 대기")
                if attempt < MAX_RETRY - 1:
                    time.sleep(delay); delay *= 2
            except (_anthropic_sdk.AuthenticationError,
                    _anthropic_sdk.PermissionDeniedError) as e:
                raise _APIFatalError(f"인증 실패: {e}") from e
            except _anthropic_sdk.APIStatusError as e:
                last_err = e
                if attempt < MAX_RETRY - 1:
                    time.sleep(delay); delay *= 2
                else:
                    raise _APIFatalError(f"API 오류: {e}") from e
            except _anthropic_sdk.APIConnectionError as e:
                last_err = e
                if attempt < MAX_RETRY - 1:
                    time.sleep(delay); delay *= 2
                else:
                    raise _APIFatalError(f"네트워크 오류: {e}") from e
        raise _APIFatalError(f"재시도 소진: {last_err}")


# --------------------------------------------------------------------------- #
# GeminiFilterAgent
# --------------------------------------------------------------------------- #
class GeminiFilterAgent(BaseFilterAgent):
    """Google Gemini API (gemini-2.0-flash) 기반 필터."""

    provider_name = "gemini"

    def __init__(self, api_key: str | None = None, **kwargs):
        super().__init__(**kwargs)
        self.api_key = api_key or os.getenv("GEMINI_API_KEY", "")
        self.model   = None
        self._init_client()

    def _init_client(self):
        if not _gemini_available:
            self.use_fallback    = True
            self.fallback_reason = "google-genai SDK 미설치"
            logger.warning(f"[GeminiFilterAgent] {self.fallback_reason}")
            return
        if not _is_valid_gemini_key(self.api_key):
            self.use_fallback    = True
            self.fallback_reason = "GEMINI_API_KEY 없음 또는 미설정"
            logger.warning(f"[GeminiFilterAgent] {self.fallback_reason}")
            return
        try:
            self.model = _genai.Client(api_key=self.api_key)
        except Exception as e:
            self.use_fallback    = True
            self.fallback_reason = f"클라이언트 초기화 실패: {e}"
            logger.warning(f"[GeminiFilterAgent] {self.fallback_reason}")

    def _call_api(self, user_msg: str) -> str:
        delay, last_err = 1.0, None
        for attempt in range(MAX_RETRY):
            try:
                # gemini-2.5-flash는 thinking이 기본 ON이라 출력 토큰을 소진해
                # JSON이 잘려 파싱 실패한다 → thinking 비활성 + JSON모드 + 토큰 확대.
                cfg_kwargs = dict(
                    system_instruction=SYSTEM_PROMPT,
                    max_output_tokens=8192,
                    response_mime_type="application/json",
                )
                if hasattr(_genai_types, "ThinkingConfig"):
                    cfg_kwargs["thinking_config"] = _genai_types.ThinkingConfig(
                        thinking_budget=0)
                resp = self.model.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=user_msg,
                    config=_genai_types.GenerateContentConfig(**cfg_kwargs),
                )
                return resp.text or ""
            except Exception as e:
                err_str = str(e).lower()
                if any(k in err_str for k in ("api_key_invalid", "permission", "unauthorized", "403")):
                    raise _APIFatalError(f"인증 실패: {e}") from e
                if any(k in err_str for k in ("quota", "rate", "429", "resource_exhausted")):
                    last_err = e
                    logger.warning(f"  rate limit (재시도 {attempt+1}/{MAX_RETRY}), {delay:.0f}초 대기")
                    if attempt < MAX_RETRY - 1:
                        time.sleep(delay); delay *= 2
                    continue
                last_err = e
                if attempt < MAX_RETRY - 1:
                    time.sleep(delay); delay *= 2
                else:
                    raise _APIFatalError(f"Gemini API 오류: {e}") from e
        raise _APIFatalError(f"재시도 소진: {last_err}")


# --------------------------------------------------------------------------- #
# 팩토리 함수
# --------------------------------------------------------------------------- #
def create_post_filter_agent(
    provider: str | None = None,
    keywords_path: str | None = None,
    progress_callback=None,
) -> BaseFilterAgent:
    """
    AI provider에 맞는 필터 에이전트를 생성한다.

    provider=None → .env의 AI_PROVIDER 값 사용 (기본값: claude)
    우선순위: 지정 provider → 반대 provider 자동 시도 → 키워드 폴백

    반환된 에이전트의 use_fallback=True이면 키워드 폴백으로 동작한다.
    """
    if provider is None:
        provider = os.getenv("AI_PROVIDER", "claude").lower().strip()

    kwargs = dict(keywords_path=keywords_path, progress_callback=progress_callback)

    # 1차: 지정 provider 시도
    if provider == "gemini":
        agent = GeminiFilterAgent(**kwargs)
        if not agent.use_fallback:
            logger.info("[AI] GeminiFilterAgent 활성화")
            return agent
        # Gemini 실패 → Claude 자동 시도
        logger.warning(f"[AI] Gemini 초기화 실패 ({agent.fallback_reason}), Claude 시도")
        fallback_agent = ClaudeFilterAgent(**kwargs)
        if not fallback_agent.use_fallback:
            logger.info("[AI] ClaudeFilterAgent 로 자동 전환")
            return fallback_agent
    else:
        # provider == "claude" 또는 기타
        agent = ClaudeFilterAgent(**kwargs)
        if not agent.use_fallback:
            logger.info("[AI] ClaudeFilterAgent 활성화")
            return agent
        # Claude 실패 → Gemini 자동 시도
        logger.warning(f"[AI] Claude 초기화 실패 ({agent.fallback_reason}), Gemini 시도")
        fallback_agent = GeminiFilterAgent(**kwargs)
        if not fallback_agent.use_fallback:
            logger.info("[AI] GeminiFilterAgent 로 자동 전환")
            return fallback_agent

    # 둘 다 실패 → 키워드 폴백 모드 에이전트 반환
    logger.warning("[AI] Claude·Gemini 모두 초기화 실패, 키워드 폴백 모드")
    agent.use_fallback    = True
    agent.fallback_reason = "Claude·Gemini API 키 모두 미설정 — 키워드 폴백 사용"
    return agent


# PostFilterAgent: 하위 호환 별칭 (기존 코드가 PostFilterAgent()를 직접 사용하는 경우)
class PostFilterAgent(ClaudeFilterAgent):
    """하위 호환 별칭. 새 코드는 create_post_filter_agent() 사용 권장."""
    pass


# --------------------------------------------------------------------------- #
# ErrorHandlerAgent
# --------------------------------------------------------------------------- #
class ErrorHandlerAgent:
    """크롤링 오류 텍스트를 받아 원인 분석 및 대처 방법을 한국어로 반환."""

    SYSTEM = (
        "당신은 웹 크롤링 오류 진단 전문가입니다. "
        "주어진 오류 메시지를 분석하여 (1) 추정 원인과 (2) 대처 방법을 "
        "한국어로 간결하게 제시하세요. 불필요한 서론 없이 핵심만 답하세요."
    )

    def __init__(self, provider: str | None = None):
        self._provider = (provider or os.getenv("AI_PROVIDER", "claude")).lower()
        self._client_claude  = None
        self._client_gemini  = None
        self._active = None   # "claude" | "gemini" | None

        claude_key = os.getenv("ANTHROPIC_API_KEY", "")
        gemini_key = os.getenv("GEMINI_API_KEY", "")

        if self._provider == "gemini" and _gemini_available and _is_valid_gemini_key(gemini_key):
            try:
                self._client_gemini = _genai.Client(api_key=gemini_key)
                self._active = "gemini"
            except Exception:
                pass
        if self._active is None and _anthropic_available and _is_valid_claude_key(claude_key):
            try:
                self._client_claude = _anthropic_sdk.Anthropic(api_key=claude_key)
                self._active = "claude"
            except Exception:
                pass
        if self._active is None and _gemini_available and _is_valid_gemini_key(gemini_key):
            try:
                self._client_gemini = _genai.Client(api_key=gemini_key)
                self._active = "gemini"
            except Exception:
                pass

    def analyze(self, error_text: str) -> str:
        if not error_text:
            return self._generic(error_text)
        user_msg = f"다음 크롤링 오류를 분석해주세요:\n{error_text}"
        try:
            if self._active == "claude":
                resp = self._client_claude.messages.create(
                    model=CLAUDE_MODEL, max_tokens=1024,
                    system=self.SYSTEM,
                    messages=[{"role": "user", "content": user_msg}],
                )
                for block in resp.content:
                    if getattr(block, "type", None) == "text":
                        return block.text.strip()
            elif self._active == "gemini":
                resp = self._client_gemini.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=user_msg,
                    config=_genai_types.GenerateContentConfig(
                        system_instruction=self.SYSTEM,
                        max_output_tokens=1024,
                    ),
                )
                return (resp.text or "").strip()
        except Exception as e:
            logger.warning(f"[ErrorHandlerAgent] API 실패: {e}")
        return self._generic(error_text)

    @staticmethod
    def _generic(error_text: str) -> str:
        et = (error_text or "").lower()
        if "timeout" in et or "timed out" in et:
            return "원인: 서버 응답 지연.\n대처: 잠시 후 재시도하거나 TIMEOUT 값을 늘리세요."
        if "ssl" in et:
            return "원인: SSL 인증서 문제.\n대처: verify=False 재시도가 이미 적용됩니다."
        if "404" in et or "not found" in et:
            return "원인: 게시판 URL 변경됨.\n대처: sites.xlsx의 URL을 최신 주소로 갱신하세요."
        if "403" in et or "forbidden" in et:
            return "원인: 접근 차단(봇 차단).\n대처: User-Agent 변경 또는 요청 간격을 늘리세요."
        if "0개" in (error_text or "") or "선택자" in (error_text or ""):
            return "원인: 페이지 구조 변경으로 CSS 선택자 미매칭.\n대처: HTML 구조를 확인하고 선택자를 보완하세요."
        return "원인: 일반 크롤링 오류.\n대처: 네트워크 상태와 사이트 접근 가능 여부를 확인 후 재시도하세요."
