import sys, os

# BASE_DIR 패턴 (EXE 호환)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

print(f"BASE_DIR: {BASE_DIR}")
print()

# 라이브러리 확인
print("[ 라이브러리 ]")
libs = {
    'requests': 'requests',
    'playwright': 'playwright',
    'bs4': 'beautifulsoup4',
    'lxml': 'lxml',
    'openpyxl': 'openpyxl',
    'anthropic': 'anthropic',
    'google.genai': 'google-genai',
    'dotenv': 'python-dotenv',
    'schedule': 'schedule',
    'loguru': 'loguru',
    'PyQt6': 'PyQt6',
}
for lib, pkg in libs.items():
    try:
        __import__(lib)
        print(f"  ✓ {pkg}")
    except ImportError:
        print(f"  ✗ {pkg} — pip install {pkg}")

# .env 확인
print()
print("[ AI 설정 ]")
env_path = os.path.join(BASE_DIR, '.env')
if not os.path.exists(env_path):
    print("  ✗ .env 파일 없음")
else:
    print("  ✓ .env 파일 존재")
    try:
        from dotenv import dotenv_values
        env = dotenv_values(env_path)
    except ImportError:
        env = {}
        with open(env_path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if '=' in line and not line.startswith('#'):
                    k, _, v = line.partition('=')
                    env[k.strip()] = v.strip()

    provider = env.get('AI_PROVIDER', 'claude')
    print(f"  AI_PROVIDER = {provider}")

    claude_key = env.get('ANTHROPIC_API_KEY', '')
    if claude_key and claude_key.startswith('sk-ant-') and '여기에' not in claude_key and len(claude_key) >= 30:
        print("  ✓ ANTHROPIC_API_KEY 설정됨")
    else:
        print("  ✗ ANTHROPIC_API_KEY 미설정 — .env에 실제 키 입력 필요")

    gemini_key = env.get('GEMINI_API_KEY', '')
    if gemini_key and '여기에' not in gemini_key and len(gemini_key) >= 20:
        print("  ✓ GEMINI_API_KEY 설정됨")
    else:
        print("  ✗ GEMINI_API_KEY 미설정 — .env에 실제 키 입력 필요 (선택사항)")

    # 실제 사용될 provider 예측
    print()
    if provider == 'gemini':
        if gemini_key and '여기에' not in gemini_key:
            print("  → Gemini API 사용 예정")
        elif claude_key and '여기에' not in claude_key:
            print("  → Gemini 키 없음, Claude API 로 자동 전환 예정")
        else:
            print("  → API 키 없음, 키워드 폴백 모드로 실행됩니다")
    else:
        if claude_key and '여기에' not in claude_key:
            print("  → Claude API 사용 예정")
        elif gemini_key and '여기에' not in gemini_key:
            print("  → Claude 키 없음, Gemini API 로 자동 전환 예정")
        else:
            print("  → API 키 없음, 키워드 폴백 모드로 실행됩니다")

# config 파일 확인
print()
print("[ 설정 파일 ]")
for fname in ['config/sites.xlsx', 'config/keywords.json']:
    fpath = os.path.join(BASE_DIR, fname)
    if os.path.exists(fpath):
        print(f"  ✓ {fname}")
    else:
        print(f"  ✗ {fname} 없음")

print()
print("환경 확인 완료.")
