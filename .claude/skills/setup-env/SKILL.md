---
name: setup-env
description: "안전점검 모니터링 시스템 개발 환경 구성. Python 가상환경 생성, requirements.txt 설치, .env 템플릿 생성, sites.xlsx 재구성, keywords.json 생성, 프로젝트 폴더 구조 생성을 수행한다. '환경 설정', '초기 설정', 'Phase 1', 'setup' 요청 시 사용."
---

# Setup Environment

안전점검 모니터링 시스템의 개발 환경을 구성한다.

## 실행 전 확인

`site_list(전체).xlsx`가 프로젝트 루트에 있는지 확인한다. 없으면 중단하고 위치를 요청한다.

## 생성할 파일 목록

### requirements.txt

```
requests==2.31.0
playwright==1.44.0
beautifulsoup4==4.12.3
lxml==5.2.0
openpyxl==3.1.4
anthropic==0.28.0
python-dotenv==1.0.1
schedule==1.2.2
loguru==0.7.2
PyQt6==6.7.0
pyinstaller==6.6.0
```

### .env 템플릿

```
ANTHROPIC_API_KEY=sk-ant-여기에_API_키_입력
CRAWL_DELAY=2
MAX_PAGES=3
LOG_LEVEL=INFO
```

### config/keywords.json

```json
{
  "primary": [
    "안전점검 수행기관 지정",
    "수행기관 지정공고",
    "안전점검기관 지정"
  ],
  "secondary": [
    "안전점검 대행기관",
    "안전진단 수행기관",
    "정기안전점검 지정",
    "시특법 수행기관"
  ],
  "exclude": [
    "입찰결과", "낙찰", "계약체결", "취소"
  ]
}
```

### config/sites.xlsx 컬럼 구조

기존 site_list(전체).xlsx를 읽어 아래 컬럼으로 재구성한다:

| 컬럼 | 설명 |
|------|------|
| 기관명 | 기관 고유명 |
| 사이트타입 | 일반/전자조달/나라장터 |
| 게시판명 | 게시판 이름 (초기: 미등록) |
| 게시판URL | 게시판 직접 URL (초기: 공백) |
| 페이지파라미터 | 다음 페이지 이동 방식 |
| 활성화 | Y/N |
| 비고 | 특이사항 |

### check_env.py

실행하면 설치된 라이브러리 버전과 API 키 존재 여부를 출력하는 검증 스크립트.

### install.bat

```batch
pip install playwright
playwright install chromium
echo 설치 완료!
pause
```

## BASE_DIR 패턴

모든 파일 경로는 이 패턴을 사용한다:

```python
import sys, os
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
```

## 완료 확인

`python check_env.py` 실행 결과가 오류 없이 출력되면 Phase 1 완료.
