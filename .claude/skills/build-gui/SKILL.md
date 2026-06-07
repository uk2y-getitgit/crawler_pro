---
name: build-gui
description: "안전점검 모니터링 시스템 PyQt6 GUI 개발. UI 코드 작성 전 반드시 4가지 스타일 시안을 먼저 제시하고 사용자 선택 후 코드 작성. 'GUI 만들어줘', 'UI 개발', 'PyQt6', '화면 만들어줘', '인터페이스' 요청 시 사용. EXE 배포 전제로 BASE_DIR 경로 패턴 필수 적용."
---

# Build GUI

PyQt6 기반 데스크톱 UI를 개발한다. **코드 작성 전 반드시 4가지 스타일 시안을 먼저 보여주고 선택을 받는다.**

## 필수 절차

1. 사용자에게 4가지 스타일 시안 제시 (SVG 또는 HTML 렌더링)
2. 각 스타일의 버튼 기능 설명 제공
3. 사용자 선택 + 수정사항 수령
4. 선택된 스타일로 실제 코드 작성 시작

## 4가지 스타일 정의

| 스타일 | 배경 | 강조 | 특징 |
|--------|------|------|------|
| 다크 대시보드 | 다크네이비 | 하늘색 | 야간 작업, 전문 관제 느낌 |
| 라이트 비즈니스 | 밝은 회백색 | 네이비 | 공문서, 업무용, 상사에게 보여주기 좋음 |
| 딥블루 모던 | 딥네이비 | 레드-오렌지 | 임팩트 있는 모던, 신규 공고 강조 |
| 화이트 클린 | 순백 | 파랑 | 가장 단순, 처음 쓰는 사람도 직관적 |

## 레이아웃 구조

```
┌─────────────────────────────────────────────────────┐
│  [크롤링 실행] [선택 실행] [중지] [설정]  ← 상단 버튼바  │
├───────────────┬─────────────────────────────────────┤
│  사이트 목록  │  수집결과 (신규: 노란 배경)          │
│  (체크박스)   │  [신규공고 보기] [엑셀 저장]         │
│  대전광역시 ☑ │  🆕 안전점검 수행기관 지정..        │
│  논산시     ☑ │  - 대전광역시 | 2026-06-01 | ...     │
│  ...          │                                     │
├───────────────┴─────────────────────────────────────┤
│  실행 로그: [2026-06-05 08:01] 대전광역시 수집 중...  │
└─────────────────────────────────────────────────────┘
```

## 파일 구조

```
main.py                 ← 앱 진입점 (BASE_DIR 설정 포함)
modules/
├── gui_main.py         ← 메인 윈도우
├── gui_settings.py     ← 설정 화면
└── crawler_thread.py   ← QThread 크롤러 래퍼
```

## QThread 크롤러 연결

```python
class CrawlerThread(QThread):
    progress = pyqtSignal(str, dict)   # event_type, data
    finished = pyqtSignal(list)        # results

    def run(self):
        crawler = Crawler(
            sites_path=...,
            history_path=...,
            progress_callback=lambda e, d: self.progress.emit(e, d)
        )
        results = crawler.run(site_filter=self.selected_sites)
        self.finished.emit(results)
```

## 신규 공고 강조

```python
# 결과 테이블에서 신규 항목 노란 배경
if item.is_new:
    cell.setBackground(QColor('#FFFF99'))
    cell.setText(f"🆕 {item.title}")
```

## EXE 호환 규칙

- `main.py` 상단에 BASE_DIR 패턴 반드시 포함
- 아이콘 파일은 `--add-data`로 패키징
- QApplication 초기화 전 high-DPI 설정 추가

## 설정 화면 기능

- **사이트 관리**: sites.xlsx 행 추가/삭제/활성화 토글
- **키워드 설정**: keywords.json primary/secondary/exclude 편집
- **스케줄 설정**: 자동 실행 시간 (HH:MM), 요일 선택

## 테스트

PyQt6 설치 후 `python main.py` 실행 → 메인 윈도우 표시 확인
