# 내부 업무 자동화 사이트

Streamlit 기반 내부 업무 자동화 사이트입니다. 현재는 CPC 광고 운영보고서 생성 기능을 중심으로 구성되어 있으며, 계약서 자동 생성과 업무 매뉴얼 메뉴는 추후 기능 확장을 위한 준비 화면으로 만들어져 있습니다.

## 파일 구성

- `app.py`: 사이트 실행 파일
- `contract_auto.py`: 최신 계약서 자동 생성 모듈
- `contract_templates_embedded.py`: 배포 시 Excel 템플릿 파일이 누락되어도 동작하도록 내장한 계약서 템플릿
- `internal_automation_site.py`: 같은 내용의 백업 실행 파일
- `requirements.txt`: 실행에 필요한 Python 패키지 목록
- `packages.txt`: Streamlit Cloud에서 OCR 엔진을 설치하기 위한 시스템 패키지 목록
- `README.md`: 실행 방법과 기능 설명
- `assets/fonts/`: 기존 CPC 보고서 출력과 같은 Pretendard PDF 폰트입니다. 업로드 용량을 줄이기 위해 실제 PDF 생성에 필요한 Medium, Bold, ExtraBold 3개만 포함합니다.
- `assets/logo.png`: 보고서 로고 표시 옵션에서 사용하는 기본 로고
- `assets/templates/`: 계약서 자동 생성용 Excel 템플릿

## 설치

Python 환경에서 필요한 패키지를 설치합니다.

```bash
pip install -r /Users/choeinhye/Desktop/requirements.txt
```

## 실행

```bash
cd /Users/choeinhye/Desktop/internal_automation_site
streamlit run app.py
```

## 현재 메뉴

- CPC 보고서 생성
- 계약서 자동 생성
- 업무 매뉴얼
  - CPC 등록 가이드
  - CPT/CPC 권한 이전 가이드
  - 백오피스 데이터 추출 가이드
  - 백오피스 CPC 등록 방법
  - 리뷰 삭제 요청 가이드
  - 자주 묻는 질문(FAQ)
  - 검색 기능

## CPC 보고서 생성

중국어 CPC 운영 데이터 파일을 업로드하면 점주 전달용 PDF와 PNG 보고서를 생성합니다. 당월 파일과 전월 파일을 여러 개 함께 업로드할 수 있으며, 날짜 데이터를 기준으로 당월/전월 비교 데이터를 자동 인식합니다.

지원 파일 형식:

- `.xlsx`
- `.xls`
- `.csv`

주요 기능:

- CPC 데이터 컬럼 자동 인식
- 여러 CPC 파일 동시 업로드
- 가장 최근 월 기준 보고서 생성
- 직전 월 데이터가 포함된 경우 전월 비교 자동 반영
- 보고서 로고 이미지 선택 삽입
- PDF 다운로드
- PNG 이미지 다운로드

## 계약서 자동 생성

1年/3个月 계약서 템플릿에 입력값을 반영해 Excel/PDF/ZIP 파일을 생성합니다.

주요 기능:

- 사업자등록증 이미지 OCR 보조 입력
- 한국어 주소 영문 주소 자동 변환 보조
- 서명/도장 이미지 삽입
- Excel 다운로드
- PDF 다운로드
- Excel + PDF ZIP 다운로드

## 확장 방법

새 메뉴를 추가하려면 `app.py`의 `NAV_ITEMS`에 항목을 추가하고, `render_page()`에서 해당 화면 렌더링 함수를 연결하면 됩니다.
