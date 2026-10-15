# 🌊 시추주상도 기반 토질 액상화 예비평가 프로그램

시추주상도(PDF, 이미지) 및 시추조사 입력 자료(Excel, CSV)를 분석하여 **지반의 토질 액상화(Liquefaction) 발생 가능성 및 본평가 대상 여부를 자동으로 판정하는 GUI 프로그램**입니다.

---

## 📌 주요 기능 (Features)

1. **시추주상도 이미지 및 PDF 자동 분석 / 미리보기**
   - PDF/이미지 형태의 시추주상도를 첨부하고 마우스 휠 및 드래그로 손쉽게 확대/축소하여 확인.
   - `PyMuPDF` 및 `Tesseract OCR` 기반으로 텍스트, 지하수위, 시추번호, 좌표, 표고, SPT-N값 자동 추출 지원.

2. **자동 의존성 안내 및 원클릭 설치 지원**
   - 필수 패키지(`PyMuPDF`, `Pillow`, `pytesseract`, `openpyxl`) 미설치 시 프로그램 내부에서 클릭 한 번으로 자동 설치 진행.
   - 설치 없이도 CSV 불러오기 및 수기 입력으로 예비평가 수행 가능.

3. **지층별/심도별 액상화 예비평가 자동화**
   - **평가 기준**:
     - 연중 최고지하수위 하부 위치 여부
     - 지반 심도 20m 이내 여부
     - 표준관입시험 타격횟수 N <= 25 여부
     - 사질토(모래, 실트질 등) 지반 여부 (암반 및 풍화암은 제외)
   - 굴진심도 및 N값 측정 개수에 맞춰 측점 간격(1.0m 또는 1.5m) 자동 산정 및 관입불가(50/5, 1/50 등) 시 기반암 자동 판정.
   - 지하수위 등 일부 정보가 누락된 경우에도 **부분자료 예비평가** 모드로 유연한 평가 진행.

4. **평가 기록 관리 및 저장/비교 기능**
   - 수행된 모든 평가 이력을 자동 기록 저장 (`liquefaction_history.json`).
   - 서로 다른 두 시추주상도의 평가 결과를 지층/심도별로 정밀 비교 및 다중 탭 Excel 파일 저장 지원.
   - 결과를 클립보드로 복사하여 Excel 등에 즉시 붙여넣기 기능 제공.

---

## 🛠️ 기술 스택 (Tech Stack)

- **Language**: Python 3.10+
- **GUI Framework**: `tkinter`, `ttk`
- **Data & Excel Processing**: `openpyxl`, `csv`, `json`, `re`
- **Document & Image Parsing (OCR)**: `PyMuPDF` (`fitz`), `Pillow` (`PIL`), `pytesseract`

---

## 🚀 시작하기 (Getting Started)

### 1. 필수 패키지 설치

터미널(CMD/PowerShell)에서 아래 명령어를 실행하여 필요한 라이브러리를 설치합니다:

```bash
pip install pymupdf pillow pytesseract openpyxl

python "코딩과 응용 프로그램/2026.10.15.py"

.
├── 코딩과 응용 프로그램/
│   └── 2026.10.15.py           # 액상화 예비평가 메인 GUI 프로그램
├── README.md                   # 프로젝트 안내 문서
└── liquefaction_history.json   # 평가 결과 저장/이력 관리 파일 (자동 생성)

```bash
pip install pymupdf pillow pytesseract openpyxl
