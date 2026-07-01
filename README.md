# NEO Expense Approval Form

경비 승인서(지출결의서) 작성 웹앱. 영수증 이미지/PDF를 업로드하면 OCR로 금액·날짜를 읽어 엑셀 양식으로 정리한다.

## 기능
- 영수증 이미지/PDF 업로드
- Windows 내장 OCR(`ocr_windows.ps1`)로 금액·날짜 추출
- openpyxl 기반 엑셀 경비 승인서 생성

## 요구사항
- Python 3.10+
- Windows (내장 OCR 엔진 사용)

## 설치
```bash
pip install -r requirements.txt
```

## 실행
```bash
python app.py
```
브라우저에서 `http://localhost:5000` 접속.

## 구조
- `app.py` — Flask 앱 (메인)
- `expense.py` — 경비 파싱 로직
- `ocr_windows.ps1` — Windows OCR 스크립트
- `CI/` — 로고 자산
