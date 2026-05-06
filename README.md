# 검사성적서 자동 생성 — 새한나노텍 / 에스앤테크

홀가공 측정 데이터(.xls A/B)를 ZIP 으로 업로드하면 자동으로 통계 계산 + 검사성적서 PDF 생성.

## 거래처
- 에이텍솔루션 (영문 라벨, 동심도 8샘플, 위치도 <0.10)
- 솔믹스 (한글 라벨, 동심도 7샘플, 위치도 <0.30)

## 기능
- ZIP 업로드 → 폴더별 A.xls + B.xls 자동 합산
- 통계 자동 계산 (Size / Roundness / Position / Concentricity)
- 거래처별 SPEC 자동 적용 + 부적합 화면 강조
- PDF 미리보기 + 합본 PDF 다운로드 + 엑셀 모음 다운로드

## 기술
- Python + Streamlit
- xlsx → PDF 변환: LibreOffice
- PDF 미리보기: PyMuPDF
