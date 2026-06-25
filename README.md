# Carrier Format Studio

GitHub Pages용 정적 웹앱입니다. Raw order spreadsheet를 업로드한 뒤 주문별 운송사를 선택하고 FEDEX, DHL, 우체국 K-packet, ACI 업로드용 파일을 생성합니다.

## 현재 구현 범위

- Raw data `.xlsx`, `.xls`, `.csv` 업로드
- `Order id` 기준 주문 그룹핑
- 주문별 운송사 드롭다운 선택
- 수출자 기본 정보 수정 및 `localStorage` 저장
- DHL CSV 생성
- 우체국 K-packet XLSX 생성
- ACI US XLSX 생성
- FEDEX `.in` 텍스트 파일 생성
- 여러 결과물을 zip으로 다운로드

## 실행

정적 파일이므로 `index.html`을 브라우저에서 열 수 있습니다.

샘플 파일 로드 버튼까지 테스트하려면 로컬 서버로 여는 것이 좋습니다.

```bash
python3 -m http.server 4173
```

그 다음 브라우저에서 아래 주소를 엽니다.

```text
http://127.0.0.1:4173/index.html
```

## 외부 라이브러리

현재 `index.html`은 CDN에서 아래 라이브러리를 불러옵니다.

- SheetJS: Excel/CSV read/write
- JSZip: 결과 파일 zip 묶기

GitHub Pages 배포 시에는 인터넷 연결이 필요합니다. 완전 오프라인 배포가 필요하면 라이브러리 파일을 프로젝트에 vendoring하는 Phase를 추가해야 합니다.

## 분석 문서

- [Phase 1 Analysis](docs/phase-1-analysis.md)
- [구현계획서](구현계획서.md)

