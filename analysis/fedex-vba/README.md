# FEDEX VBA Extraction Report

작성일: 2026-06-24

## 추출 결과

`sample/fedex sample.xlsm`에서 VBA 추출을 완료했다.

설치한 도구:

- `oletools`
- 설치 위치: `.tools/py`

추출 산출물:

- 전체 VBA 덤프: `analysis/fedex-vba/all-vba.txt`
- 모듈별 파일: `analysis/fedex-vba/modules/`
- 핵심 모듈: `analysis/fedex-vba/modules/modHKCT.bas`

추출된 모듈 수: 15개

## 핵심 매크로

`.in` 파일 생성 로직은 `modHKCT.bas`의 `Process_Click()`에 있다.

관련 함수:

| 이름 | 역할 |
| --- | --- |
| `Process_Click()` | `Trans` 시트를 읽어 `.in` 파일 생성 |
| `FileHandler()` | 저장 파일명/경로 선택. 기본 파일명은 `YYYYMMDD_hhmmss.in` |
| `NextAWB()` | Tracking number 자동 배정 |
| `CountryCode()` | Country 시트에서 국가명으로 국가코드 조회 |
| `Tracking_Click()` | `Trans` 시트의 field `1123` 컬럼에 tracking number 입력 |

## `.in` 생성 규칙

VBA 기준 생성 흐름:

1. `Trans` 시트를 활성화한다.
2. `Trans.UsedRange`의 행/열 수를 구한다.
3. 데이터는 4행부터 마지막 행까지 순회한다.
4. A열, 즉 `Transaction ID`가 비어 있으면 해당 행은 사실상 출력 대상이 아니다.
5. 현재 행의 A열 값이 이전 행의 A열 값과 다르면 새 shipment로 판단한다.
6. 새 shipment 시작 시 `0,"020"`을 출력한다.
7. `Default` 시트의 C열 값이 비어 있지 않은 행을 모두 출력한다.
   - field number: `Default` 시트 B열
   - value: `Default` 시트 C열
8. `Trans` 시트의 모든 컬럼을 순회한다.
   - field number: 3행
   - multiplier: 2행
   - value: 현재 데이터 행
9. field number와 value가 모두 비어 있지 않으면 출력한다.
10. 2행 multiplier가 있고 value가 숫자이면 `value * multiplier`를 출력한다.
11. 같은 Transaction ID가 연속으로 나오면 추가 commodity로 판단한다.
12. 추가 commodity 행에서는 `GroupFields` 시트 A열에 존재하는 field만 출력한다.
13. 추가 commodity field number는 `fieldNumber-2`, `fieldNumber-3` 형식으로 출력한다.
14. shipment의 마지막 행이면 `99,""`를 출력한다.

## 현재 웹앱 FEDEX 구현과의 차이

초기 `index.html`의 `generateFedexIn()`은 샘플 기반 1차 구현이었고, VBA와 아래 차이가 있었다.

| 항목 | 현재 구현 | VBA 실제 규칙 |
| --- | --- | --- |
| 기본값 | 코드에 일부 field를 하드코딩 | `Default` 시트 B/C열 전체를 순회 |
| Trans 컬럼 출력 | 코드에 일부 field를 하드코딩 | `Trans` 시트 3행 field number가 있는 모든 non-empty 셀 출력 |
| multiplier | 미반영 또는 일부 값 직접 계산 | `Trans` 시트 2행 값이 있으면 숫자 value에 곱함 |
| 다품목 처리 | 첫 품목 중심, 일부 합산 | 같은 Transaction ID의 추가 행은 `GroupFields` 등록 field만 `field-n`으로 출력 |
| 종료 field | `99,""`를 각 주문 끝에 넣는 의도는 있음 | 현재 행 A열이 다음 행 A열과 다를 때만 출력 |
| 파일명 | `FEDEX_YYYY-MM-DD.in` | VBA 기본값은 `YYYYMMDD_hhmmss.in` |

## 구현 권장안

FEDEX 생성 로직은 하드코딩 방식보다 아래 구조가 맞다.

1. FEDEX 템플릿 구조를 설정 데이터로 이식한다.
2. `Default` 시트 B/C열을 그대로 설정화한다.
3. `Trans` 시트 1~3행을 field schema로 사용한다.
4. raw order를 `Trans` 데이터 행 형태로 먼저 만든다.
5. VBA `Process_Click()` 규칙을 TypeScript/JavaScript로 그대로 재현한다.
6. 샘플 `fedex sample.in`과 byte/text diff로 검증한다.

## 적용 및 검증

2026-06-24에 `index.html`의 `generateFedexIn()`을 VBA 규칙 기반으로 교체했다.

적용 내용:

- `FEDEX_DEFAULT_FIELDS`: `Default` 시트 B/C열 출력 규칙 이식
- `FEDEX_TRANS_FIELDS`: `020 Transaction` 3행 field number 이식
- `FEDEX_TRANS_MULTIPLIERS`: `020 Transaction` 2행 multiplier 이식
- `FEDEX_GROUP_FIELDS`: `Grouping Fields` 시트 A열 이식
- `buildFedexTransRows()`: 주문/품목을 VBA의 `Trans` 데이터 행 형태로 변환
- `generateFedexIn()`: VBA `Process_Click()`의 새 shipment/추가 commodity/종료 field 출력 규칙 재현

검증 결과:

| 검증 | 결과 |
| --- | --- |
| `node --check` on extracted `index.html` script | 통과 |
| Python으로 VBA 알고리즘을 `fedex sample.xlsm`에 적용 후 `fedex sample.in`과 비교 | 262/262 lines exact match |
| `index.html`의 `generateFedexIn()`을 Node VM에서 호출해 샘플 주문 `142774` 다품목 구간과 비교 | 55/55 lines exact match |
