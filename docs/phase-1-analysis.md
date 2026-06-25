# Phase 1 Analysis: Carrier Spreadsheet Conversion

작성일: 2026-06-24

## 확인한 샘플

| 파일 | 역할 | 분석 결과 |
| --- | --- | --- |
| `sample/raw data sample.xlsx` | 입력 raw data | `Products` 단일 시트. `Order id` 기준 그룹핑, 다품목 주문은 같은 주문번호의 여러 행으로 표현됨. |
| `sample/dhl sample.csv` | DHL 출력 샘플 | 단일 CSV. 헤더 순서 유지 후 주문/품목 데이터를 채우는 방식이 적합함. |
| `sample/aci sample_US.xlsx` | ACI US 출력 샘플 | `Form`, `국가명_영문명`, `참고사항` 3개 시트. 값 입력은 `Form` 기준. |
| `sample/fedex sample.xlsm` | FEDEX 원본 템플릿 | 13개 시트와 `vbaProject.bin` 포함. 핵심 시트는 `020 Transaction`, `기본값(필수 입력)`, `Field#`, `Grouping Fields`. |
| `sample/fedex sample.in` | FEDEX 출력 샘플 | CSV text. `fieldNumber,"value"` 행 반복 구조이며 주문별 레코드가 이어짐. |
| `sample/code.rtf`, `sample/index.html` | 기존 Apps Script 구현 | 우체국 K-packet 엑셀 생성 규칙은 `index.html`의 `buildKoreaPostRow`, `downloadKoreaPostExcel`에 존재. |

## 표준 주문 모델

앱 내부에서는 raw data를 아래 모델로 정규화한다.

```ts
type Order = {
  orderId: string;
  orderNo: string;
  selectedCarrier: "FEDEX" | "DHL" | "KOREA_POST" | "ACI";
  receiver: {
    name: string;
    phone: string;
    email: string;
    country: string;
    state: string;
    city: string;
    postalCode: string;
    address1: string;
    address2: string;
    taxId: string;
  };
  package: {
    weightKg: number;
    length: number;
    width: number;
    height: number;
    boxType: string;
  };
  items: Array<{
    description: string;
    quantity: number;
    price: number;
    hsCode: string;
    originCountry: string;
  }>;
};
```

## Raw Data 컬럼 매핑

샘플 기준 고정 컬럼:

| Raw 컬럼 | 내부 필드 |
| --- | --- |
| `Order id` | `orderId` |
| `Order no` | `orderNo` |
| `Product Name(상품명)` | `items[].description` |
| `Sale price` | `items[].price` |
| `Quantity` | `items[].quantity` |
| `Ship via` | 초기 추천 운송사에만 사용 가능. 실제 변환은 화면 드롭다운 선택값 기준. |
| `Receiver name` | `receiver.name` |
| `Phone no` | `receiver.phone` |
| `Country` | `receiver.country` |
| `Province` | `receiver.state` |
| `City` | `receiver.city` |
| `Address1` | `receiver.address1` |
| `Address2` | `receiver.address2` |
| `Zip` | `receiver.postalCode` |
| `Seller E-mail` | `receiver.email` |
| `Tax ID` | `receiver.taxId` |
| `Total Weight` | `package.weightKg` |
| `Recommended Box Type` | `package.boxType` |

## 운송사 선택 규칙

최종 변환 운송사는 raw data의 `Ship via`가 아니라 페이지의 운송사 드롭다운 선택값을 기준으로 한다.

`Ship via`는 선택값 자동 추천에만 사용할 수 있다. 사용자가 드롭다운을 바꾸면 Generate 시점에는 사용자 선택값이 우선한다.

## 수출자 정보 규칙

Dropshipping 모델이므로 `company`와 `contact`는 기본값을 넣지 않는다. 두 필드는 shipment마다 달라질 수 있어 항상 공백으로 시작하고 사용자가 직접 입력한다.

전화번호, 이메일, 주소, 우편번호, 사업자번호처럼 고정 운영 정보는 기본값으로 제공하되 설정 패널에서 수정 가능하게 둔다.

## DHL

| 항목 | 구현 방식 |
| --- | --- |
| 출력 포맷 | CSV |
| 기준 헤더 | `sample/dhl sample.csv`의 헤더 |
| 다품목 주문 | 우선 품목별 행 유지. 같은 주문번호의 여러 품목은 여러 CSV 행으로 출력 |
| 인코딩 | UTF-8 BOM 포함 CSV 권장 |

## 우체국 K-packet

기존 `index.html`에서 확인한 규칙을 이식한다.

| 항목 | 구현 방식 |
| --- | --- |
| 출력 포맷 | XLSX |
| 시트명 | `우체국등록` |
| 파일명 | `KoreaPost_Bulk_YYYY-MM-DD.xlsx` |
| 상품구분 | Generate 전 또는 설정에서 `Gift`/`Merchandise` 선택 |
| 헤더 | 1행 제목, 2행 컬럼명, 3~5행 빈 행, 6행부터 데이터 |
| 다품목 | 상품명, 수량, 가격, 순중량, HS CODE, 생산지를 `;`로 연결 |
| HS CODE | 기본 `3304991000` |
| 생산지 | 기본 `KR` |
| 순중량 | 품목별 `30`g |
| 박스 실중량 | `package.weightKg * 1000` |
| EMS구분/물품종류 | CL/MX는 `R`/`re`, 그 외 국가는 `K`/`rl` |
| 보험가입여부 | `N` |

## ACI

| 항목 | 구현 방식 |
| --- | --- |
| 출력 포맷 | XLSX |
| 기준 시트 | `Form` |
| 국가별 양식 | 현재 US 샘플만 있음. `ACI_US_YYYY-MM-DD.xlsx`를 우선 구현 |
| 다품목 | 샘플 구조상 품목별 행 또는 주문별 첫 행+품목행 혼합 가능성이 있어, 1차 구현은 품목별 행으로 안정화 |
| 고정값 | `출발도시 코드=082`, `PAYMENT=DDP`, `도착국가=USA`, `화폐 단위=USD` |

## FEDEX

| 항목 | 구현 방식 |
| --- | --- |
| 출력 포맷 | `.in` text file |
| 구조 | `fieldNumber,"value"` 행 반복 |
| 기본값 | `기본값(필수 입력)` 시트의 sender/default field number 구조를 코드 설정으로 이식 |
| 주문값 | `020 Transaction` 시트 3행 field number 기준 |
| 다품목 | `Grouping Fields`의 반복 필드 후보를 참고. 1차 구현은 품목별 `.in` 필드를 주문 레코드 안에서 순서대로 반복 |
| 한계 | VBA 내부 예외 처리까지 100% 보장하려면 추가 `.in` 샘플 비교 필요 |

## Phase 1 결론

모든 운송사는 GitHub Pages 정적 앱에서 1차 변환 구현 가능하다.

남은 검증 리스크는 FEDEX `.in`의 국가별 예외, ACI 국가별 템플릿 차이, 각 운송사 홈페이지 실제 업로드 검증이다.
