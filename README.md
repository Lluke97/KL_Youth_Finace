# youth-budget — 청년부 예산 대시보드

청년부 카페 수입/지출(jusumcafe.up.railway.app), 회비, 기타 수동 입력을 한 곳에 모아
현재 잔액과 추이를 보여주는 대시보드. cafe-pos와 같은 패턴(Express + JSON 파일 DB,
Railway 볼륨)으로 만들어졌습니다.

## 구조
- `db.js` — JSON 파일 기반 저장소. `seed-data/transactions.seed.json`은 집사님이
  정리하신 `Open_Cafe-2026.xlsx` 과거 내역을 1회 임포트한 결과로, **볼륨이
  비어있는 최초 배포 시에만** 초기값으로 쓰입니다. 그 뒤로는 `data/transactions.json`
  (Railway 볼륨)이 실제 데이터입니다.
- `lib/cafeSync.js` — cafe-pos의 관리자 API(`/api/admin/orders`, `/api/admin/expenses`)에서
  새 내역을 끌어옵니다. 지출 카테고리 중 "기타"는 카페 운영비인지 청년부 일반지출인지
  애매해서 자동으로 `needs-review`로 태깅되고, 관리자 화면에서 분류를 정해주면 됩니다.
- `lib/sheetsSync.js` — 구글 시트로 전체 내역을 내보내는 기능(서비스 계정 필요, 아래 참고).
- `scripts/import-excel.js` — 엑셀 과거 데이터 1회 임포트용 (재실행해도 중복 안 됨).
- `scripts/sync-cafe.js` — cafe-pos 동기화를 커맨드라인/크론에서 실행.
- 정산(실제 현금/터치앤고/은행 잔액)은 거래 장부와 별도로 `settlements`에 쌓입니다.
  cafe-pos에서 "정산하기"를 누르면 `/api/webhook/cafe-settlement`로 자동 전달되고,
  관리자 화면에서 직접 입력할 수도 있습니다. `seed-data/settlements.seed.json`은
  볼륨이 비어있을 때만 쓰이는 초기값입니다.

## 환경변수
| 변수 | 설명 |
|---|---|
| `ADMIN_USER`, `ADMIN_PASS` | 관리자 페이지(`/admin`) 로그인 |
| `ADMIN_USER2`, `ADMIN_PASS2` | (선택) 보조 관리자 계정 |
| `REQUIRE_ADMIN_AUTH` | `false`로 설정하면 `/admin`과 관리자 API 전부 로그인 없이 열림 (기본값: 로그인 필요) |
| `CAFE_BASE_URL` | 기본값 `https://jusumcafe.up.railway.app` |
| `CAFE_ADMIN_USER`, `CAFE_ADMIN_PASS` | cafe-pos 쪽 `ADMIN_USER`/`ADMIN_PASS`와 동일한 값 |
| `SETTLEMENT_WEBHOOK_SECRET` | cafe-pos가 정산 내용을 보낼 때 쓰는 공유 비밀값 (cafe-pos의 같은 변수와 일치해야 함) |
| `GOOGLE_SHEET_ID` | (선택) 내보낼 구글 시트 ID |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | (선택) 구글 서비스계정 키 JSON 전체 내용 |

## 구글 시트 연동 설정
1. Google Cloud Console에서 프로젝트 생성 → "Google Sheets API" 활성화
2. IAM → 서비스 계정 생성 → Keys 탭 → JSON 키 다운로드
3. 쓸 구글 시트를 열고 Share에서 서비스계정 이메일(`...iam.gserviceaccount.com`)을
   편집자로 추가
4. Railway 변수에 `GOOGLE_SHEET_ID`(시트 URL의 ID 부분)와 `GOOGLE_SERVICE_ACCOUNT_JSON`
   (키 JSON 파일 내용 그대로)을 등록

## 로컬 실행
```
npm install
node scripts/import-excel.js /path/to/Open_Cafe-2026.xlsx   # 처음 1회만
ADMIN_USER=admin ADMIN_PASS=yourpass npm start
```
