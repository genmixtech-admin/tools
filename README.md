# tools

젠믹스텍 사내 업무 도구 (GitHub Pages로 게시)

| 도구 | 파일 | 웹 주소 |
|---|---|---|
| 출장비 정산 (v1.9) | `travel-expense.html` | https://genmixtech-admin.github.io/tools/travel-expense.html |
| 유류대 지급신청 (v1.8) | `fuel-claim.html` | https://genmixtech-admin.github.io/tools/fuel-claim.html |

- 원본: Obsidian `100_Intranet_경영정보/tools/`
- 관리자: genmixtech.ai@gmail.com
- 입력 데이터는 각 사용자 브라우저에만 저장되며 서버로 전송되지 않습니다.
- 회사 정보·개인정보를 이 저장소(공개)에 올리지 마세요. (GitHub 사용 정책 참조)

## 메일 발송 (v1.5~)
- `mail-sender.gs`: genmixtech.ai@gmail.com 계정의 Google Apps Script(웹 앱)로 배포. 페이지가 엑셀을 보내면 Gmail로 첨부 발송.
- 하루 발송 한도 50건, 모든 발송은 관리 계정에 숨은 참조로 기록.

## 출장비 정산 · 엑셀 불러오기 (v1.9~)
- 이 도구로 내보낸 엑셀(v1.0~v1.9 형식)을 '엑셀 불러오기'로 열면 내역과 출장 정보가 화면에 반영되어 이어서 입력할 수 있습니다.
- 출장 구분을 '해외'로 고르면 항공·숙박·해외교통·국내교통·통신·식음료 기본 항목이 '입력 필요' 상태로 만들어집니다.
- 화면 제목 옆과 하단, 엑셀 4행에 도구 버전이 표시됩니다.

## 유류대 지급신청 (v1.8)
- `fuel-backend.gs` (v12): genmixtech.ai@gmail.com 계정의 Google Apps Script(웹 앱)로 배포. 오피넷 경기 평균가, 도로거리(장소 LOOKUP 시트 → 카카오/네이버), 통행료 영수증 OCR, 엑셀 첨부 메일 발송.
- API 키·접속 토큰·LOOKUP 시트 ID는 Apps Script 스크립트 속성에만 보관하며 이 저장소에는 없습니다.
- 자주 가는 곳(명칭·주소·별칭)은 사내 구글 시트 「유류대 장소 LOOKUP」에서 관리합니다. 주소 목록은 이 공개 저장소에 올리지 않습니다.
- 사용자는 처음 한 번 담당자에게 받은 접속 토큰만 입력합니다.
- 거리 점검: 제출된 주행기록을 사내 시트 '주행이력' 탭에 쌓아 구간별 중앙값과 비교, ±15% 또는 ±5km를 넘으면 경고(사유 입력 필수).
