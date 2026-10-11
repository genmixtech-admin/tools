# tools

젠믹스텍 사내 업무 도구 (GitHub Pages로 게시)

| 도구 | 파일 | 웹 주소 |
|---|---|---|
| 출장비 정산 (v1.8) | `travel-expense.html` | https://genmixtech-admin.github.io/tools/travel-expense.html |
| 유류대 지급신청 (v1.0) | `fuel-claim.html` | https://genmixtech-admin.github.io/tools/fuel-claim.html |

- 원본: Obsidian `100_Intranet_경영정보/tools/`
- 관리자: genmixtech.ai@gmail.com
- 입력 데이터는 각 사용자 브라우저에만 저장되며 서버로 전송되지 않습니다.
- 회사 정보·개인정보를 이 저장소(공개)에 올리지 마세요. (GitHub 사용 정책 참조)

## 메일 발송 (v1.5~)
- `mail-sender.gs`: genmixtech.ai@gmail.com 계정의 Google Apps Script(웹 앱)로 배포. 페이지가 엑셀을 보내면 Gmail로 첨부 발송.
- 하루 발송 한도 50건, 모든 발송은 관리 계정에 숨은 참조로 기록.

## 유류대 지급신청 (v1.0)
- 백엔드: 별도 Google Apps Script 웹 앱(오피넷 경기 평균가 · 네이버 Maps 거리 · 카카오 장소검색 · 영수증 OCR · 메일 발송). API 키와 접속 토큰은 Apps Script 스크립트 속성에만 보관하며 이 저장소에는 없습니다.
- 사용자는 처음 한 번 담당자에게 받은 접속 토큰만 입력합니다.
