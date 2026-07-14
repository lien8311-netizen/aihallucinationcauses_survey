# AI 창작물 불쾌한 골짜기 설문

구성 파일:

- `survey.html`: 응답자용 단일 파일 설문
- `admin.html`: 관리자용 단일 파일 대시보드
- `Code.gs`: Google Sheets 바인딩 Apps Script 백엔드

## 배포 방법

1. 새 구글 스프레드시트를 만듭니다.
2. **확장 프로그램 > Apps Script**를 열고 기본 코드를 지운 뒤 `Code.gs` 전체를 붙여넣어 저장합니다.
3. **배포 > 새 배포 > 유형: 웹 앱**을 선택합니다. 실행 계정은 **나**, 액세스 권한은 **모든 사용자**로 지정하고 배포합니다. 권한 요청이 나오면 승인합니다.
4. 발급된 `/exec` 웹 앱 URL을 `survey.html`과 `admin.html` 상단 `CONFIG.WEB_APP_URL`에 똑같이 입력합니다. 두 파일의 자극 URL·정체 매핑도 실제 파일에 맞게 확인하고, `admin.html`의 `ADMIN_CODE`를 원하는 값으로 바꿉니다.
5. `survey.html`에서 테스트 응답 1건을 제출하고 스프레드시트의 `Responses` 시트에 한 행이 생기는지 확인합니다. 이어 `admin.html`에서 접속 코드를 입력해 조회되는지 확인합니다.

정적 호스팅(학교 서버, GitHub Pages 등)에서 HTML 파일과 자극 파일을 함께 제공하세요. 로컬 파일(`file://`)보다 HTTPS 호스팅을 권장합니다. Apps Script 코드를 변경한 경우 **배포 관리 > 수정 > 새 버전**으로 다시 배포해야 반영됩니다.

같은 브라우저의 중복 제출 차단을 테스트 중 초기화하려면 개발자 도구 Console에서 다음을 실행합니다.

```js
localStorage.removeItem('uviSurveySubmitted');
localStorage.removeItem('uviSurveyPending');
```
