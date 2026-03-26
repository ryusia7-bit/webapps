# 척도검사 Apps Script 웹앱

이 폴더는 `MindMap 다시서기`의 `Apps Script 웹앱` 1차 MVP입니다.

## 포함 기능

- 척도 8종 선택
- 응답자 정보/문항 응답 입력
- 결과 계산과 인쇄
- Google Sheets 직접 저장
- 최근 저장 결과 조회
- 척도 마스터 자동 동기화

## 구성 파일

- `Code.gs`: Apps Script 서버 함수와 시트 저장 로직
- `Index.html`: 웹앱 화면 껍데기
- `Styles.html`: 화면/인쇄 스타일
- `ClientApp.html`: 클라이언트 계산/렌더링 로직
- `QuestionnaireBundle.html`: 척도 번들 데이터

## 배포 순서

1. `clasp push`
2. Apps Script에서 `배포 > 새 배포 > 웹 앱`
3. 실행 사용자 `배포하는 사용자`
4. 접근 권한 `모든 사용자(로그인 여부와 관계없이)` 또는 `Anyone`
5. 배포 URL로 접속 후 앱 내부 로그인

## 참고

- 기본 대상 스프레드시트 ID는 `11y5p7Cp_yN2vggMOlCwn4pKNBEmio-CmkK25Nyd2nIk`입니다.
- 웹앱 첫 실행 시 척도 마스터 시트와 저장 시트를 자동 점검합니다.
- 현재 구현본은 Apps Script 프로젝트 `1RAhgFNBQylhMIJddrta9_rxttuHfvcsIfyccGP5_lYtaCSP9oh8xcJ0F`에 푸시되었습니다.
- 앱 접근은 공개로 열어두고, 앱 내부 로그인으로 사용자를 제한합니다.
- 초기 로그인 계정은 `admin0109` 또는 `ryusia7@homeless.go.kr`이며, 비밀번호는 `admin0109`입니다.
- 로그인 화면에는 가입신청이 있고, 관리자 계정으로 로그인하면 계정 등록/수정/비활성화와 가입신청 처리를 할 수 있습니다.
