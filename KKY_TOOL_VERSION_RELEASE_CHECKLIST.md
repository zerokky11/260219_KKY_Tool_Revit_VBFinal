# KKY TOOL Version Release Checklist

사용자가 "버전 업데이트해줘", "홈페이지에 올리고 푸시해줘"라고 요청하면 이 순서대로 진행한다.

## 1. 릴리즈 기준 확인

- `KKY_TOOL_RELEASE_GATE.md`를 먼저 읽고, 설치/업데이트/홈페이지 배포 위험을 확인한다.
- 현재 정식 배포 범위는 Revit 2019/2021/2023/2025 기준으로 본다.
- Revit 2027은 별도 검증 전까지 설치 파일명, ZIP 패키지, `latest.json`, 홈페이지 링크에 포함하지 않는다.
- `git status --short`로 작업 중인 변경을 확인하고, `KKY_FamilyBrowser_*` 파일은 사용자가 명시하지 않으면 건드리지 않는다.

## 2. 변경 내용 정리

- 이번 버전에 들어간 사용자-facing 변경점을 2~4개 문장으로 정리한다.
- Revit 모델 수정 로직, 설치 방식, 업데이트 방식이 바뀌었는지 확인하고 바뀐 경우 위험을 설명한다.
- 기능 검증이 부족한 항목은 "현장 테스트 필요"로 표시하고 완료처럼 쓰지 않는다.

## 3. 버전 번호 동기화

- `Compile/KKY_Tool_Compiler.iss`의 `MyAppVersion`을 새 버전으로 맞춘다.
- `KKY_Tool_Revit_2019-2023/My Project/AssemblyInfo.vb`의 `AssemblyVersion`, `AssemblyFileVersion`, `AssemblyInformationalVersion`을 맞춘다.
- `KKY_Tool_Revit_2019-2023/Resources/HubUI/js/core/topbar.js`의 `APP_VERSION_FALLBACK`을 맞춘다.
- Hub JS/CSS를 바꿨다면 `Resources/HubUI/index.html`과 `Resources/HubUI/js/main.js`의 `?v=` 캐시 버전을 갱신한다.

## 4. 빌드와 배포 산출물

- 정식 배포는 `Compile/Build-Release.ps1 -Notes "<변경점>|<변경점>|<변경점>"`로 실행한다.
- 정식 설치 EXE는 `Sever/Release/official/`에 생성되어야 한다.
- 테스트 설치 EXE는 `Compile/Build-TestVersion.ps1`로 만들며 `Sever/Release/test/`에 생성되어야 한다.
- 업데이트 ZIP, `latest.json`, `release-history.json`, `index.html`은 기존처럼 `Sever/Release/` 루트에서 관리한다.
- 정식 설치 URL은 `https://update.zerokky.com/official/KKY_Tool_Revit(2019,21,23,25)_v<버전>.exe` 형식이어야 한다.
- 업데이트 ZIP URL은 `https://update.zerokky.com/KKY_Tool_Revit(2019,21,23,25)_v<버전>.zip` 형식이어야 한다.

## 5. 검증

- 빌드가 Revit 2019/2021/2023/2025 출력물, 정식 설치 EXE, 업데이트 ZIP을 모두 만들었는지 확인한다.
- `Compile/Verify-KKYToolRelease.ps1 -Version <버전>` 검증을 통과해야 한다.
- `latest.json`의 `version`, `url`, `publishedAt`, `notes`와 `release-history.json` 첫 항목이 새 버전인지 확인한다.
- `Sever/Release` 루트에 설치 EXE가 직접 남아 있지 않은지 확인한다.

## 6. Git 정리와 푸시

- 커밋 대상에는 관련 소스, 릴리즈 스크립트/문서, 홈페이지/피드/배포 산출물만 명시적으로 포함한다.
- `_build`, `_buildcheck`, `_temp_build`, `artifacts/release-stage`, 임시 테스트 파일은 커밋하지 않는다.
- 커밋 메시지는 `Release KKY Tool v<버전>` 형식을 기본으로 한다.
- 커밋 후 현재 브랜치로 `git push origin <브랜치>`를 실행한다.
