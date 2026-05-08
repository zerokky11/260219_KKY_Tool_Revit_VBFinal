# KKY TOOL Version Release Checklist

사용자가 "버전 업데이트해줘", "홈페이지에 올리고 푸시해줘"라고 요청하면 이 순서대로 진행한다.

## 1. 릴리즈 기준 확인

- `KKY_TOOL_RELEASE_GATE.md`를 먼저 읽고, 설치/업데이트/홈페이지 배포 위험을 확인한다.
- 현재 안정 배포 범위는 Revit 2019/2021/2023/2025로 본다.
- Revit 2027은 별도 검증 전까지 파일명, ZIP 패키지, `latest.json`, 홈페이지 링크에 포함하지 않는다.
- `git status --short`로 작업 중인 변경을 확인하고, Family Browser 파일은 사용자가 명시하지 않으면 건드리지 않는다.

## 2. 릴리즈 내용 정리

- 이번 버전에 들어갈 변경점을 사용자 화면 기준으로 2~4개 문장으로 정리한다.
- Revit 모델 수정 로직, 설치 방식, 업데이트 방식이 바뀌었는지 확인하고, 바뀐 경우 위험을 설명한다.
- 기능적으로 확인되지 않은 항목은 "현장 테스트 필요"로 남기고 완료처럼 쓰지 않는다.

## 3. 버전 번호 동기화

- `Compile/KKY_Tool_Compiler.iss`의 `MyAppVersion`을 새 버전으로 맞춘다.
- `KKY_Tool_Revit_2019-2023/My Project/AssemblyInfo.vb`의 `AssemblyVersion`, `AssemblyFileVersion`, `AssemblyInformationalVersion`을 맞춘다.
- `KKY_Tool_Revit_2019-2023/Resources/HubUI/js/core/topbar.js`의 `APP_VERSION_FALLBACK`을 맞춘다.
- 허브 JS/CSS를 바꿨다면 `Resources/HubUI/index.html`과 `Resources/HubUI/js/main.js`의 `?v=` 캐시 번호를 새 날짜/릴리즈 기준으로 갱신한다.
- 필요하면 `docs/update-feed.sample.json`, `docs/update-hosting-setup.md`, `docs/test_build_history.md`의 현재 예시 버전도 맞춘다.

## 4. 홈페이지/업데이트 피드 반영

- `Compile/Build-Release.ps1 -Notes "<변경점>|<변경점>|<변경점>"`로 릴리즈 빌드를 실행한다.
- 스크립트가 `Sever/Release/latest.json`, `Sever/Release/release-history.json`, `Sever/Release/index.html`을 새 버전으로 갱신했는지 확인한다.
- 홈페이지 문구에는 `Release build`, `API`, `Apps Script`, 내부 설정명 같은 구현자용 표현을 쓰지 않는다.

## 5. 빌드/검증

- 릴리즈 빌드가 Revit 2019/2021/2023/2025 산출물, 설치 EXE, 업데이트 ZIP을 모두 만들었는지 확인한다.
- `Compile/Verify-KKYToolRelease.ps1 -Version <버전>` 검증을 통과해야 한다.
- `latest.json`의 `version`, `url`, `publishedAt`, `notes`와 `release-history.json` 첫 항목이 새 버전인지 확인한다.
- 새 ZIP/EXE 파일명이 `KKY_Tool_Revit(2019,21,23,25)_v<버전>` 형식인지 확인한다.

## 6. Git 정리와 푸시

- 빌드 산출물 중 배포에 필요한 `Sever/Release`의 새 EXE/ZIP/JSON/홈페이지 파일만 포함한다.
- `_build`, `_buildcheck`, `_temp_build`, `artifacts/release-stage`, 임시 테스트 파일은 커밋하지 않는다.
- `git add`는 관련 소스, 릴리즈 스크립트/문서, 홈페이지/피드/배포 산출물만 명시적으로 지정한다.
- 커밋 메시지는 `Release KKY Tool v<버전>` 형식으로 남긴다.
- 커밋 후 현재 브랜치로 `git push origin <브랜치>`를 실행한다.
