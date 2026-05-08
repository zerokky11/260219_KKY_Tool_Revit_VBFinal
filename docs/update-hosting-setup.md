# KKY Tool Update Hosting Setup

## 목적

KKY TOOL은 `latest.json`을 확인해 새 버전 여부를 판단하고, 사용자가 업데이트를 선택하면 지정된 ZIP 또는 EXE 파일을 다운로드한다.

권장 운영 방식은 다음과 같다.

- 홈페이지: `https://update.zerokky.com`
- 업데이트 피드: `https://update.zerokky.com/latest.json`
- 배포 파일: `https://update.zerokky.com/KKY_Tool_Revit(2019,21,23,25)_v버전.zip`
- 업데이트 내역: `https://update.zerokky.com/updates.html`
- 기능 요청: `https://update.zerokky.com/requests.html`

현재 안정 배포 범위는 Revit 2019/2021/2023/2025다. Revit 2027은 2027용 AddIn 파일과 RevitAPI 참조를 별도로 검증한 뒤 배포 파일명과 업데이트 피드에 포함한다.

## latest.json 형식

```json
{
  "version": "2.25",
  "url": "https://update.zerokky.com/KKY_Tool_Revit(2019,21,23,25)_v2.25.zip",
  "publishedAt": "2026-05-08",
  "notes": "View Navigator 창 전환 안정화, 문서 색상 구분 개선, 납품용 RVT 정리 결과 폴더 선택창 개선을 반영했습니다."
}
```

필드 기준:

- `version`: 사용자가 볼 최신 버전
- `url`: 다운로드할 ZIP 또는 EXE 주소
- `publishedAt`: 배포일
- `notes`: 앱과 홈페이지에 표시할 핵심 변경 내용

## 애드인 설정

설정 파일:

```text
KKY_Tool_Revit_2019-2023/Resources/update-config.json
```

예시:

```json
{
  "feedUrls": [
    "https://update.zerokky.com/latest.json"
  ],
  "downloadDirectory": "%LOCALAPPDATA%\\KKY_Tool_Revit\\Updates"
}
```

동작 기준:

- `feedUrls` 순서대로 접근한다.
- 먼저 연결되는 피드를 사용한다.
- `feedUrl`도 하위 호환용으로 동작하지만, 새 구성은 `feedUrls`를 사용한다.

## 홈페이지 문구 운영

홈페이지 문구는 `docs/kky-tool-homepage-update-copy.md`를 기준으로 정리한다.

업데이트 문구는 실제 구현된 기능만 설명한다. 아직 테스트 중인 기능은 배포 노트에 안정 기능처럼 쓰지 않는다.

## 배포 순서

1. 설치/업데이트 패키지를 빌드한다.
2. `KKY_TOOL_RELEASE_GATE.md`를 확인한다.
3. `latest.json`, `release-history.json`, `Compile/KKY_Tool_Compiler.iss`의 `MyAppVersion`, `AssemblyInformationalVersion`, Hub 표시 버전을 같은 버전으로 맞춘다.
4. ZIP 또는 EXE 파일을 서버에 업로드한다.
5. `latest.json`의 버전, 다운로드 주소, 변경 내용을 갱신한다.
6. `release-history.json`에 최신 버전 기록을 추가한다.
7. 홈페이지에서 다운로드 링크와 업데이트 내역을 확인한다.
8. Revit에서 `업데이트 확인`을 실행해 실제 피드가 읽히는지 확인한다.
9. Windows 프로그램 추가/제거의 버전 표시가 Hub 표시 버전과 같은지 확인한다.

## 트래픽 참고

200 MB 파일 기준:

- 50회 다운로드: 약 10 GB
- 100회 다운로드: 약 20 GB
- 500회 다운로드: 약 100 GB

회사 보안망에서 EXE 다운로드가 차단될 수 있으므로 ZIP 업데이트 경로를 우선 운영하고, 필요하면 보조 미러를 추가한다.

## 실무 메모

- HTTPS 443 포트 기반 운영을 기본으로 한다.
- 클라우드 공유 링크보다 일반 도메인 호스팅이 차단 가능성이 낮다.
- 배포 파일명을 바꾸면 `latest.json`, 홈페이지 링크, 업데이트 내역을 함께 맞춘다.
- 프로그램 추가/제거에 이전 버전이 남아 있으면 설치 정보가 아직 갱신되지 않은 상태일 수 있으므로, 업데이트 적용 또는 재설치 후 다시 확인한다.
- `_build`, `_temp_build`, `_buildcheck` 폴더는 검증 과정에서 남은 이전 산출물을 포함할 수 있으므로, 최신 버전 판단은 소스/패키징 입력과 새로 만든 배포 산출물 기준으로 한다.
- 업데이트 실패 메시지는 사용자 조치와 관리자 전달 정보를 함께 보여줘야 한다.
