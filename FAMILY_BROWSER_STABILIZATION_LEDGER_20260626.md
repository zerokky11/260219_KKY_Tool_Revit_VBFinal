# Family Browser Stabilization Ledger - 2026-06-26

이 문서는 Family Browser 복구/안정화 작업이 다시 꼬이지 않도록 고정하는 작업 기준표다.

## 현재 원칙

- 설치파일 생성보다 기준선 고정, 백업, 기능별 검증을 우선한다.
- 수정 전에는 전체 source checkpoint와 수정 파일 백업을 만든다.
- `FamilyBrowserDashboardHtmlForm.cs`는 UI 대부분을 담당하므로 큰 덩어리 교체를 금지하고 작은 패치만 적용한다.
- 정상 동작으로 확인되지 않은 디컴파일 소스나 예전 백업을 현재 기준선처럼 사용하지 않는다.
- 빌드 성공은 기능 복구 증거가 아니다. 기능별 검증 항목을 별도로 남긴다.

## 이번 작업 기준

- 시작 checkpoint: `artifacts/source-checkpoints/20260626-142200-stabilization-audit-p0`
- 수정 파일 백업 폴더: `artifacts/source-file-backups/20260626/20260626-142200-stabilization-audit-p0`
- 작업 방식: P0부터 좁게 수정하고 빌드/코드 검토 후 다음 항목으로 이동

## 요구사항 대비 상태표

| ID | 기능 | 사용자가 요구한 동작 | 현재 관찰/위험 | 우선순위 | 상태 |
| --- | --- | --- | --- | --- | --- |
| FB-P0-001 | 표준 RVT 변경 추적 | 애드인 창을 열지 않아도 표준 RVT 저장/동기화 시 패밀리 로드/삭제/이름변경 로그가 확정되어야 함 | `App.cs`에서 `DocumentChanged`, `DocumentSaving`, `DocumentSavingAs`, `DocumentSynchronizedWithCentral` 구독 추가 | P0 | 1차 수정 완료 |
| FB-P0-002 | 파일별 권한 차단 | 파일별 권한 설정이 있는 대상 RVT에서만 패밀리 로드/타입명/패밀리명 변경을 막고, 미등록 파일은 막지 않아야 함 | `ShouldEnableProtectedChangeUpdater`, `HandleDocumentChanged`, updater 실행 경로를 현재 문서가 FileGuard 대상일 때만 진행하도록 축소 | P0 | 1차 수정 완료 |
| FB-P0-003 | 프로젝트 열림 속도 | 프로젝트 열 때 정책/스캔/인덱스 같은 무거운 작업을 하지 않아야 함 | 문서 활성화 경로는 가볍게 유지하고, 브라우저 초기 렌더 중 최신 프로젝트 스캔 캐시 로드는 1200ms로 제한 | P0 | 1차 수정 완료 |
| FB-P0-004 | 상세항목 새창 | 패밀리 로드/시스템 타입 로드에서 상세항목은 별도 창으로 뜨고 본문 리스트는 넓게 써야 함 | 패밀리/시스템 타입 탭에서만 detached detail 모드 적용, 다른 탭에서 상세창 호출 방지 | P0 | 1차 수정 완료 |
| FB-P0-005 | 탭 스크롤 | 모든 탭은 내용이 길면 스크롤로 확인 가능해야 함 | 이번 P0 패치에서 직접 수정하지 않음. 다음 패치 전 CSS/JS 스크롤 규칙부터 확인 필요 | P0 | 대기 |
| FB-P0-006 | 하위 패밀리 필터 | 다른 패밀리 안에 로드된 FamilyName이 표준 Excel 리스트에도 있으면 하위패밀리로 보고 로드 리스트에서 제외 | nested 후보를 표준 Excel FamilyName 인덱스로 한 번 더 제한하고, 표준 리스트가 없으면 하위패밀리 확정 처리하지 않음 | P0 | 1차 수정 완료 |
| FB-P0-007 | 언어/문자열 | 한글/영어 전환 시 화면 문자열과 결과창 문자열이 설정 언어에 맞아야 함 | 일부 소스 문자열 자체가 깨져 있을 수 있음 | P0 | 조사 중 |
| FB-P1-001 | 결과창 | 검사/스캔/로드/Excel import 결과창을 일관된 HTML 스타일로 표시 | MessageBox/TaskDialog/WinForms 혼재 | P1 | 대기 |
| FB-P1-002 | 표준 Excel import | Excel 시트는 입력이 아니라 목록 선택, import 후 과한 전체 refresh 금지 | 시트 선택은 있으나 import 후 refresh 병목 가능 | P1 | 대기 |
| FB-P1-003 | 3D 미리보기 | 정밀 스캔 PNG가 있으면 상세항목에서 중앙 정렬로 표시, 객체 잘림 없어야 함 | capture/display 양쪽 점검 필요 | P1 | 대기 |
| FB-P1-004 | fingerprint | CSV/Lookup Table/파라미터/타입 비교가 표준과 프로젝트 스캔에서 같은 기준이어야 함 | CSV는 포함되어 보이나 캐시 갱신 흐름 점검 필요 | P1 | 대기 |
| FB-P1-005 | 시스템 타입 | 라우팅 프리퍼런스/세그먼트/의존성 차이를 상세 표로 설명 | 엔진 일부 반영, UI 요약 점검 필요 | P1 | 대기 |

## 이번 P0 수정 순서

1. `App.cs` 이벤트 구독 누락 여부 수정.
2. FileGuard가 현재 문서가 대상 파일일 때만 무거운 판단/Updater를 사용하도록 축소.
3. 상세항목 새창과 탭 스크롤 CSS/JS 충돌 정리.
4. nested child 판정과 상세 표시를 FamilyName 기준으로 통일.
5. 깨진 문자열은 우선 startup/loading/detail/file-guard 핵심 경로부터 복구.

## 검증 체크리스트

| 검증 | 방법 | 결과 |
| --- | --- | --- |
| 2019-2023 빌드 | `dotnet build KKY_FamilyBrowser_RevitHost_2019-2023/KKY_FamilyBrowser_RevitHost.csproj -v:minimal` | 성공, 오류 0 |
| 2025 빌드 | `dotnet build KKY_FamilyBrowser_RevitHost_2025/KKY_FamilyBrowser_RevitHost_2025.csproj -v:minimal` | 성공, 오류 0 |
| 2027 빌드 | `dotnet build KKY_FamilyBrowser_RevitHost_2027/KKY_FamilyBrowser_RevitHost_2027.csproj -v:minimal` | 성공, 오류 0 |
| FileGuard 코드 검토 | 대상 파일이 아니면 native guard/Updater가 no-op인지 확인 | 1차 통과: 대상 문서가 아니면 DocumentChanged/updater 경로 return |
| 상세항목 코드 검토 | browser load/type tabs 외 다른 탭에서 detail window가 뜨지 않는지 확인 | 1차 통과: `families`/`systems` 탭에서만 detached detail 갱신 |
| nested 코드 검토 | 표준 Excel 리스트에 포함된 nested FamilyName만 로드 목록에서 제외되는지 확인 | 1차 통과: nested summary/list filtering 모두 standard list index를 사용 |

## 작업 로그

- 2026-06-26 14:22: 전체 source checkpoint 생성.
- 2026-06-26 14:22: 안정화 ledger 작성 시작.
- 2026-06-26 14:29: `App.cs` 이벤트 구독 복구 및 FileGuard 대상 문서 축소 패치 적용.
- 2026-06-26 14:31: 2019-2023, 2025, 2027 일반 빌드 성공 확인.
- 2026-06-26 14:45: UI/nested P0 edit backup created at `artifacts/source-file-backups/20260626/20260626-ui-detail-nested-p0`.
- 2026-06-26 14:45: Browser tabs now set `fb-detached-detail`; standard-list-empty nested fallback is disabled so nested display/list hiding only trusts standard Excel family names.
- 2026-06-26 14:45: Detached detail window default Korean title restored to `선택 항목 상세`.
- 2026-06-26 14:45: 2019-2023, 2025, 2027 builds passed again after UI/nested patch; errors 0.
- 2026-06-26 15:12: Startup/cache P0 edit backup created at `artifacts/source-file-backups/20260626/20260626-startup-cache-p0`.
- 2026-06-26 15:12: Lightweight dashboard cache key no longer enumerates latest project scan record files while `_deferExpensiveStartupLookups` is active.
- 2026-06-26 15:12: Latest project scan cache loading during lightweight dashboard refresh is now bounded to 1200 ms; if it times out the dashboard falls back to standard-list/name-category rows and logs `project-scan-cache-load-timeout`.
- 2026-06-26 15:13: 2019-2023, 2025, 2027 builds passed after startup/cache patch; errors 0.
- 2026-06-26 15:18: Ledger status updated so completed P0 patches and untouched pending items are separated. Next work must start from this ledger instead of guessing from older backups.
- 2026-06-26 15:16: Installer built for external PC testing: `artifacts/family-browser/installers/KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v0.1_p0-stabilization-test_Setup.exe`.
- 2026-06-26 15:16: Mail-sized package built: `artifacts/family-browser/mail-packages/20260626_03.zip` (`13.6 MB`, SHA256 `71CF7CB4803C08104C1F6A8E185E650C0592FD60693E0B6231EE6B01472FE979`).
