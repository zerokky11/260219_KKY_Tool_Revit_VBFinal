# KKY Family Browser Next Chat Handoff - 2026-07-06

## 0. 절대 주의: 작업 폴더

새 채팅에서 가장 먼저 확인할 것:

```powershell
Set-Location 'C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628'
Get-Location
```

실제 패밀리 브라우저 1.0 작업 폴더는 아래다.

```text
C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628
```

비슷한 폴더가 하나 더 있다.

```text
C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit
```

이 폴더는 현재 Codex 기본 workspace root로 잡힐 수 있지만, 최근 패밀리 브라우저 수정/빌드/설치 검증을 진행한 폴더가 아니다. 새 채팅에서 이 폴더를 기준으로 작업하면 다른 버전을 만질 위험이 크다.

## 1. 새 채팅 첫 지시문

새 채팅을 시작하면 이렇게 말하면 된다.

```text
C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\FAMILY_BROWSER_NEXT_CHAT_HANDOFF_20260706.md 를 먼저 읽고, 같은 폴더의 FAMILY_BROWSER_BUTTON_AUDIT.md 도 이어서 읽어. 작업 폴더를 반드시 KKY_Tool_Revit_full_20260628 로 맞춘 뒤 패밀리 브라우저 수정을 이어가. 2019/2023/2025/2027 네 버전을 모두 같이 반영하고, 수정 전 백업, 정적 QA, 빌드, 스테이지 검증, 설치, 설치본 검증까지 유지해.
```

## 2. 기준 문서

반드시 이 순서로 읽고 시작한다.

1. `FAMILY_BROWSER_NEXT_CHAT_HANDOFF_20260706.md`
2. `FAMILY_BROWSER_BUTTON_AUDIT.md`
3. 필요 시 `FAMILY_BROWSER_1_0_BLUEPRINT_20260629.md`
4. 필요 시 `FAMILY_BROWSER_INTENT_AUDIT_20260626.md`

`FAMILY_BROWSER_BUTTON_AUDIT.md` 는 버튼/기능 감사의 원장이다. 새 수정, 발견 문제, 검증 결과는 계속 거기에 업데이트해야 한다.

## 3. 대상 Revit 버전과 호스트 프로젝트

항상 4개 버전을 같이 본다.

```text
2019
2023
2025
2027
```

주요 호스트 프로젝트:

```text
KKY_FamilyBrowser_RevitHost_2019-2023
KKY_FamilyBrowser_RevitHost_2025
KKY_FamilyBrowser_RevitHost_2027
```

2019와 2023은 같은 `KKY_FamilyBrowser_RevitHost_2019-2023` 프로젝트 산출물을 사용한다. 2025와 2027은 별도 프로젝트다. 대시보드 관련 수정은 대부분 세 폴더에 동일하게 반영해야 한다.

주요 파일:

```text
KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDashboardHtmlForm.cs
KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserDashboardHtmlForm.cs
KKY_FamilyBrowser_RevitHost_2027\FamilyBrowserDashboardHtmlForm.cs

KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserDetachedDetailWindow.cs
KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserDetachedDetailWindow.cs
KKY_FamilyBrowser_RevitHost_2027\FamilyBrowserDetachedDetailWindow.cs

KKY_FamilyBrowser_RevitHost_2019-2023\FileGuardHtmlConfigurationForm.cs
KKY_FamilyBrowser_RevitHost_2025\FileGuardHtmlConfigurationForm.cs
KKY_FamilyBrowser_RevitHost_2027\FileGuardHtmlConfigurationForm.cs

KKY_FamilyBrowser_RevitHost_2019-2023\FamilyBrowserModernMessageDialog.cs
KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserModernMessageDialog.cs
KKY_FamilyBrowser_RevitHost_2027\FamilyBrowserModernMessageDialog.cs

KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1
```

## 4. 최근 마지막 상태

마지막으로 설치 검증까지 끝난 상태:

```text
2026-06-30 16:42 KST
```

설치 위치:

```text
C:\ProgramData\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser_RevitHost.addin
C:\ProgramData\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser_RevitHost.addin
C:\ProgramData\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser_RevitHost_2025.addin
C:\ProgramData\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser_RevitHost_2027.addin
```

마지막 검증:

```powershell
.\KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1
.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')
.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')
.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')
```

빌드는 통과했다. 경고는 기존 디컴파일/WinForms/WindowsDesktop SDK 계열 경고가 남아있다. 마지막 수정 때문에 생긴 컴파일 오류는 없었다.

## 5. 최근 중요 수정 요약

### 5.1 대시보드/선택창/파일가드 메시지 통일

사용자가 결과창, 글씨 깨짐, 한/영 전환, WinForm/MessageBox 잔존에 대해 강하게 문제 제기했다.

수정 내용:

- `FamilyBrowserModernMessageDialog.cs` 를 세 호스트 프로젝트에 추가.
- `ShowDashboardMessage` 와 `ShowDashboardChoiceMessage` 가 새 공통 모던 다이얼로그를 사용.
- `StandardFamilySelectionHtmlForm` 의 "하나 이상의 패밀리를 선택하세요" 기본 MessageBox 제거.
- `FileGuardHtmlConfigurationForm` 의 clear-all 확인 기본 MessageBox 제거.
- 정적 QA에 회귀 방지 추가.

주의:

- `FamilyBrowserDashboardHtmlForm.cs` 안에 오래된 `DashboardMessageDialog` nested class 는 아직 남아있지만, 현재 호출 경로에서는 사용하지 않는다.
- Revit 외부 명령 / native command guard 쪽 `TaskDialog.Show` 는 아직 남아있다. 이건 대시보드 내부 UI가 아니라 Revit 네이티브 명령 알림이라 별도 설계 판단이 필요하다.

### 5.2 Current Model Check 결과 export

사용자가 "검증 결과를 csv로 내보낼 필요 없다. xlsx 버튼만 두자" 라고 정정했다.

현재 의도:

- 결과창에는 `Excel 추출` 버튼 유지.
- `.xlsx` 만 저장.
- `.csv` 결과 export 는 제거.
- 단, 차이점 컬럼은 유지.

중요:

- 사용자가 말한 CSV는 결과 export CSV가 아니었다.
- 진짜 의미는 Revit family 내부 lookup CSV / size table 이다.

### 5.3 Revit family 내부 lookup CSV / size table fingerprint

사용자가 말한 CSV:

```text
.rfa 패밀리 안에서 CSV import / size table 로 타입별 치수 값을 다르게 정의하는 그 CSV
```

현재 의도:

- 표준 패밀리와 프로젝트 패밀리의 lookup table 이름, 컬럼, 행/셀 값, size definition 이 다르면 "다름"으로 인식해야 한다.

수정 내용:

- `LoadableFamilyContentSignatureService.cs` 에서 `lookup-tables=` signature 유지.
- `FamilySizeTableManager.GetFamilySizeTableManager(Document, ElementId)` 2-argument API reflection 추가.
- `familyDocument.OwnerFamily.Id` 기준 owner family id resolution 추가.
- `AsValueString` cell capture 포함.
- `ProjectStandardComparisonService.cs` 에 `lookup tables` difference classification 확인.

남은 확인:

- 실제 lookup CSV 포함 `.rfa` 로 Revit runtime check 필요.

### 5.4 패밀리/시스템 타입 상세 항목 창

의도:

- 오른쪽 inline detail 이 아니라 별도 창으로 떠야 한다.
- 패밀리 탭과 시스템 타입 탭 모두 동일.

수정된 핵심:

- `about:kkyfb:` navigation 도 action 으로 인정.
- row click 은 selection 만 수행.
- detached detail 은 toolbar `상세 항목` 버튼에서만 명시적으로 열림.
- 기존 `selectRow(this,true)` / `kkyOriginalSelectRow` 회귀 방지.

남은 확인:

- Revit UI에서 실제로 패밀리/시스템 타입 행 선택 후 `상세 항목` 버튼을 눌러 창이 뜨는지 확인.

### 5.5 필터/선택항목 로드/초기화 버튼

사용자가 버튼 무응답을 보고했다.

수정된 핵심:

- row click 에서 detail-window navigation 을 제거.
- 필터, 선택 초기화, 체크 로드/적용이 browser-only JS 또는 dashboard router 로 정상 흐르도록 복구.
- 권한 false path 가 조용히 return 하지 않고 메시지를 표시하도록 보완.

남은 확인:

- Revit UI에서 사용자가 직접 필터, 공종 tree filter, `선택 항목 로드`, `초기화` 클릭 확인.

### 5.6 표준 RVT / 표준 목록 미등록 empty state

관리자 모드일 때 Family/System Type 탭에서 표준 목록이 준비되지 않은 경우:

- `표준 패밀리 등록하기` 액션이 표시됨.
- Admin Settings 의 standard list registration 영역으로 이동.

남은 확인:

- 정밀 스캔은 되어 있고 Excel/JSON 표준 목록은 없는 상태에서 버튼이 보이는지 Revit에서 확인.

### 5.7 Permissions / Guard 탭

사용자 의도:

```text
권한/차단 탭에서는 파일별 권한 적용 설정만 사용.
역할목록 / 권한 Excel 영역은 제거.
```

현재 상태:

- Permissions / Guard 탭은 file-specific guard 중심으로 정리됨.
- role list, permission Excel, project rule, O/X diagnostic table 은 표시하지 않음.
- hidden compatibility route 는 일부 남아있을 수 있으나 visible UI 에 노출하지 않는 게 의도다.

## 6. 반드시 지켜야 할 작업 규칙

1. 수정 전 백업 생성.
2. 백업 경로를 `FAMILY_BROWSER_BUTTON_AUDIT.md` 에 기록.
3. 2019-2023 / 2025 / 2027 세 호스트 프로젝트를 같이 수정.
4. 단순 문자열 치환 말고 handler 흐름까지 추적.
5. Revit runtime 확인이 필요한 건 추측으로 Fixed 처리하지 말고 `Needs Revit Check` 로 기록.
6. 사용자가 일부러 탭 분리한 관리자 설정을 임의로 다시 합치지 말 것.
7. `권한/차단` 탭에 role list / permission Excel UI 를 되살리지 말 것.
8. Current Model Check 결과 export 에 CSV 를 되살리지 말 것.
9. "CSV 차이" 라고 하면 `.rfa` 내부 lookup CSV / size table 을 의미한다고 먼저 가정할 것.
10. 기존 dirty worktree 를 임의로 revert 하지 말 것.

## 7. 백업 위치 패턴

기존 백업:

```text
C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\_backups
C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628\artifacts\source-file-backups
```

새 백업 예시:

```powershell
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path (Resolve-Path .) "_backups\some-fix-$stamp"
New-Item -ItemType Directory -Force -Path $backup | Out-Null
```

백업 대상은 최소:

```text
FAMILY_BROWSER_BUTTON_AUDIT.md
KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1
수정할 2019-2023 파일
수정할 2025 파일
수정할 2027 파일
```

## 8. 정적 QA와 빌드 커맨드

수정 후 기본 순서:

```powershell
Set-Location 'C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628'

.\KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1

.\KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027') -Configuration Release

.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')

.\KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1 -Years @('2019','2023','2025','2027')

.\KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1 -Installed -Years @('2019','2023','2025','2027')
```

## 9. 아직 남은 리스크 / 다음에 볼 항목

### 9.1 Revit runtime 확인 필요

아래는 코드/정적 QA만으로 완전히 확정하지 말 것.

- 패밀리/시스템 타입 `상세 항목` 별도 창 실제 팝업.
- Family/System Type 탭 필터, 선택 초기화, 선택 항목 로드/적용 실제 반응.
- 정밀 스캔 후 표준 목록 미등록 상태의 `표준 패밀리 등록하기` 액션.
- lookup CSV / size table 포함 `.rfa` 의 precise fingerprint difference.
- Request store 연결 상태에서 unregistered family/system row 의 `요청` action prefill.

### 9.2 디자인/정책 판단 필요

- Revit external command / native command guard 의 `TaskDialog.Show` 를 모던 다이얼로그로 바꿀지.
- `DashboardMessageDialog` nested class 를 삭제할지. 현재 호출되지 않지만 남아있다.
- 결과창 전체를 완전 HTML-like shell 로 더 통일할지. 현재 `CurrentModelCheckResultDialog`, `FamilyLoadResultDialog` 는 custom WinForms Form 이지만 이미 크기/줄바꿈은 개선됨.

### 9.3 사용자가 특히 민감하게 본 것

- 글씨 깨짐.
- 한/영 전환 안 됨.
- 결과창/상세창 디자인 불일치.
- 버튼 눌러도 반응 없는 상태.
- CSV 의미 오해.
- 관리자 설정 탭을 의도와 다르게 합치는 것.
- 권한/차단 탭에 role list / permission Excel 을 되살리는 것.

## 10. 현재 git 상태 관련 주의

이 repo 는 이미 매우 dirty 하다. 많은 파일이 modified / untracked / deleted 로 보일 수 있다. 이는 이 대화 전후 누적 작업물이다.

새 채팅에서는 절대 아래를 하지 말 것:

```powershell
git reset --hard
git checkout -- .
```

사용자가 명시적으로 요청하지 않는 한 기존 변경을 되돌리지 않는다. 패밀리 브라우저 수정과 무관한 dirty 상태는 무시한다.

## 11. 새 채팅이 처음 실행하면 좋은 확인 명령

```powershell
Set-Location 'C:\Users\kkyki\Desktop\Codex Project\KKY_Tool_Revit_full_20260628'

Get-Content .\FAMILY_BROWSER_BUTTON_AUDIT.md -TotalCount 180

rg -n "FamilyBrowserModernMessageDialog|MessageBox\.Show|TaskDialog\.Show|DashboardMessageDialog|CurrentModelCheckResultDialog|FamilyLoadResultDialog" KKY_FamilyBrowser_RevitHost_2019-2023 KKY_FamilyBrowser_RevitHost_2025 KKY_FamilyBrowser_RevitHost_2027

.\KKY_FamilyBrowser_Compile\Test-FamilyBrowserUiStatic.ps1
```

현재 기대:

- `FamilyBrowserModernMessageDialog` 는 세 호스트 프로젝트에 있어야 한다.
- 대시보드/선택창/파일가드 직접 `MessageBox.Show` 는 없어야 한다.
- 새 공통 다이얼로그 내부 fallback `MessageBox.Show` 는 남아있을 수 있다.
- `TaskDialog.Show` 는 Revit external/native command 쪽에 남아있다.

## 12. 사용자와 대화 톤

사용자는 존대를 원하지 않는다. 결과와 현상을 논리적으로 짧게 말하는 걸 선호한다.

좋은 응답 형태:

```text
맞아. 이건 아직 덜 된 상태야.
원인은 A고, 코드상으로는 B 경로를 타고 있어.
수정은 C까지 하고, Revit 런타임 확인은 D로 남길게.
```

피해야 할 것:

- 확실하지 않은데 "완료"라고 말하기.
- 코드상 확인 없이 UI 의도 추측하기.
- 빌드/설치 검증 없이 "될 것 같다"로 마무리하기.
