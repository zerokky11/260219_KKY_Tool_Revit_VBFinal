# KKY_RvtBatchCleaner_CSharp

Revit 2019~2023 호환을 기준으로 만든 단일 기능 C# 프로젝트입니다.

## 기능 요약

- 다중 RVT 파일 선택
- 가능하면 `DetachAndDiscardWorksets` 로 열기
- Manage Links / CAD / Import / Image / PointCloud 삭제 시도
- Group 해제, Assembly 해제
- 기본 isometric 3D 뷰 1개 생성
- 새 3D 뷰 이름 지정
- Detail Level = Fine
- Display Style = ShadingWithEdges
- View Template 제거
- Phase Filter = Show All
- Phase = New Construction
- Starting View = 새 3D 뷰
- View 파라미터 최대 5개 입력/적용
- Model Categories 전체 표시 후 특정 카테고리 숨김
- Cable/Conduit/Duct/Pipe Fittings 서브카테고리 중 `End` 또는 `Cut` 포함 항목 숨김
- Annotation Categories 숨김
- Analytical Model Categories 숨김
- Imported Categories 표시
- 미사용 View Filter 삭제
- 필터 XML 내보내기 / 불러오기
- 필터 적용 / 미적용 / 뷰 비었을 때 자동 적용
- Purge 5회 반복 기본값
  - Revit 2024+ : `Document.GetUnusedElements()` reflection 사용
  - Revit 2019~2023 : Legacy purge 엔진 사용
- `_Detached` 제거 후 결과 폴더에 저장

## 프로젝트 구성

- `App.cs` : 리본 등록
- `Commands/CmdOpenBatchCleaner.cs` : 실행 커맨드
- `UI/BatchCleanerForm.cs` : WinForms UI
- `Services/BatchCleanService.cs` : 메인 처리 로직
- `Services/RevitViewFilterProfileService.cs` : 필터 XML / 생성 / 적용
- `Services/LegacyPurgeService.cs` : 19~23용 정리 + 24+ reflection purge

## 빌드

기본 Revit 참조 경로:

```xml
<RevitApiDir>C:\Program Files\Autodesk\Revit 2019</RevitApiDir>
```

빌드할 때 다른 버전을 쓰면 MSBuild 속성으로 덮어쓰면 됩니다.

예:

```powershell
msbuild KKY_RvtBatchCleaner_CSharp.csproj /p:Configuration=Release /p:RevitApiDir="C:\Program Files\Autodesk\Revit 2023"
```

## 주의

이 프로젝트는 실제 배치 정리 동작까지 포함했지만, Revit 모델 상태에 따라 삭제 불가 항목이 남을 수 있습니다.
특히 2019~2023의 purge 는 공식 `GetUnusedElements()` 가 없어서 Legacy purge 엔진으로 최대한 근접하게 처리합니다.

실제 배포 전 반드시 아래 시나리오를 직접 테스트해야 합니다.

1. workshared RVT 1개
2. non-workshared RVT 1개
3. CAD/Image/PointCloud/Revit Link 포함 파일
4. View Template / Filter / Text Type 많이 들어있는 파일
5. 결과 저장 후 열기 확인



## Build / Deploy

이 프로젝트는 빌드 시 `.addin` 파일과 DLL을 자동으로 `%APPDATA%\Autodesk\Revit\Addins\<RevitYear>` 경로에 배포합니다.

기본값:
- `RevitYear=2019`
- `RevitApiDir=C:\Program Files\Autodesk\Revit <RevitYear>`
- `DeployToRevitAddins=true`

예시:

```bat
msbuild KKY_RvtBatchCleaner_CSharp.csproj /t:Build /p:Configuration=Release /p:RevitYear=2019
msbuild KKY_RvtBatchCleaner_CSharp.csproj /t:Build /p:Configuration=Release /p:RevitYear=2023
msbuild KKY_RvtBatchCleaner_CSharp.csproj /t:Build /p:Configuration=Release /p:RevitApiDir="C:\Program Files\Autodesk\Revit 2023"
```

간편 빌드:

```bat
build_revit.bat
build_revit.bat 2023
```

배포를 막고 빌드만 하고 싶으면:

```bat
msbuild KKY_RvtBatchCleaner_CSharp.csproj /t:Build /p:DeployToRevitAddins=false
```
