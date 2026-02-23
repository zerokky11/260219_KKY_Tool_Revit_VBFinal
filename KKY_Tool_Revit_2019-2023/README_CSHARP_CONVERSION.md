# C# Conversion Bootstrap

VB 기반 `KKY_Tool_Revit` 구성과 동일한 빌드 축(2019/2021/2023/2025)을 C# 솔루션에서도 유지하도록 정리했습니다.

## 포함 내용
- `KKY_Tool_Revit_CSharp.sln` 추가
  - `KKY_Tool_Revit_2019-2023.csproj` (net48)
  - `KKY_Tool_Revit_2025.csproj` (net8.0-windows)
- `KKY_Tool_Revit_2019-2023.csproj`에 VB 프로젝트와 동일한 다중 연도 배포 타깃 추가
  - 기본 컴파일 기준: `AddinYear=2019`
  - 1회 빌드 후 `2019;2021;2023` 산출물/`addin` 배포
- `KKY_Tool_Revit_2025.csproj` 추가
  - 기존 C# 코드 링크 + 2025 전용 Compat(C#) 포함
  - 빌드 후 `ProgramData\Autodesk\Revit\Addins\2025`에 addin 파일 생성
- 2025 호환용 C# shim 추가
  - `Compat/JavaScriptSerializer.cs`
  - `Compat/BuiltInParameterGroupCompat.cs`

## 참고
- 기존 VB 프로젝트(`KKY_Tool_Revit.vbproj`, `KKY_Tool_Revit_2025.vbproj`)는 그대로 유지됩니다.
- Hub UI 리소스(HTML/CSS/JS)는 변경하지 않았습니다.
