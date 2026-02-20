# C# Conversion Bootstrap

이 커밋은 VB 기반 `KKY_Tool_Revit_2019-2023`의 C# 전환을 위한 시작점을 추가합니다.

## 포함 내용
- `KKY_Tool_Revit_2019-2023.csproj` 추가 (SDK 스타일, .NET Framework 4.8, WinForms/WPF 활성화)
- Shared Parameter 상태/파서 C# 구현
- GUID Audit 서비스 C# 구현(프로젝트/패밀리 기본 감사 + 고정 DataTable 스키마)
- `UiBridgeExternalEvent`의 GUID 이벤트 맵 C# 구현

## 참고
- 기존 VB 프로젝트(`KKY_Tool_Revit.vbproj`) 및 리소스는 그대로 유지됩니다.
- Hub UI 리소스(HTML/CSS/JS)는 변경하지 않았습니다.
