# CS/VB Parity Quick Report

자동 비교(라인 수, 메서드 수)로 C# 전환 누락 의심 파일을 식별한 리포트입니다.

| File | VB lines | CS lines | VB methods | CS methods |
|---|---:|---:|---:|---:|
| `APP.vb` | 131 | 152 | 7 | 7 |
| `Commands/CmdOpenHub.vb` | 33 | 25 | 1 | 1 |
| `Exports/ConnectorExport.vb` | 47 | 55 | 3 | 3 |
| `Exports/DuplicateExport.vb` | 167 | 196 | 11 | 11 |
| `Exports/FamilyLinkAuditExport.vb` | 151 | 169 | 7 | 7 |
| `Exports/PointsExport.vb` | 29 | 29 | 2 | 2 |
| `Infrastructure/ElementIdCompat.vb` | 55 | 35 | 3 | 3 |
| `Infrastructure/ExcelCore.vb` | 831 | 956 | 30 | 30 |
| `Infrastructure/ExcelExportStyleRegistry.vb` | 278 | 298 | 20 | 20 |
| `Infrastructure/ExcelStyleHelper.vb` | 133 | 125 | 9 | 9 |
| `Infrastructure/ResourceExtractor.vb` | 106 | 118 | 3 | 2 |
| `Infrastructure/ResultTableFilter.vb` | 49 | 47 | 2 | 2 |
| `My Project/Application.Designer.vb` | 13 | 9 | 0 | 0 |
| `My Project/AssemblyInfo.vb` | 32 | 15 | 0 | 0 |
| `My Project/Resources.Designer.vb` | 62 | 40 | 0 | 0 |
| `My Project/Settings.Designer.vb` | 73 | 33 | 1 | 0 |
| `Services/ConnectorDiagnosticsService.vb` | 1301 | 569 | 48 | 26 |
| `Services/DuplicateAnalysisService.vb` | 234 | 230 | 7 | 7 |
| `Services/ExportPointsService.vb` | 335 | 366 | 17 | 17 |
| `Services/FamilyLinkAuditService.vb` | 565 | 634 | 16 | 15 |
| `Services/GuidAuditService.vb` | 1261 | 1068 | 36 | 35 |
| `Services/HubCommonOptionsStorageService.vb` | 67 | 77 | 4 | 3 |
| `Services/ParamPropagateService.vb` | 1975 | 137 | 49 | 4 |
| `Services/SegmentPmsCheckService.vb` | 2623 | 241 | 82 | 13 |
| `Services/SharedParamBatchService.vb` | 1614 | 476 | 54 | 15 |
| `Services/SharedParameterStatusService.vb` | 166 | 187 | 4 | 3 |
| `UI/Hub/ExcelProgressReporter.vb` | 84 | 97 | 3 | 3 |
| `UI/Hub/HubHostWindow.vb` | 296 | 390 | 15 | 15 |
| `UI/Hub/UiBridgeExternalEvent.Connector.vb` | 1189 | 254 | 52 | 18 |
| `UI/Hub/UiBridgeExternalEvent.Core.vb` | 430 | 570 | 17 | 21 |
| `UI/Hub/UiBridgeExternalEvent.Duplicate.vb` | 978 | 754 | 23 | 20 |
| `UI/Hub/UiBridgeExternalEvent.Export.vb` | 583 | 683 | 31 | 31 |
| `UI/Hub/UiBridgeExternalEvent.FamilyLinkAudit.vb` | 337 | 408 | 12 | 13 |
| `UI/Hub/UiBridgeExternalEvent.Guid.vb` | 369 | 342 | 12 | 11 |
| `UI/Hub/UiBridgeExternalEvent.Multi.vb` | 1075 | 156 | 50 | 5 |
| `UI/Hub/UiBridgeExternalEvent.ParamProp.vb` | 210 | 237 | 6 | 6 |
| `UI/Hub/UiBridgeExternalEvent.SegmentPms.vb` | 1014 | 190 | 34 | 18 |
| `UI/Hub/UiBridgeExternalEvent.SharedParamBatch.vb` | 214 | 286 | 7 | 8 |

## Flagged (CS lines < 60% of VB lines)

- `My Project/AssemblyInfo.vb` (VB 32 / CS 15)
- `My Project/Settings.Designer.vb` (VB 73 / CS 33)
- `Services/ConnectorDiagnosticsService.vb` (VB 1301 / CS 569)
- `Services/ParamPropagateService.vb` (VB 1975 / CS 137)
- `Services/SegmentPmsCheckService.vb` (VB 2623 / CS 241)
- `Services/SharedParamBatchService.vb` (VB 1614 / CS 476)
- `UI/Hub/UiBridgeExternalEvent.Connector.vb` (VB 1189 / CS 254)
- `UI/Hub/UiBridgeExternalEvent.Multi.vb` (VB 1075 / CS 156)
- `UI/Hub/UiBridgeExternalEvent.SegmentPms.vb` (VB 1014 / CS 190)

## Update Log

- 2026-02-20: `Infrastructure/ElementIdCompat.cs`를 VB 조건부 분기와 동일하게 수정 (`REVIT2025` 분기 반영).
- 2026-02-20: `UI/Hub/UiBridgeExternalEvent.Guid.cs`에서 GUID Export 미연결 placeholder를 제거하고 VB와 유사한 실행/내보내기/상세조회 플로우로 확장.
- 2026-02-20: `Services/GuidAuditService.cs`를 대폭 확장하여 VB의 다중 문서 오픈/감사/실패처리/Family 감사 로직을 반영.

## Remaining High-Gap Targets (next)

- `Services/ParamPropagateService.cs`
- `Services/SegmentPmsCheckService.cs`
- `Services/SharedParamBatchService.cs`
- `UI/Hub/UiBridgeExternalEvent.Multi.cs`
- `UI/Hub/UiBridgeExternalEvent.Connector.cs`
- `UI/Hub/UiBridgeExternalEvent.SegmentPms.cs`
