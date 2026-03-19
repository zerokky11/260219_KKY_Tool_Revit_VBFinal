Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiFloorInfoOptions
            Public Property Enabled As Boolean
            Public Property ParameterName As String = String.Empty
            Public Property BaseLevelName As String = String.Empty
            Public Property LevelRules As List(Of FloorInfoReviewService.LevelRule) = New List(Of FloorInfoReviewService.LevelRule)()
        End Class

        Private Shared _multiFloorInfoRows As List(Of FloorInfoReviewService.ReviewRow)
        Private Shared _multiFloorInfoFileSummaries As List(Of FloorInfoReviewService.FileSummary)
        Private Shared _multiFloorInfoWarnings As List(Of String)

        Private Sub HandleFloorInfoConfigLoad(app As UIApplication, payload As Object)
            Try
                Dim uidoc = If(app Is Nothing, Nothing, app.ActiveUIDocument)
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    SendToWeb("floorinfo:config-loaded", New With {
                        .ok = False,
                        .message = "활성 문서가 없습니다.",
                        .levels = New List(Of Object)(),
                        .warnings = New List(Of String)()
                    })
                    Return
                End If

                Dim snapshot = FloorInfoReviewService.ReadConfig(doc)
                Dim levels = If(snapshot.Levels, New List(Of FloorInfoReviewService.LevelOption)()) _
                    .Select(Function(level) New With {
                        .levelId = level.LevelId,
                        .levelName = level.LevelName,
                        .absoluteZFt = level.AbsoluteZFt,
                        .absoluteZMm = level.AbsoluteZMm,
                        .isBaseLevel = level.IsBaseLevel
                    }).ToList()

                SendToWeb("floorinfo:config-loaded", New With {
                    .ok = True,
                    .documentTitle = snapshot.DocumentTitle,
                    .levels = levels,
                    .warnings = If(snapshot.Warnings, New List(Of String)())
                })
            Catch ex As Exception
                SendToWeb("floorinfo:config-loaded", New With {
                    .ok = False,
                    .message = ex.Message,
                    .levels = New List(Of Object)(),
                    .warnings = New List(Of String)()
                })
                SendToWeb("revit:error", New With {.message = "층정보 설정 로드 실패: " & ex.Message})
            End Try
        End Sub

        Private Function ParseFloorInfo(fd As Dictionary(Of String, Object)) As MultiFloorInfoOptions
            Dim opt As New MultiFloorInfoOptions()
            Dim obj = GetDictValue(fd, "floorinfo")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.ParameterName = SafeStr(GetDictValue(d, "parameterName"))
            opt.BaseLevelName = SafeStr(GetDictValue(d, "baseLevelName"))

            Dim rawRules = GetDictValue(d, "levelRules")
            Dim arr = TryCast(rawRules, System.Collections.IEnumerable)
            If arr IsNot Nothing AndAlso Not TypeOf rawRules Is String Then
                For Each item In arr
                    Dim ruleDict = ToDict(item)
                    Dim levelName = SafeStr(GetDictValue(ruleDict, "levelName"))
                    If String.IsNullOrWhiteSpace(levelName) Then Continue For
                    opt.LevelRules.Add(New FloorInfoReviewService.LevelRule With {
                        .LevelName = levelName,
                        .ExpectedValue = SafeStr(GetDictValue(ruleDict, "expectedValue")),
                        .AbsoluteZFt = ToDouble(GetDictValue(ruleDict, "absoluteZFt"), 0.0R)
                    })
                Next
            End If

            If String.IsNullOrWhiteSpace(opt.BaseLevelName) AndAlso opt.LevelRules.Count > 0 Then
                opt.BaseLevelName = opt.LevelRules(0).LevelName
            End If

            Return opt
        End Function

        Private Sub RunFloorInfoMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.FloorInfo Is Nothing OrElse Not _multiRequest.FloorInfo.Enabled Then Return

            Dim settings As New FloorInfoReviewService.Settings With {
                .ParameterName = If(_multiRequest.FloorInfo.ParameterName, String.Empty),
                .BaseLevelName = If(_multiRequest.FloorInfo.BaseLevelName, String.Empty),
                .LevelRules = If(_multiRequest.FloorInfo.LevelRules, New List(Of FloorInfoReviewService.LevelRule)())
            }

            Dim result = FloorInfoReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "층정보 Z 검토 실행 중", $"{safeName} · {msg}")
                End Sub)

            If _multiFloorInfoRows Is Nothing Then _multiFloorInfoRows = New List(Of FloorInfoReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiFloorInfoRows.AddRange(result.Rows)
            End If

            If _multiFloorInfoFileSummaries Is Nothing Then _multiFloorInfoFileSummaries = New List(Of FloorInfoReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiFloorInfoFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiFloorInfoFileSummaries.Add(New FloorInfoReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .Total = 0,
                    .Issues = 0,
                    .Near = 0
                })
            End If

            If _multiFloorInfoWarnings Is Nothing Then _multiFloorInfoWarnings = New List(Of String)()
            If result IsNot Nothing AndAlso result.Warnings IsNot Nothing Then
                For Each warning In result.Warnings
                    If String.IsNullOrWhiteSpace(warning) Then Continue For
                    _multiFloorInfoWarnings.Add($"{safeName}: {warning}")
                Next
            End If
        End Sub

        Private Sub ClearMultiFloorInfoCache()
            _multiFloorInfoRows = Nothing
            _multiFloorInfoFileSummaries = Nothing
            _multiFloorInfoWarnings = Nothing
        End Sub

        Private Function GetMultiFloorInfoRowCount() As Integer
            Return If(_multiFloorInfoRows, New List(Of FloorInfoReviewService.ReviewRow)()).Count
        End Function

        Private Function BuildFloorInfoMultiSummary() As Object
            Dim rows = If(_multiFloorInfoRows, New List(Of FloorInfoReviewService.ReviewRow)())
            Dim fileSummaries = BuildFloorInfoFileSummaries()
            Dim totalEvaluated As Integer = 0
            Dim totalIssues As Integer = 0

            For Each item In If(_multiFloorInfoFileSummaries, New List(Of FloorInfoReviewService.FileSummary)())
                totalEvaluated += item.Total
                totalIssues += item.Issues
            Next

            Dim parameterName As String = ""
            Dim baseLevelName As String = ""
            Dim ruleCount As Integer = 0
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FloorInfo IsNot Nothing Then
                parameterName = If(_multiRequest.FloorInfo.ParameterName, "")
                baseLevelName = If(_multiRequest.FloorInfo.BaseLevelName, "")
                ruleCount = If(_multiRequest.FloorInfo.LevelRules, New List(Of FloorInfoReviewService.LevelRule)()).Count
            End If

            Dim warningCount As Integer = If(_multiFloorInfoWarnings, New List(Of String)()).Count
            Return New With {
                .key = "floorinfo",
                .label = "층정보 파라미터 Z 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"검토 파라미터: {If(String.IsNullOrWhiteSpace(parameterName), "(미설정)", parameterName)}",
                    $"주 레벨: {If(String.IsNullOrWhiteSpace(baseLevelName), "(미설정)", baseLevelName)}",
                    $"레벨 규칙 수: {ruleCount}개",
                    $"평가된 객체 수: {totalEvaluated}개",
                    $"이슈 행 수: {rows.Count}행",
                    $"파일별 이슈 수 합계: {totalIssues}건",
                    $"경고 수: {warningCount}건"
                },
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildFloorInfoFileSummaries() As List(Of Object)
            Dim summaries = If(_multiFloorInfoFileSummaries, New List(Of FloorInfoReviewService.FileSummary)())
            Dim orderedNames As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each path In _multiRequest.RvtPaths
                    Dim safeName As String = GetSafeMultiFileName(path)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            If _multiRunItems IsNot Nothing Then
                For Each item In _multiRunItems
                    If item Is Nothing Then Continue For
                    Dim safeName As String = GetSafeMultiFileName(item.File)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            For Each item In summaries
                If item Is Nothing Then Continue For
                Dim safeName As String = GetSafeMultiFileName(item.File)
                If String.IsNullOrWhiteSpace(safeName) Then Continue For
                If seen.Add(safeName) Then orderedNames.Add(safeName)
            Next

            Dim result As New List(Of Object)()
            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issues As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = ""

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.Total
                    issues = summary.Issues
                    statusText = If(String.IsNullOrWhiteSpace(summary.Status), "success", summary.Status)
                    reason = If(summary.Reason, "")
                End If

                If _multiRunItems IsNot Nothing Then
                    For Each runItem In _multiRunItems
                        If runItem Is Nothing Then Continue For
                        If Not String.Equals(GetSafeMultiFileName(runItem.File), fileName, StringComparison.OrdinalIgnoreCase) Then Continue For
                        If Not String.IsNullOrWhiteSpace(runItem.Status) Then statusText = runItem.Status
                        If Not String.IsNullOrWhiteSpace(runItem.Reason) Then reason = runItem.Reason
                        Exit For
                    Next
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = total,
                    .issues = issues,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Sub ExportFloorInfo(doAutoFit As Boolean, excelMode As String)
            Dim rows = If(_multiFloorInfoRows, New List(Of FloorInfoReviewService.ReviewRow)())
            Dim summaries = If(_multiFloorInfoFileSummaries, New List(Of FloorInfoReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "층정보 검토 결과가 없습니다."})
                Return
            End If

            Dim table As DataTable = FloorInfoReviewService.BuildExportTable(rows)
            ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
            Dim saved = ExcelCore.PickAndSaveXlsx("FloorInfo Review", table, $"FloorInfoReview_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", "floorinfo")

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("floorinfo", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

    End Class

End Namespace
