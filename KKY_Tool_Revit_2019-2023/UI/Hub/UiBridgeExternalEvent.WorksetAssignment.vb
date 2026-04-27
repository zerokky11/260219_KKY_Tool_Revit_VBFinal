Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Autodesk.Revit.DB
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiWorksetAssignmentOptions
            Public Property Enabled As Boolean
            Public Property ExpectedWorksetName As String = WorksetAssignmentReviewService.DefaultExpectedWorksetName
            Public Property FlaggedWorksetName As String = String.Empty
        End Class

        Private Shared _multiWorksetAssignmentRows As List(Of WorksetAssignmentReviewService.ReviewRow)
        Private Shared _multiWorksetAssignmentFileSummaries As List(Of WorksetAssignmentReviewService.FileSummary)

        Private Function ParseWorksetAssignment(fd As Dictionary(Of String, Object)) As MultiWorksetAssignmentOptions
            Dim opt As New MultiWorksetAssignmentOptions()
            Dim obj = GetDictValue(fd, "worksetassignment")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.ExpectedWorksetName = SafeStr(GetDictValue(d, "expectedWorksetName")).Trim()
            opt.FlaggedWorksetName = SafeStr(GetDictValue(d, "flaggedWorksetName")).Trim()
            If String.IsNullOrWhiteSpace(opt.ExpectedWorksetName) Then
                opt.ExpectedWorksetName = WorksetAssignmentReviewService.DefaultExpectedWorksetName
            End If
            Return opt
        End Function

        Private Sub RunWorksetAssignmentMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.WorksetAssignment Is Nothing OrElse Not _multiRequest.WorksetAssignment.Enabled Then Return

            Dim commonTargetFilter As String = String.Empty
            Dim commonExcludeTargetFilter As String = String.Empty
            Dim commonExtraParamNames As New List(Of String)()
            Dim allowedElementIds As New List(Of Integer)()
            Dim hasAllowedElementScope As Boolean = False
            If _multiRequest.Common IsNot Nothing Then
                commonTargetFilter = SafeStr(_multiRequest.Common.TargetFilter)
                commonExcludeTargetFilter = SafeStr(_multiRequest.Common.ExcludeTargetFilter)
                commonExtraParamNames = ParseExtraParams(SafeStr(_multiRequest.Common.ExtraParams))
                hasAllowedElementScope = TryBuildCommonScopeIds(doc, commonTargetFilter, commonExcludeTargetFilter, allowedElementIds)
            End If

            Dim settings As New WorksetAssignmentReviewService.Settings With {
                .ExpectedWorksetName = If(_multiRequest.WorksetAssignment.ExpectedWorksetName, WorksetAssignmentReviewService.DefaultExpectedWorksetName),
                .FlaggedWorksetName = If(_multiRequest.WorksetAssignment.FlaggedWorksetName, String.Empty),
                .HasAllowedElementScope = hasAllowedElementScope,
                .AllowedElementIds = allowedElementIds,
                .ExtraParameterNames = commonExtraParamNames
            }

            Dim result = WorksetAssignmentReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "웍셋 배정 검토 실행 중", $"{safeName} · {msg}")
                End Sub)

            If _multiWorksetAssignmentRows Is Nothing Then _multiWorksetAssignmentRows = New List(Of WorksetAssignmentReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiWorksetAssignmentRows.AddRange(result.Rows)
            End If

            If _multiWorksetAssignmentFileSummaries Is Nothing Then _multiWorksetAssignmentFileSummaries = New List(Of WorksetAssignmentReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiWorksetAssignmentFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiWorksetAssignmentFileSummaries.Add(New WorksetAssignmentReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .TotalReviewed = 0,
                    .ErrorCount = 0,
                    .OkCount = 0,
                    .Reason = "검토 결과 없음"
                })
            End If
        End Sub

        Private Sub ClearMultiWorksetAssignmentCache()
            _multiWorksetAssignmentRows = Nothing
            _multiWorksetAssignmentFileSummaries = Nothing
        End Sub

        Private Function BuildWorksetAssignmentMultiSummary() As Object
            Dim summaries = If(_multiWorksetAssignmentFileSummaries, New List(Of WorksetAssignmentReviewService.FileSummary)())
            Dim totalReviewed As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TotalReviewed))
            Dim totalErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ErrorCount))
            Dim totalOk As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.OkCount))
            Dim expectedWorksetName As String = WorksetAssignmentReviewService.DefaultExpectedWorksetName

            If _multiRequest IsNot Nothing AndAlso
               _multiRequest.WorksetAssignment IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(_multiRequest.WorksetAssignment.ExpectedWorksetName) Then
                expectedWorksetName = _multiRequest.WorksetAssignment.ExpectedWorksetName.Trim()
            End If

            Dim flaggedWorksetName As String = ""
            If _multiRequest IsNot Nothing AndAlso
               _multiRequest.WorksetAssignment IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(_multiRequest.WorksetAssignment.FlaggedWorksetName) Then
                flaggedWorksetName = _multiRequest.WorksetAssignment.FlaggedWorksetName.Trim()
            End If

            Return New With {
                .key = "worksetassignment",
                .label = "웍셋 배정 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"기준 workset: {expectedWorksetName}",
                    If(String.IsNullOrWhiteSpace(flaggedWorksetName), "오류 대상 workset: (미입력)", $"오류 대상 workset: {flaggedWorksetName}"),
                    $"검토 객체 수: {totalReviewed}개",
                    $"오류 객체 수: {totalErrors}개",
                    $"정상 객체 수: {totalOk}개",
                    $"엑셀 결과 행 수: {If(_multiWorksetAssignmentRows, New List(Of WorksetAssignmentReviewService.ReviewRow)()).Count}행"
                },
                .fileSummaries = BuildWorksetAssignmentFileSummaries()
            }
        End Function

        Private Function BuildWorksetAssignmentFileSummaries() As List(Of Object)
            Dim summaries = If(_multiWorksetAssignmentFileSummaries, New List(Of WorksetAssignmentReviewService.FileSummary)())
            Dim orderedNames = BuildOrderedMultiFileNames(summaries.Select(Function(item) If(item Is Nothing, "", item.File)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issues As Integer = 0
                Dim okCount As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = ""

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.TotalReviewed
                    issues = summary.ErrorCount
                    okCount = summary.OkCount
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

        Private Sub ExportWorksetAssignment(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiWorksetAssignmentRows, New List(Of WorksetAssignmentReviewService.ReviewRow)())
            Dim summaries = If(_multiWorksetAssignmentFileSummaries, New List(Of WorksetAssignmentReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "웍셋 배정 검토 결과가 없습니다."})
                Return
            End If

            Dim orderedNames = BuildOrderedMultiFileNames(
                rows.Select(Function(row) If(row Is Nothing, "", row.File)),
                summaries.Select(Function(summary) If(summary Is Nothing, "", summary.File)))

            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim defaultFileName As String
            If orderedNames.Count >= 2 OrElse requestedCount >= 2 Then
                defaultFileName = $"WorksetAssignment_Selected {Math.Max(orderedNames.Count, requestedCount)} Files.xlsx"
            Else
                defaultFileName = $"WorksetAssignment_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Dim saved As String = ""
            If Not String.IsNullOrWhiteSpace(outputFolder) OrElse orderedNames.Count >= 2 Then
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                For Each fileName In orderedNames
                    Dim perFileRows = rows.
                        Where(Function(row) row IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(row.File), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    If perFileRows.Count = 0 Then Continue For

                    Dim baseName As String = fileName
                    Try
                        baseName = System.IO.Path.GetFileNameWithoutExtension(fileName)
                    Catch
                    End Try
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = fileName

                    Dim table = WorksetAssignmentReviewService.BuildExportTable(perFileRows)
                    ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
                    sheetList.Add(New KeyValuePair(Of String, DataTable)(baseName, table))
                    Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                    SetSplitExportIssueCount(fileIssueCounts, baseName, If(summary Is Nothing, perFileRows.Count, summary.ErrorCount))
                Next

                If sheetList.Count = 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "웍셋 배정 검토 결과가 없습니다."})
                    Return
                End If

                If Not String.IsNullOrWhiteSpace(outputFolder) Then
                    Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "worksetassignment", "웍셋배정검토", "Workset Assignment Review", sheetList, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                    If savedCount <= 0 Then
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                    Else
                        SendSplitExportCompleted(outputFolder, savedCount)
                    End If
                    Return
                End If

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="worksetassignment", exportKind:="worksetassignment", exportLocale:=exportLocale)
            Else
                Dim table = WorksetAssignmentReviewService.BuildExportTable(rows)
                ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
                saved = ExcelCore.PickAndSaveXlsx("Workset Assignment Review", table, defaultFileName, doAutoFit, "hub:multi-progress", "worksetassignment", exportLocale:=exportLocale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("worksetassignment", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

    End Class

End Namespace
