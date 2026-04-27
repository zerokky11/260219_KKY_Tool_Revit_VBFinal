Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiProjectParameterDuplicationOptions
            Public Property Enabled As Boolean
            Public Property Scope As String = ProjectParameterDuplicationReviewService.ScopeAll
            Public Property ParameterNames As List(Of String) = New List(Of String)()
        End Class

        Private Shared _multiParameterDuplicationRows As List(Of ProjectParameterDuplicationReviewService.ReviewRow)
        Private Shared _multiParameterDuplicationFileSummaries As List(Of ProjectParameterDuplicationReviewService.FileSummary)

        Private Function ParseProjectParameterDuplication(fd As Dictionary(Of String, Object)) As MultiProjectParameterDuplicationOptions
            Dim opt As New MultiProjectParameterDuplicationOptions()
            Dim obj = GetDictValue(fd, "parameterduplication")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Scope = NormalizeProjectParameterDuplicationScope(SafeStr(GetDictValue(d, "scope")))
            opt.ParameterNames = ParseStringList(d, "parameterNames") _
                .Select(Function(x) SafeStr(x).Trim()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
            Return opt
        End Function

        Private Shared Function NormalizeProjectParameterDuplicationScope(value As String) As String
            Dim normalized As String = SafeStr(value).Trim().ToLowerInvariant()
            If String.Equals(normalized, ProjectParameterDuplicationReviewService.ScopeSelected, StringComparison.OrdinalIgnoreCase) Then
                Return ProjectParameterDuplicationReviewService.ScopeSelected
            End If
            Return ProjectParameterDuplicationReviewService.ScopeAll
        End Function

        Private Sub HandleProjectParameterDuplicationPickSharedParams(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Shared Parameter TXT (*.txt)|*.txt|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                dlg.Multiselect = False
                dlg.Title = "공유파라미터 TXT 선택"
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then
                    SendToWebAfterDialog("parameterduplication:sharedparams-picked", New With {
                        .ok = False,
                        .cancelled = True
                    })
                    Return
                End If

                Try
                    Dim parameterNames = LoadSharedParameterNamesFromTxt(dlg.FileName)
                    If parameterNames.Count = 0 Then
                        SendToWebAfterDialog("parameterduplication:sharedparams-picked", New With {
                            .ok = False,
                            .path = dlg.FileName,
                            .message = "TXT에서 공유파라미터 이름을 찾지 못했습니다."
                        })
                        Return
                    End If

                    SendToWebAfterDialog("parameterduplication:sharedparams-picked", New With {
                        .ok = True,
                        .path = dlg.FileName,
                        .parameterNames = parameterNames,
                        .count = parameterNames.Count
                    })
                Catch ex As Exception
                    SendToWebAfterDialog("parameterduplication:sharedparams-picked", New With {
                        .ok = False,
                        .path = dlg.FileName,
                        .message = ex.Message
                    })
                End Try
            End Using
        End Sub

        Private Shared Function LoadSharedParameterNamesFromTxt(filePath As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return result

            For Each rawLine In File.ReadLines(filePath)
                If String.IsNullOrWhiteSpace(rawLine) Then Continue For

                Dim parts = rawLine.Split(New String() {vbTab}, StringSplitOptions.None)
                If parts.Length < 3 Then Continue For
                If Not String.Equals(If(parts(0), "").Trim(), "PARAM", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim name As String = If(parts(2), "").Trim()
                If String.IsNullOrWhiteSpace(name) Then Continue For
                If seen.Add(name) Then result.Add(name)
            Next

            result.Sort(StringComparer.CurrentCultureIgnoreCase)
            Return result
        End Function

        Private Sub RunProjectParameterDuplicationMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.ProjectParameterDuplication Is Nothing OrElse Not _multiRequest.ProjectParameterDuplication.Enabled Then Return

            Dim settings As New ProjectParameterDuplicationReviewService.Settings With {
                .Scope = NormalizeProjectParameterDuplicationScope(_multiRequest.ProjectParameterDuplication.Scope),
                .ParameterNames = If(_multiRequest.ProjectParameterDuplication.ParameterNames, New List(Of String)())
            }

            Dim result = ProjectParameterDuplicationReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "Project Parameter 중복 검토 실행 중", $"{safeName} · {msg}")
                End Sub)

            If _multiParameterDuplicationRows Is Nothing Then _multiParameterDuplicationRows = New List(Of ProjectParameterDuplicationReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiParameterDuplicationRows.AddRange(result.Rows)
            End If

            If _multiParameterDuplicationFileSummaries Is Nothing Then _multiParameterDuplicationFileSummaries = New List(Of ProjectParameterDuplicationReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiParameterDuplicationFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiParameterDuplicationFileSummaries.Add(New ProjectParameterDuplicationReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .TotalReviewed = 0,
                    .ErrorCount = 0,
                    .OkCount = 0,
                    .Reason = "검토 결과가 없습니다."
                })
            End If
        End Sub

        Private Sub ClearMultiProjectParameterDuplicationCache()
            _multiParameterDuplicationRows = Nothing
            _multiParameterDuplicationFileSummaries = Nothing
        End Sub

        Private Function BuildProjectParameterDuplicationMultiSummary() As Object
            Dim summaries = If(_multiParameterDuplicationFileSummaries, New List(Of ProjectParameterDuplicationReviewService.FileSummary)())
            Dim totalReviewed As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TotalReviewed))
            Dim totalErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ErrorCount))
            Dim totalOk As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.OkCount))
            Dim scopeLabel As String = "전체 추가 파라미터"
            Dim selectedCount As Integer = 0

            If _multiRequest IsNot Nothing AndAlso _multiRequest.ProjectParameterDuplication IsNot Nothing Then
                Dim scope = NormalizeProjectParameterDuplicationScope(_multiRequest.ProjectParameterDuplication.Scope)
                selectedCount = If(_multiRequest.ProjectParameterDuplication.ParameterNames, New List(Of String)()).Count
                If String.Equals(scope, ProjectParameterDuplicationReviewService.ScopeSelected, StringComparison.OrdinalIgnoreCase) Then
                    scopeLabel = $"지정 파라미터 {selectedCount}개"
                End If
            End If

            Return New With {
                .key = "parameterduplication",
                .label = "Project Parameter 중복 검토",
                .lines = New String() {
                    $"선택 파일 수 {GetRequestedMultiFileCount()}개",
                    $"검토 범위: {scopeLabel}",
                    $"검토 파라미터 수 {totalReviewed}개",
                    $"중복 오류 수 {totalErrors}개",
                    $"정상 파라미터 수 {totalOk}개",
                    $"엑셀 결과 행 수 {If(_multiParameterDuplicationRows, New List(Of ProjectParameterDuplicationReviewService.ReviewRow)()).Count}개"
                },
                .fileSummaries = BuildProjectParameterDuplicationFileSummaries()
            }
        End Function

        Private Function BuildProjectParameterDuplicationFileSummaries() As List(Of Object)
            Dim summaries = If(_multiParameterDuplicationFileSummaries, New List(Of ProjectParameterDuplicationReviewService.FileSummary)())
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
                    .near = okCount,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Sub ExportProjectParameterDuplication(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiParameterDuplicationRows, New List(Of ProjectParameterDuplicationReviewService.ReviewRow)())
            Dim summaries = If(_multiParameterDuplicationFileSummaries, New List(Of ProjectParameterDuplicationReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "Project Parameter 중복 검토 결과가 없습니다."})
                Return
            End If

            Dim orderedNames = BuildOrderedMultiFileNames(
                rows.Select(Function(row) If(row Is Nothing, "", row.File)),
                summaries.Select(Function(summary) If(summary Is Nothing, "", summary.File)))

            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim defaultFileName As String
            If orderedNames.Count >= 2 OrElse requestedCount >= 2 Then
                defaultFileName = $"ProjectParameterDuplication_Selected {Math.Max(orderedNames.Count, requestedCount)} Files.xlsx"
            Else
                defaultFileName = $"ProjectParameterDuplication_{Date.Now:yyyyMMdd_HHmm}.xlsx"
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
                        baseName = IO.Path.GetFileNameWithoutExtension(fileName)
                    Catch
                    End Try
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = fileName

                    Dim table = ProjectParameterDuplicationReviewService.BuildExportTable(perFileRows)
                    ExcelCore.EnsureNoDataRow(table, "검토 결과가 없습니다.")
                    sheetList.Add(New KeyValuePair(Of String, DataTable)(baseName, table))
                    Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                    SetSplitExportIssueCount(fileIssueCounts, baseName, If(summary Is Nothing, 0, summary.ErrorCount))
                Next

                If sheetList.Count = 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "Project Parameter 중복 검토 결과가 없습니다."})
                    Return
                End If

                If Not String.IsNullOrWhiteSpace(outputFolder) Then
                    Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "parameterduplication", "ProjectParameter중복검토", "Project Parameter Duplication Review", sheetList, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                    If savedCount <= 0 Then
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                    Else
                        SendSplitExportCompleted(outputFolder, savedCount)
                    End If
                    Return
                End If

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="parameterduplication", exportKind:="parameterduplication", exportLocale:=exportLocale)
            Else
                Dim table = ProjectParameterDuplicationReviewService.BuildExportTable(rows)
                ExcelCore.EnsureNoDataRow(table, "검토 결과가 없습니다.")
                saved = ExcelCore.PickAndSaveXlsx("Project Parameter Duplication Review", table, defaultFileName, doAutoFit, "hub:multi-progress", "parameterduplication", exportLocale:=exportLocale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("parameterduplication", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

    End Class

End Namespace
