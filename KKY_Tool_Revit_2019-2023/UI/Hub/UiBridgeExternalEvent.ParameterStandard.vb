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
Imports NPOI.SS.UserModel
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiParameterStandardOptions
            Public Property Enabled As Boolean
            Public Property CriteriaExcelPath As String = String.Empty
            Public Property CriteriaRules As List(Of ParameterStandardReviewService.CriteriaRule) = New List(Of ParameterStandardReviewService.CriteriaRule)()
            Public Property CriteriaParameterCount As Integer
            Public Property CriteriaValueCount As Integer
            Public Property CriteriaSheetCount As Integer
            Public Property BlankAllowedCount As Integer
            Public Property WarningCount As Integer
        End Class

        Private NotInheritable Class ParameterStandardCriteriaSnapshot
            Public Property SourcePath As String = String.Empty
            Public Property CriteriaRules As List(Of ParameterStandardReviewService.CriteriaRule) = New List(Of ParameterStandardReviewService.CriteriaRule)()
            Public Property ParameterCount As Integer
            Public Property ValueCount As Integer
            Public Property SheetCount As Integer
            Public Property BlankAllowedCount As Integer
            Public Property Warnings As List(Of String) = New List(Of String)()
        End Class

        Private Shared _multiParameterStandardRows As List(Of ParameterStandardReviewService.ReviewRow)
        Private Shared _multiParameterStandardFileSummaries As List(Of ParameterStandardReviewService.FileSummary)
        Private Shared _multiParameterStandardWarnings As List(Of String)

        Private Sub HandleParameterStandardPickCriteria(app As UIApplication, payload As Object)
            Try
                Using dlg As New WinForms.OpenFileDialog()
                    dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
                    dlg.Title = "속성 모수 검토 기준 엑셀 선택"
                    dlg.RestoreDirectory = True
                    If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                    Dim snapshot = LoadParameterStandardCriteria(dlg.FileName)
                    SendToWebAfterDialog("parameterstandard:criteria-picked", New With {
                        .ok = True,
                        .path = dlg.FileName,
                        .parameterCount = snapshot.ParameterCount,
                        .valueCount = snapshot.ValueCount,
                        .sheetCount = snapshot.SheetCount,
                        .blankAllowedCount = snapshot.BlankAllowedCount,
                        .warningCount = snapshot.Warnings.Count,
                        .warnings = snapshot.Warnings.Take(12).ToArray()
                    })
                End Using
            Catch ex As Exception
                SendToWebAfterDialog("parameterstandard:criteria-picked", New With {
                    .ok = False,
                    .message = ex.Message
                })
            End Try
        End Sub

        Private Function ParseParameterStandard(fd As Dictionary(Of String, Object)) As MultiParameterStandardOptions
            Dim opt As New MultiParameterStandardOptions()
            Dim obj = GetDictValue(fd, "parameterstandard")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.CriteriaExcelPath = SafeStr(GetDictValue(d, "criteriaExcelPath")).Trim()
            opt.CriteriaParameterCount = ToInt(GetDictValue(d, "criteriaParameterCount"), 0)
            opt.CriteriaValueCount = ToInt(GetDictValue(d, "criteriaValueCount"), 0)
            opt.CriteriaSheetCount = ToInt(GetDictValue(d, "criteriaSheetCount"), 0)
            opt.BlankAllowedCount = ToInt(GetDictValue(d, "blankAllowedCount"), 0)
            opt.WarningCount = ToInt(GetDictValue(d, "warningCount"), 0)
            Return opt
        End Function

        Private Sub PrepareParameterStandardCriteria(req As MultiRunRequest)
            If req Is Nothing OrElse req.ParameterStandard Is Nothing OrElse Not req.ParameterStandard.Enabled Then Return

            Dim excelPath As String = SafeStr(req.ParameterStandard.CriteriaExcelPath).Trim()
            If String.IsNullOrWhiteSpace(excelPath) Then
                Throw New InvalidOperationException("속성 모수 검토 기준 엑셀 파일을 선택해 주세요.")
            End If

            Dim snapshot = LoadParameterStandardCriteria(excelPath)
            req.ParameterStandard.CriteriaExcelPath = snapshot.SourcePath
            req.ParameterStandard.CriteriaRules = snapshot.CriteriaRules
            req.ParameterStandard.CriteriaParameterCount = snapshot.ParameterCount
            req.ParameterStandard.CriteriaValueCount = snapshot.ValueCount
            req.ParameterStandard.CriteriaSheetCount = snapshot.SheetCount
            req.ParameterStandard.BlankAllowedCount = snapshot.BlankAllowedCount
            req.ParameterStandard.WarningCount = snapshot.Warnings.Count

            _multiParameterStandardWarnings = snapshot.Warnings.ToList()
        End Sub

        Private Sub RunParameterStandardMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.ParameterStandard Is Nothing OrElse Not _multiRequest.ParameterStandard.Enabled Then Return

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

            Dim settings As New ParameterStandardReviewService.Settings With {
                .CriteriaRules = If(_multiRequest.ParameterStandard.CriteriaRules, New List(Of ParameterStandardReviewService.CriteriaRule)()),
                .HasAllowedElementScope = hasAllowedElementScope,
                .AllowedElementIds = If(allowedElementIds, New List(Of Integer)()),
                .CommonTargetFilterText = commonTargetFilter,
                .CommonExcludeTargetFilterText = commonExcludeTargetFilter,
                .ExtraParameterNames = commonExtraParamNames
            }

            Dim result = ParameterStandardReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "속성 모수 검토 실행 중", $"{safeName} - {msg}")
                End Sub)

            If _multiParameterStandardRows Is Nothing Then _multiParameterStandardRows = New List(Of ParameterStandardReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiParameterStandardRows.AddRange(result.Rows)
            End If

            If _multiParameterStandardFileSummaries Is Nothing Then _multiParameterStandardFileSummaries = New List(Of ParameterStandardReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiParameterStandardFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiParameterStandardFileSummaries.Add(New ParameterStandardReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .TargetElementCount = 0,
                    .ParameterCount = 0,
                    .TotalReviewed = 0,
                    .ErrorCount = 0,
                    .OkCount = 0,
                    .BlankAllowedCount = 0,
                    .Reason = "검토 결과가 없습니다."
                })
            End If
        End Sub

        Private Sub ClearMultiParameterStandardCache()
            _multiParameterStandardRows = Nothing
            _multiParameterStandardFileSummaries = Nothing
            _multiParameterStandardWarnings = Nothing
        End Sub

        Private Function GetMultiParameterStandardRowCount() As Integer
            Return If(_multiParameterStandardRows, New List(Of ParameterStandardReviewService.ReviewRow)()).Count
        End Function

        Private Function BuildParameterStandardMultiSummary() As Object
            Dim summaries = If(_multiParameterStandardFileSummaries, New List(Of ParameterStandardReviewService.FileSummary)())
            Dim totalTargets As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TargetElementCount))
            Dim totalReviewed As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TotalReviewed))
            Dim totalErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ErrorCount))
            Dim totalOk As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.OkCount))
            Dim criteriaFileName As String = "(미선택)"
            Dim parameterCount As Integer = 0
            Dim valueCount As Integer = 0
            Dim blankAllowedCount As Integer = 0
            If _multiRequest IsNot Nothing AndAlso _multiRequest.ParameterStandard IsNot Nothing Then
                criteriaFileName = GetParameterStandardCriteriaFileLabel(_multiRequest.ParameterStandard.CriteriaExcelPath)
                parameterCount = Math.Max(_multiRequest.ParameterStandard.CriteriaParameterCount, If(_multiRequest.ParameterStandard.CriteriaRules, New List(Of ParameterStandardReviewService.CriteriaRule)()).Count)
                valueCount = _multiRequest.ParameterStandard.CriteriaValueCount
                blankAllowedCount = _multiRequest.ParameterStandard.BlankAllowedCount
            End If

            Dim commonFilterLabel As String = "없음"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(_multiRequest.Common.TargetFilter) OrElse Not String.IsNullOrWhiteSpace(_multiRequest.Common.ExcludeTargetFilter) Then
                    commonFilterLabel = "적용"
                End If
            End If

            Dim warningCount As Integer = If(_multiParameterStandardWarnings, New List(Of String)()).Count
            Return New With {
                .key = "parameterstandard",
                .label = "속성 모수 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"기준 엑셀: {criteriaFileName}",
                    $"기준 파라미터 수: {parameterCount}개",
                    $"허용값 수: {valueCount}개",
                    $"공란 허용 파라미터: {blankAllowedCount}개",
                    $"공통 검토 대상 필터: {commonFilterLabel}",
                    $"검토 대상 객체 수: {totalTargets}개",
                    $"검토 건수: {totalReviewed}건",
                    $"기준 불일치 오류: {totalErrors}건",
                    $"정상: {totalOk}건",
                    $"기준 엑셀 경고: {warningCount}건",
                    $"저장 결과 행 수: {GetMultiParameterStandardRowCount()}행"
                },
                .fileSummaries = BuildParameterStandardFileSummaries()
            }
        End Function

        Private Function BuildParameterStandardFileSummaries() As List(Of Object)
            Dim summaries = If(_multiParameterStandardFileSummaries, New List(Of ParameterStandardReviewService.FileSummary)())
            Dim orderedNames = BuildOrderedMultiFileNames(summaries.Select(Function(item) If(item Is Nothing, "", item.File)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issues As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = ""

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.TotalReviewed
                    issues = summary.ErrorCount
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

        Private Sub ExportParameterStandard(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiParameterStandardRows, New List(Of ParameterStandardReviewService.ReviewRow)())
            Dim summaries = If(_multiParameterStandardFileSummaries, New List(Of ParameterStandardReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "속성 모수 검토 결과가 없습니다."})
                Return
            End If

            Dim orderedNames = BuildOrderedMultiFileNames(
                rows.Select(Function(row) If(row Is Nothing, "", row.File)),
                summaries.Select(Function(summary) If(summary Is Nothing, "", summary.File)))

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each fileName In orderedNames
                Dim fileRows = rows.Where(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase)).ToList()
                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                Dim table = ParameterStandardReviewService.BuildExportTable(fileRows)
                Dim sheetName = BuildParameterStandardSheetName(fileName)
                sheets.Add(New KeyValuePair(Of String, DataTable)(sheetName, table))
                SetSplitExportIssueCount(fileIssueCounts, sheetName, If(summary Is Nothing, 0, summary.ErrorCount))
            Next

            If sheets.Count = 0 Then
                sheets.Add(New KeyValuePair(Of String, DataTable)("Review", ParameterStandardReviewService.BuildExportTable(rows)))
            End If

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "parameterstandard", "ParameterStandardReview", "Review", sheets, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(
                sheets,
                BuildParameterStandardDefaultExcelName(),
                doAutoFit,
                "hub:multi-progress",
                sheetKeyOverride:="parameterstandard",
                exportKind:="parameterstandard",
                exportLocale:=exportLocale)

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("parameterstandard", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Function LoadParameterStandardCriteria(path As String) As ParameterStandardCriteriaSnapshot
            Dim safePath As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(safePath) Then
                Throw New InvalidOperationException("속성 모수 검토 기준 엑셀 파일 경로가 비어 있습니다.")
            End If
            If Not File.Exists(safePath) Then
                Throw New FileNotFoundException("기준 엑셀 파일을 찾을 수 없습니다.", safePath)
            End If

            Dim snapshot As New ParameterStandardCriteriaSnapshot With {
                .SourcePath = safePath
            }
            Dim uniqueRules As New Dictionary(Of String, ParameterStandardReviewService.CriteriaRule)(StringComparer.OrdinalIgnoreCase)
            Dim formatter As New DataFormatter()

            Using stream As New FileStream(safePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim workbook = WorkbookFactory.Create(stream)
                For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                    Dim sheet = workbook.GetSheetAt(sheetIndex)
                    If sheet Is Nothing Then Continue For

                    Dim parameterName As String = CleanParameterStandardText(sheet.SheetName)
                    If String.IsNullOrWhiteSpace(parameterName) Then Continue For

                    snapshot.SheetCount += 1
                    Dim headerParameterName As String = GetParameterStandardCellText(sheet, 0, 1, formatter)
                    If Not String.IsNullOrWhiteSpace(headerParameterName) AndAlso Not String.Equals(parameterName, headerParameterName, StringComparison.OrdinalIgnoreCase) Then
                        snapshot.Warnings.Add($"{sheet.SheetName}: 시트명({parameterName})과 B1 파라미터명({headerParameterName})이 다릅니다. 시트명을 기준으로 검토합니다.")
                    End If

                    Dim allowedValues As New List(Of String)()
                    Dim seenValues As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    Dim allowBlank As Boolean = False

                    For rowIndex As Integer = 1 To sheet.LastRowNum
                        Dim value As String = GetParameterStandardCellText(sheet, rowIndex, 1, formatter)
                        If String.IsNullOrWhiteSpace(value) Then Continue For

                        If IsParameterStandardBlankToken(value) Then
                            allowBlank = True
                            Continue For
                        End If

                        If seenValues.Add(value) Then
                            allowedValues.Add(value)
                        End If
                    Next

                    If allowedValues.Count = 0 AndAlso Not allowBlank Then
                        snapshot.Warnings.Add($"{sheet.SheetName}: B2 아래 기준값이 없어 이 시트는 건너뜁니다.")
                        Continue For
                    End If

                    If uniqueRules.ContainsKey(parameterName) Then
                        snapshot.Warnings.Add($"{sheet.SheetName}: 같은 파라미터 기준 시트가 이미 있어 중복 시트는 건너뜁니다.")
                        Continue For
                    End If

                    uniqueRules.Add(parameterName, New ParameterStandardReviewService.CriteriaRule With {
                        .ParameterName = parameterName,
                        .SheetName = sheet.SheetName,
                        .HeaderParameterName = headerParameterName,
                        .AllowedValues = allowedValues,
                        .AllowBlank = allowBlank
                    })
                Next
            End Using

            snapshot.CriteriaRules = uniqueRules.Values _
                .OrderBy(Function(item) item.ParameterName, StringComparer.OrdinalIgnoreCase) _
                .ToList()
            snapshot.ParameterCount = snapshot.CriteriaRules.Count
            snapshot.ValueCount = snapshot.CriteriaRules.Sum(Function(item) If(item Is Nothing OrElse item.AllowedValues Is Nothing, 0, item.AllowedValues.Count))
            snapshot.BlankAllowedCount = snapshot.CriteriaRules.Where(Function(item) item IsNot Nothing AndAlso item.AllowBlank).Count()

            If snapshot.ParameterCount = 0 Then
                Throw New InvalidOperationException("기준 엑셀에서 사용할 파라미터 기준값을 찾지 못했습니다. 각 시트명은 파라미터명, B2 아래는 허용값이어야 합니다.")
            End If

            Return snapshot
        End Function

        Private Function GetParameterStandardCellText(sheet As ISheet, rowIndex As Integer, columnIndex As Integer, formatter As DataFormatter) As String
            If sheet Is Nothing OrElse rowIndex < 0 OrElse columnIndex < 0 Then Return String.Empty
            Dim row = sheet.GetRow(rowIndex)
            If row Is Nothing Then Return String.Empty
            Dim cell = row.GetCell(columnIndex)
            If cell Is Nothing Then Return String.Empty
            Return CleanParameterStandardText(formatter.FormatCellValue(cell))
        End Function

        Private Function CleanParameterStandardText(value As String) As String
            Dim text As String = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(text) Then Return String.Empty
            text = text.Trim()
            text = text.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ").Replace(ControlChars.Tab, " ")
            Return String.Join(" ", text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))
        End Function

        Private Function IsParameterStandardBlankToken(value As String) As Boolean
            Dim text As String = CleanParameterStandardText(value)
            If String.IsNullOrWhiteSpace(text) Then Return False
            Dim legacyBlankWithParens As String = String.Concat("(", ChrW(&H6028), ChrW(&HB4EC), "?)")
            Dim legacyBlankPlain As String = String.Concat(ChrW(&H6028), ChrW(&HB4EC), "?")
            Return String.Equals(text, "(공란)", StringComparison.OrdinalIgnoreCase) _
                OrElse String.Equals(text, "공란", StringComparison.OrdinalIgnoreCase) _
                OrElse String.Equals(text, legacyBlankWithParens, StringComparison.OrdinalIgnoreCase) _
                OrElse String.Equals(text, legacyBlankPlain, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function BuildParameterStandardDefaultExcelName() As String
            Dim baseName As String = String.Empty
            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing AndAlso _multiRequest.RvtPaths.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiRequest.RvtPaths(0)))
            ElseIf _multiParameterStandardFileSummaries IsNot Nothing AndAlso _multiParameterStandardFileSummaries.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiParameterStandardFileSummaries(0).File))
            End If

            If String.IsNullOrWhiteSpace(baseName) Then
                Return $"ParameterStandardReview_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Return $"{baseName}_ParameterStandardReview.xlsx"
        End Function

        Private Function BuildParameterStandardSheetName(fileName As String) As String
            Dim safeName As String = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(fileName))
            If String.IsNullOrWhiteSpace(safeName) Then safeName = "Review"
            Return safeName
        End Function

        Private Function GetParameterStandardCriteriaFileLabel(path As String) As String
            Dim safePath As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(safePath) Then Return "(미선택)"
            Try
                Return System.IO.Path.GetFileName(safePath)
            Catch
                Return safePath
            End Try
        End Function

    End Class

End Namespace
