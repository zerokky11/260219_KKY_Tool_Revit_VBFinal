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

        Private Class MultiFamilySuitabilityOptions
            Public Property Enabled As Boolean
            Public Property CriteriaExcelPath As String = String.Empty
            Public Property CriteriaRules As List(Of FamilySuitabilityReviewService.CriteriaRule) = New List(Of FamilySuitabilityReviewService.CriteriaRule)()
            Public Property CriteriaRowCount As Integer
            Public Property CriteriaUniqueCount As Integer
            Public Property CriteriaSheetCount As Integer
            Public Property MatchReviewText As String = String.Empty
            Public Property MismatchReviewText As String = String.Empty
            Public Property FilterRules As List(Of FamilySuitabilityReviewService.FilterRule) = New List(Of FamilySuitabilityReviewService.FilterRule)()
        End Class

        Private NotInheritable Class FamilySuitabilityCriteriaSnapshot
            Public Property SourcePath As String = String.Empty
            Public Property CriteriaRules As List(Of FamilySuitabilityReviewService.CriteriaRule) = New List(Of FamilySuitabilityReviewService.CriteriaRule)()
            Public Property RowCount As Integer
            Public Property UniqueCount As Integer
            Public Property SheetCount As Integer
        End Class

        Private NotInheritable Class FamilySuitabilityHeaderMap
            Public Property HeaderRowIndex As Integer
            Public Property CategoryIndex As Integer = -1
            Public Property FamilyIndex As Integer = -1
            Public Property TypeIndex As Integer = -1
        End Class

        Private Shared _multiFamilySuitabilityRows As List(Of FamilySuitabilityReviewService.ReviewRow)
        Private Shared _multiFamilySuitabilityFileSummaries As List(Of FamilySuitabilityReviewService.FileSummary)
        Private Shared _multiFamilySuitabilityWarnings As List(Of String)

        Private Sub HandleFamilySuitabilityPickCriteria(app As UIApplication, payload As Object)
            Try
                Using dlg As New WinForms.OpenFileDialog()
                    dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
                    dlg.Title = "Family 적합성 기준 엑셀 선택"
                    dlg.RestoreDirectory = True
                    If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                    Dim snapshot = LoadFamilySuitabilityCriteria(dlg.FileName)
                    SendToWebAfterDialog("familysuitability:criteria-picked", New With {
                        .ok = True,
                        .path = dlg.FileName,
                        .rowCount = snapshot.RowCount,
                        .uniqueCount = snapshot.UniqueCount,
                        .sheetCount = snapshot.SheetCount
                    })
                End Using
            Catch ex As Exception
                SendToWebAfterDialog("familysuitability:criteria-picked", New With {
                    .ok = False,
                    .message = ex.Message
                })
            End Try
        End Sub

        Private Function ParseFamilySuitability(fd As Dictionary(Of String, Object)) As MultiFamilySuitabilityOptions
            Dim opt As New MultiFamilySuitabilityOptions()
            Dim obj = GetDictValue(fd, "familysuitability")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.CriteriaExcelPath = SafeStr(GetDictValue(d, "criteriaExcelPath")).Trim()
            opt.CriteriaRowCount = ToInt(GetDictValue(d, "criteriaRowCount"), 0)
            opt.CriteriaUniqueCount = ToInt(GetDictValue(d, "criteriaComboCount"), 0)
            opt.CriteriaSheetCount = ToInt(GetDictValue(d, "criteriaSheetCount"), 0)
            opt.MatchReviewText = SafeStr(GetDictValue(d, "matchReviewText")).Trim()
            opt.MismatchReviewText = SafeStr(GetDictValue(d, "mismatchReviewText")).Trim()

            For Each item In EnumeratePayloadItems(GetDictValue(d, "filterRules"))
                Dim ruleDict = ToDict(item)
                opt.FilterRules.Add(New FamilySuitabilityReviewService.FilterRule With {
                    .Target = NormalizeFamilySuitabilityFilterTarget(SafeStr(GetDictValue(ruleDict, "target"))),
                    .Keyword = SafeStr(GetDictValue(ruleDict, "keyword")).Trim(),
                    .ReviewText = SafeStr(GetDictValue(ruleDict, "reviewText")).Trim()
                })
            Next

            Return opt
        End Function

        Private Sub PrepareFamilySuitabilityCriteria(req As MultiRunRequest)
            If req Is Nothing OrElse req.FamilySuitability Is Nothing OrElse Not req.FamilySuitability.Enabled Then Return

            Dim excelPath As String = SafeStr(req.FamilySuitability.CriteriaExcelPath).Trim()
            If String.IsNullOrWhiteSpace(excelPath) Then
                Throw New InvalidOperationException("Family 적합성 기준 엑셀 파일을 선택하세요.")
            End If

            Dim snapshot = LoadFamilySuitabilityCriteria(excelPath)
            req.FamilySuitability.CriteriaExcelPath = snapshot.SourcePath
            req.FamilySuitability.CriteriaRules = snapshot.CriteriaRules
            req.FamilySuitability.CriteriaRowCount = snapshot.RowCount
            req.FamilySuitability.CriteriaUniqueCount = snapshot.UniqueCount
            req.FamilySuitability.CriteriaSheetCount = snapshot.SheetCount
        End Sub

        Private Sub RunFamilySuitabilityMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.FamilySuitability Is Nothing OrElse Not _multiRequest.FamilySuitability.Enabled Then Return

            Dim settings As New FamilySuitabilityReviewService.Settings With {
                .CriteriaRules = If(_multiRequest.FamilySuitability.CriteriaRules, New List(Of FamilySuitabilityReviewService.CriteriaRule)()),
                .MatchReviewText = If(_multiRequest.FamilySuitability.MatchReviewText, String.Empty),
                .MismatchReviewText = If(_multiRequest.FamilySuitability.MismatchReviewText, String.Empty),
                .FilterRules = If(_multiRequest.FamilySuitability.FilterRules, New List(Of FamilySuitabilityReviewService.FilterRule)())
            }

            Dim result = FamilySuitabilityReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "Family 적합성 검토 실행 중", $"{safeName} · {msg}")
                End Sub)

            If _multiFamilySuitabilityRows Is Nothing Then _multiFamilySuitabilityRows = New List(Of FamilySuitabilityReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiFamilySuitabilityRows.AddRange(result.Rows)
            End If

            If _multiFamilySuitabilityFileSummaries Is Nothing Then _multiFamilySuitabilityFileSummaries = New List(Of FamilySuitabilityReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiFamilySuitabilityFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiFamilySuitabilityFileSummaries.Add(New FamilySuitabilityReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .Total = 0,
                    .Issues = 0,
                    .Near = 0,
                    .Reason = "집계 가능한 객체가 없습니다."
                })
            End If

            If _multiFamilySuitabilityWarnings Is Nothing Then _multiFamilySuitabilityWarnings = New List(Of String)()
            If result IsNot Nothing AndAlso result.Warnings IsNot Nothing Then
                For Each warning In result.Warnings
                    If String.IsNullOrWhiteSpace(warning) Then Continue For
                    _multiFamilySuitabilityWarnings.Add($"{safeName}: {warning}")
                Next
            End If
        End Sub

        Private Sub ClearMultiFamilySuitabilityCache()
            _multiFamilySuitabilityRows = Nothing
            _multiFamilySuitabilityFileSummaries = Nothing
            _multiFamilySuitabilityWarnings = Nothing
        End Sub

        Private Function GetMultiFamilySuitabilityRowCount() As Integer
            Return If(_multiFamilySuitabilityRows, New List(Of FamilySuitabilityReviewService.ReviewRow)()).Count
        End Function

        Private Function BuildFamilySuitabilityMultiSummary() As Object
            Dim rows = If(_multiFamilySuitabilityRows, New List(Of FamilySuitabilityReviewService.ReviewRow)())
            Dim summaries = If(_multiFamilySuitabilityFileSummaries, New List(Of FamilySuitabilityReviewService.FileSummary)())
            Dim fileSummaries = BuildFamilySuitabilityFileSummaries()

            Dim totalElements As Integer = 0
            Dim matchCount As Integer = 0
            Dim mismatchCount As Integer = 0
            Dim filterCount As Integer = 0
            For Each row In rows
                If row Is Nothing Then Continue For
                totalElements += row.ElementCount
                Select Case SafeStr(row.ReviewSource).Trim().ToUpperInvariant()
                    Case "MATCH"
                        matchCount += 1
                    Case "FILTER"
                        filterCount += 1
                    Case Else
                        mismatchCount += 1
                End Select
            Next

            Dim criteriaFileName As String = "(미선택)"
            Dim criteriaCount As Integer = 0
            Dim filterRuleCount As Integer = 0
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FamilySuitability IsNot Nothing Then
                criteriaFileName = GetFamilySuitabilityCriteriaFileLabel(_multiRequest.FamilySuitability.CriteriaExcelPath)
                criteriaCount = Math.Max(_multiRequest.FamilySuitability.CriteriaUniqueCount, If(_multiRequest.FamilySuitability.CriteriaRules, New List(Of FamilySuitabilityReviewService.CriteriaRule)()).Count)
                filterRuleCount = CountActiveFamilySuitabilityFilters(_multiRequest.FamilySuitability.FilterRules)
            End If

            Dim warningCount As Integer = If(_multiFamilySuitabilityWarnings, New List(Of String)()).Count
            Return New With {
                .key = "familysuitability",
                .label = "Family 적합성 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"기준 엑셀: {criteriaFileName}",
                    $"기준 조합 수: {criteriaCount}개",
                    $"필터 규칙 수: {filterRuleCount}개",
                    $"사용 객체 수: {totalElements}개",
                    $"결과 행 수: {rows.Count}행",
                    $"기준 일치 / 미일치 / 필터 적용: {matchCount} / {mismatchCount} / {filterCount}",
                    $"경고 수: {warningCount}건"
                },
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildFamilySuitabilityFileSummaries() As List(Of Object)
            Dim summaries = If(_multiFamilySuitabilityFileSummaries, New List(Of FamilySuitabilityReviewService.FileSummary)())
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
                Dim near As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = String.Empty

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.Total
                    issues = summary.Issues
                    near = summary.Near
                    statusText = If(String.IsNullOrWhiteSpace(summary.Status), "success", summary.Status)
                    reason = If(summary.Reason, String.Empty)
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
                    .near = near,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function FindFamilySuitabilityRunItem(fileName As String) As MultiRunItem
            If String.IsNullOrWhiteSpace(fileName) OrElse _multiRunItems Is Nothing Then Return Nothing

            Return _multiRunItems.FirstOrDefault(Function(item) item IsNot Nothing AndAlso
                String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Function ResolveFamilySuitabilityEmptyMessage(fileName As String,
                                                             summary As FamilySuitabilityReviewService.FileSummary) As String
            Dim runItem = FindFamilySuitabilityRunItem(fileName)
            If runItem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(runItem.Reason) Then
                Return runItem.Reason
            End If

            If summary IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(summary.Reason) Then
                Return summary.Reason
            End If

            Return "집계 가능한 객체가 없습니다."
        End Function

        Private Function TrimFamilySuitabilityExportColumns(table As DataTable) As DataTable
            If table Is Nothing Then Return Nothing

            Dim allowed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "Category",
                "Family",
                "Type",
                "No. of Elements",
                "Review",
                "__ReviewEn",
                "__ReviewKo",
                "Status"
            }

            Dim removeNames As New List(Of String)()
            For Each column As DataColumn In table.Columns
                If column Is Nothing Then Continue For
                If allowed.Contains(column.ColumnName) Then Continue For
                removeNames.Add(column.ColumnName)
            Next

            For Each name In removeNames
                table.Columns.Remove(name)
            Next

            Return table
        End Function

        Private Sub ExportFamilySuitability(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiFamilySuitabilityRows, New List(Of FamilySuitabilityReviewService.ReviewRow)())
            Dim summaries = If(_multiFamilySuitabilityFileSummaries, New List(Of FamilySuitabilityReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "Family 적합성 검토 결과가 없습니다."})
                Return
            End If

            Dim orderedNames As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each path In _multiRequest.RvtPaths
                    Dim safeName As String = GetSafeMultiFileName(path)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            For Each summary In summaries
                If summary Is Nothing Then Continue For
                Dim safeName As String = GetSafeMultiFileName(summary.File)
                If String.IsNullOrWhiteSpace(safeName) Then Continue For
                If seen.Add(safeName) Then orderedNames.Add(safeName)
            Next

            If orderedNames.Count = 0 Then
                For Each row In rows
                    If row Is Nothing Then Continue For
                    Dim safeName As String = GetSafeMultiFileName(row.File)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each fileName In orderedNames
                Dim fileRows = rows.Where(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase)).ToList()
                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                Dim emptyMessage As String = "집계 가능한 객체가 없습니다."
                If summary IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(summary.Reason) Then
                    emptyMessage = summary.Reason
                End If
                emptyMessage = ResolveFamilySuitabilityEmptyMessage(fileName, summary)
                Dim table = TrimFamilySuitabilityExportColumns(FamilySuitabilityReviewService.BuildExportTable(fileRows, emptyMessage))
                Dim sheetName = BuildFamilySuitabilitySheetName(fileName)
                sheets.Add(New KeyValuePair(Of String, DataTable)(sheetName, table))
                SetSplitExportIssueCount(fileIssueCounts, sheetName, If(summary Is Nothing, 0, summary.Issues))
            Next

            If sheets.Count = 0 Then
                sheets.Add(New KeyValuePair(Of String, DataTable)("Review", TrimFamilySuitabilityExportColumns(FamilySuitabilityReviewService.BuildExportTable(rows))))
            End If

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "familysuitability", "Family적합성검토", "Review", sheets, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(
                sheets,
                BuildFamilySuitabilityDefaultExcelName(),
                doAutoFit,
                "hub:multi-progress",
                sheetKeyOverride:="familysuitability",
                exportKind:="familysuitability",
                exportLocale:=exportLocale)

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("familysuitability", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Function BuildFamilySuitabilityDefaultExcelName() As String
            Dim baseName As String = String.Empty
            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing AndAlso _multiRequest.RvtPaths.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiRequest.RvtPaths(0)))
            ElseIf _multiFamilySuitabilityFileSummaries IsNot Nothing AndAlso _multiFamilySuitabilityFileSummaries.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiFamilySuitabilityFileSummaries(0).File))
            End If

            If String.IsNullOrWhiteSpace(baseName) Then
                Return $"FamilySuitabilityReview_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Return $"{baseName}_FamilySuitabilityReview.xlsx"
        End Function

        Private Function BuildFamilySuitabilitySheetName(fileName As String) As String
            Dim safeName As String = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(fileName))
            If String.IsNullOrWhiteSpace(safeName) Then safeName = "Review"
            Return safeName
        End Function

        Private Function CountActiveFamilySuitabilityFilters(filters As IEnumerable(Of FamilySuitabilityReviewService.FilterRule)) As Integer
            Return If(filters, Enumerable.Empty(Of FamilySuitabilityReviewService.FilterRule)()) _
                .Where(Function(rule) rule IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(rule.Keyword) AndAlso Not String.IsNullOrWhiteSpace(rule.ReviewText)) _
                .Count()
        End Function

        Private Function GetFamilySuitabilityCriteriaFileLabel(path As String) As String
            Dim safePath As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(safePath) Then Return "(미선택)"
            Try
                Return System.IO.Path.GetFileName(safePath)
            Catch
                Return safePath
            End Try
        End Function

        Private Function NormalizeFamilySuitabilityFilterTarget(value As String) As String
            Dim normalized As String = SafeStr(value).Trim().ToLowerInvariant()
            If normalized = "family" Then Return "family"
            If normalized = "type" Then Return "type"
            Return "familyOrType"
        End Function

        Private Function LoadFamilySuitabilityCriteria(path As String) As FamilySuitabilityCriteriaSnapshot
            Dim safePath As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(safePath) Then
                Throw New InvalidOperationException("Family 적합성 기준 엑셀 파일 경로가 비어 있습니다.")
            End If
            If Not File.Exists(safePath) Then
                Throw New FileNotFoundException("기준 엑셀 파일을 찾을 수 없습니다.", safePath)
            End If

            Dim snapshot As New FamilySuitabilityCriteriaSnapshot With {
                .SourcePath = safePath
            }
            Dim uniqueRules As New Dictionary(Of String, FamilySuitabilityReviewService.CriteriaRule)(StringComparer.OrdinalIgnoreCase)
            Dim formatter As New DataFormatter()

            Using stream As New FileStream(safePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim workbook = WorkbookFactory.Create(stream)
                For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                    Dim sheet = workbook.GetSheetAt(sheetIndex)
                    Dim header = FindFamilySuitabilityHeaderMap(sheet, formatter)
                    If header Is Nothing Then Continue For

                    snapshot.SheetCount += 1
                    For rowIndex As Integer = header.HeaderRowIndex + 1 To sheet.LastRowNum
                        Dim row = sheet.GetRow(rowIndex)
                        If row Is Nothing Then Continue For

                        Dim category As String = GetFamilySuitabilityCellText(row, header.CategoryIndex, formatter)
                        Dim familyName As String = GetFamilySuitabilityCellText(row, header.FamilyIndex, formatter)
                        Dim typeName As String = GetFamilySuitabilityCellText(row, header.TypeIndex, formatter)
                        If String.IsNullOrWhiteSpace(category) AndAlso String.IsNullOrWhiteSpace(familyName) AndAlso String.IsNullOrWhiteSpace(typeName) Then
                            Continue For
                        End If
                        If String.IsNullOrWhiteSpace(category) OrElse String.IsNullOrWhiteSpace(familyName) OrElse String.IsNullOrWhiteSpace(typeName) Then
                            Continue For
                        End If

                        snapshot.RowCount += 1
                        Dim key As String = BuildFamilySuitabilityCriteriaKey(category, familyName, typeName)
                        If String.IsNullOrWhiteSpace(key) Then Continue For
                        If uniqueRules.ContainsKey(key) Then Continue For

                        uniqueRules.Add(key, New FamilySuitabilityReviewService.CriteriaRule With {
                            .Category = category,
                            .Family = familyName,
                            .TypeName = typeName
                        })
                    Next
                Next
            End Using

            If snapshot.SheetCount = 0 Then
                Throw New InvalidOperationException("기준 엑셀에서 Category / Family / Type 헤더를 찾지 못했습니다.")
            End If

            snapshot.CriteriaRules = uniqueRules.Values _
                .OrderBy(Function(item) item.Category, StringComparer.OrdinalIgnoreCase) _
                .ThenBy(Function(item) item.Family, StringComparer.OrdinalIgnoreCase) _
                .ThenBy(Function(item) item.TypeName, StringComparer.OrdinalIgnoreCase) _
                .ToList()
            snapshot.UniqueCount = snapshot.CriteriaRules.Count

            If snapshot.UniqueCount = 0 Then
                Throw New InvalidOperationException("기준 엑셀에 유효한 Category / Family / Type 조합이 없습니다.")
            End If

            Return snapshot
        End Function

        Private Function FindFamilySuitabilityHeaderMap(sheet As ISheet, formatter As DataFormatter) As FamilySuitabilityHeaderMap
            If sheet Is Nothing Then Return Nothing

            For rowIndex As Integer = sheet.FirstRowNum To sheet.LastRowNum
                Dim row = sheet.GetRow(rowIndex)
                If row Is Nothing Then Continue For

                Dim map As New FamilySuitabilityHeaderMap With {
                    .HeaderRowIndex = rowIndex
                }

                Dim lastCell As Integer = CInt(row.LastCellNum) - 1
                For colIndex As Integer = 0 To lastCell
                    Dim headerText As String = NormalizeFamilySuitabilityHeader(GetFamilySuitabilityCellText(row, colIndex, formatter))
                    If headerText = "category" AndAlso map.CategoryIndex < 0 Then
                        map.CategoryIndex = colIndex
                    ElseIf headerText = "family" AndAlso map.FamilyIndex < 0 Then
                        map.FamilyIndex = colIndex
                    ElseIf headerText = "type" AndAlso map.TypeIndex < 0 Then
                        map.TypeIndex = colIndex
                    End If
                Next

                If map.CategoryIndex >= 0 AndAlso map.FamilyIndex >= 0 AndAlso map.TypeIndex >= 0 Then
                    Return map
                End If
            Next

            Return Nothing
        End Function

        Private Function GetFamilySuitabilityCellText(row As IRow, columnIndex As Integer, formatter As DataFormatter) As String
            If row Is Nothing OrElse columnIndex < 0 Then Return String.Empty
            Dim cell = row.GetCell(columnIndex)
            If cell Is Nothing Then Return String.Empty
            Return CleanFamilySuitabilityText(formatter.FormatCellValue(cell))
        End Function

        Private Function NormalizeFamilySuitabilityHeader(value As String) As String
            Return CleanFamilySuitabilityText(value).Replace(" ", String.Empty).ToLowerInvariant()
        End Function

        Private Function BuildFamilySuitabilityCriteriaKey(category As String, familyName As String, typeName As String) As String
            Dim normCategory As String = CleanFamilySuitabilityText(category).ToLowerInvariant()
            Dim normFamily As String = CleanFamilySuitabilityText(familyName).ToLowerInvariant()
            Dim normType As String = CleanFamilySuitabilityText(typeName).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(normCategory) OrElse String.IsNullOrWhiteSpace(normFamily) OrElse String.IsNullOrWhiteSpace(normType) Then
                Return String.Empty
            End If
            Return $"{normCategory}|{normFamily}|{normType}"
        End Function

        Private Function CleanFamilySuitabilityText(value As String) As String
            Dim text As String = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(text) Then Return String.Empty
            text = text.Trim()
            text = text.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ").Replace(ControlChars.Tab, " ")
            Return String.Join(" ", text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))
        End Function

    End Class

End Namespace
