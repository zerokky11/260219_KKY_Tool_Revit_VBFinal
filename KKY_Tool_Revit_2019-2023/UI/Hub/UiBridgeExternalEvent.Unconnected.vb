Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiUnconnectedOptions
            Public Property Enabled As Boolean
            Public Property IncludeCenterAxisCheck As Boolean
            Public Property CenterAxisTol As Double = 0.5R
            Public Property CenterAxisUnit As String = "mm"
        End Class

        Private Shared _multiUnconnectedRows As List(Of UnconnectedConnectorReviewService.ReviewRow)
        Private Shared _multiUnconnectedFileSummaries As List(Of UnconnectedConnectorReviewService.FileSummary)

        Private Function ParseUnconnected(fd As Dictionary(Of String, Object)) As MultiUnconnectedOptions
            Dim opt As New MultiUnconnectedOptions()
            Dim obj = GetDictValue(fd, "unconnected")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.IncludeCenterAxisCheck = ToBool(GetDictValue(d, "includeCenterAxisCheck"))
            opt.CenterAxisTol = ToDouble(GetDictValue(d, "centerAxisTol"), 0.5R)
            If opt.CenterAxisTol <= 0 Then opt.CenterAxisTol = 0.5R
            opt.CenterAxisUnit = NormalizeTapAlignUnit(SafeStr(GetDictValue(d, "centerAxisUnit")))
            If String.IsNullOrWhiteSpace(opt.CenterAxisUnit) Then opt.CenterAxisUnit = "mm"
            Return opt
        End Function

        Private Sub RunUnconnectedMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.Unconnected Is Nothing OrElse Not _multiRequest.Unconnected.Enabled Then Return

            Dim commonTargetFilter As String = String.Empty
            Dim commonExcludeTargetFilter As String = String.Empty
            Dim allowedElementIds As New List(Of Integer)()
            Dim hasAllowedElementScope As Boolean = False
            If _multiRequest.Common IsNot Nothing Then
                commonTargetFilter = SafeStr(_multiRequest.Common.TargetFilter)
                commonExcludeTargetFilter = SafeStr(_multiRequest.Common.ExcludeTargetFilter)
                hasAllowedElementScope = TryBuildCommonScopeIds(doc, commonTargetFilter, commonExcludeTargetFilter, allowedElementIds)
            End If

            Dim settings As New UnconnectedConnectorReviewService.Settings With {
                .HasAllowedElementScope = hasAllowedElementScope,
                .AllowedElementIds = If(allowedElementIds, New List(Of Integer)()),
                .CommonTargetFilterText = commonTargetFilter,
                .CommonExcludeTargetFilterText = commonExcludeTargetFilter
            }

            Dim result = UnconnectedConnectorReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "미연결 검토 실행 중", $"{safeName} · {msg}")
                End Sub)

            Dim resultRows As IEnumerable(Of UnconnectedConnectorReviewService.ReviewRow) = Nothing
            If result IsNot Nothing Then resultRows = result.Rows
            Dim fullyUnconnectedElementIds As HashSet(Of String) = BuildFullyUnconnectedElementIdSet(resultRows)

            Dim centerAxisTargetCount As Integer = 0
            Dim centerAxisRows = BuildUnconnectedCenterAxisRows(doc,
                                                                safeName,
                                                                commonTargetFilter,
                                                                commonExcludeTargetFilter,
                                                                basePct,
                                                                fullyUnconnectedElementIds,
                                                                centerAxisTargetCount)

            Dim documentRows As New List(Of UnconnectedConnectorReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                documentRows.AddRange(result.Rows)
            End If
            If centerAxisRows IsNot Nothing AndAlso centerAxisRows.Count > 0 Then
                documentRows.AddRange(centerAxisRows)
            End If
            If documentRows.Count > 0 Then
                UnconnectedConnectorReviewService.ApplyGroupItemTexts(documentRows)
            End If

            If _multiUnconnectedRows Is Nothing Then _multiUnconnectedRows = New List(Of UnconnectedConnectorReviewService.ReviewRow)()
            _multiUnconnectedRows.AddRange(documentRows)

            If _multiUnconnectedFileSummaries Is Nothing Then _multiUnconnectedFileSummaries = New List(Of UnconnectedConnectorReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                For Each summary In result.FileSummaries
                    ApplyCenterAxisSummary(summary, centerAxisTargetCount, If(centerAxisRows, New List(Of UnconnectedConnectorReviewService.ReviewRow)()).Count)
                Next
                _multiUnconnectedFileSummaries.AddRange(result.FileSummaries)
            Else
                Dim summary As New UnconnectedConnectorReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .TargetElementCount = 0,
                    .ConnectorCount = 0,
                    .ErrorCount = 0,
                    .FullErrorCount = 0,
                    .PartialErrorCount = 0,
                    .OkCount = 0,
                    .Reason = "검토 결과가 없습니다."
                }
                ApplyCenterAxisSummary(summary, centerAxisTargetCount, If(centerAxisRows, New List(Of UnconnectedConnectorReviewService.ReviewRow)()).Count)
                _multiUnconnectedFileSummaries.Add(summary)
            End If
        End Sub

        Private Function BuildUnconnectedCenterAxisRows(doc As Document,
                                                        safeName As String,
                                                        commonTargetFilter As String,
                                                        commonExcludeTargetFilter As String,
                                                        basePct As Double,
                                                        fullyUnconnectedElementIds As ISet(Of String),
                                                        ByRef targetCount As Integer) As List(Of UnconnectedConnectorReviewService.ReviewRow)
            targetCount = 0
            Dim rows As New List(Of UnconnectedConnectorReviewService.ReviewRow)()
            If _multiRequest Is Nothing OrElse _multiRequest.Unconnected Is Nothing OrElse Not _multiRequest.Unconnected.IncludeCenterAxisCheck Then Return rows
            If doc Is Nothing Then Return rows

            Dim tol As Double = If(_multiRequest.Unconnected.CenterAxisTol > 0, _multiRequest.Unconnected.CenterAxisTol, 0.5R)
            Dim unit As String = NormalizeTapAlignUnit(_multiRequest.Unconnected.CenterAxisUnit)
            If String.IsNullOrWhiteSpace(unit) Then unit = "mm"
            Dim domain As String = "all"
            Dim combinedTargetFilter = commonTargetFilter
            targetCount = TapAlignmentReviewService.CountTargetsOnDocument(doc, domain, combinedTargetFilter, commonExcludeTargetFilter)

            Dim tapRows = TapAlignmentReviewService.RunOnDocument(doc,
                                                                  tol,
                                                                  unit,
                                                                  domain,
                                                                  New List(Of String)(),
                                                                  combinedTargetFilter,
                                                                  commonExcludeTargetFilter,
                                                                  Sub(pct, msg)
                                                                      Dim fraction As Double = Math.Max(0.0R, Math.Min(CDbl(pct), 1.0R))
                                                                      Dim overallPct = ((basePct + fraction / Math.Max(_multiTotal, 1)) * 100.0R)
                                                                      ReportMultiProgress(overallPct, "중심축 연결 검토 실행 중", $"{safeName} · {msg}")
                                                                  End Sub)

            If tapRows Is Nothing Then Return rows

            Dim skippedFullyUnconnected As Integer = 0
            For Each tapRow In tapRows
                If tapRow Is Nothing OrElse Not IsTapAlignIssueRow(tapRow) Then Continue For

                Dim elementId As String = NormalizeElementIdText(ReadField(tapRow, "ElementId"))
                If IsFullyUnconnectedElementId(elementId, fullyUnconnectedElementIds) Then
                    skippedFullyUnconnected += 1
                    Continue For
                End If

                Dim categoryName As String = ReadField(tapRow, "Category")
                Dim typeName As String = ReadField(tapRow, "Type")
                Dim familyName As String = ReadField(tapRow, "Family")

                rows.Add(New UnconnectedConnectorReviewService.ReviewRow With {
                    .File = safeName,
                    .ItemBase = ResolveUnconnectedCenterAxisItemBase(tapRow),
                    .Id = elementId,
                    .Name = typeName,
                    .Result = "오류",
                    .Content = BuildUnconnectedCenterAxisContent(tapRow, unit),
                    .Etc = String.Empty,
                    .Category = categoryName,
                    .Family = familyName,
                    .IssueKind = "centeraxis",
                    .ConnectorCount = 0,
                    .UnconnectedCount = 0
                })
            Next

            If skippedFullyUnconnected > 0 Then
                targetCount = Math.Max(0, targetCount - skippedFullyUnconnected)
            End If

            Return rows
        End Function

        Private Shared Function BuildFullyUnconnectedElementIdSet(rows As IEnumerable(Of UnconnectedConnectorReviewService.ReviewRow)) As HashSet(Of String)
            Dim ids As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If rows Is Nothing Then Return ids

            For Each row In rows
                If row Is Nothing Then Continue For
                If row.ConnectorCount <= 0 Then Continue For
                If row.UnconnectedCount < row.ConnectorCount Then Continue For

                Dim id As String = NormalizeElementIdText(row.Id)
                If Not String.IsNullOrWhiteSpace(id) Then ids.Add(id)
            Next

            Return ids
        End Function

        Private Shared Function IsFullyUnconnectedElementId(elementId As String, fullyUnconnectedElementIds As ISet(Of String)) As Boolean
            If fullyUnconnectedElementIds Is Nothing OrElse fullyUnconnectedElementIds.Count = 0 Then Return False

            Dim id As String = NormalizeElementIdText(elementId)
            If String.IsNullOrWhiteSpace(id) Then Return False
            Return fullyUnconnectedElementIds.Contains(id)
        End Function

        Private Shared Function NormalizeElementIdText(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty
            Return value.Trim().TrimEnd(","c)
        End Function

        Private Sub ApplyCenterAxisSummary(summary As UnconnectedConnectorReviewService.FileSummary,
                                           targetCount As Integer,
                                           errorCount As Integer)
            If summary Is Nothing Then Return

            Dim enabled As Boolean = _multiRequest IsNot Nothing AndAlso
                                     _multiRequest.Unconnected IsNot Nothing AndAlso
                                     _multiRequest.Unconnected.IncludeCenterAxisCheck
            summary.CenterAxisEnabled = enabled
            If Not enabled Then Return

            summary.CenterAxisTargetCount = Math.Max(0, targetCount)
            summary.CenterAxisErrorCount = Math.Max(0, errorCount)

            Dim axisReason As String
            If summary.CenterAxisTargetCount <= 0 Then
                axisReason = "중심축 검토 대상 없음"
            ElseIf summary.CenterAxisErrorCount <= 0 Then
                axisReason = "중심축 오류 없음"
            Else
                axisReason = $"중심축 오류 {summary.CenterAxisErrorCount}건"
            End If

            If String.IsNullOrWhiteSpace(summary.Reason) Then
                summary.Reason = axisReason
            ElseIf summary.Reason.IndexOf(axisReason, StringComparison.OrdinalIgnoreCase) < 0 Then
                summary.Reason = summary.Reason & " / " & axisReason
            End If
        End Sub

        Private Shared Function BuildUnconnectedCenterAxisContent(row As Dictionary(Of String, Object), unit As String) As String
            Dim typeLabel As String = ResolveUnconnectedCenterAxisTypeLabel(row)
            Dim typeName As String = ReadField(row, "HostType")
            If String.IsNullOrWhiteSpace(typeName) Then typeName = ReadField(row, "Type")
            If String.IsNullOrWhiteSpace(typeName) Then typeName = "알 수 없음"

            Dim status As String = ReadField(row, "Status").Trim().ToUpperInvariant()
            If String.Equals(status, "NO_HOST", StringComparison.OrdinalIgnoreCase) Then
                Return $"[{typeLabel}]: [{typeName}] 연결 중심축을 확인할 수 없습니다."
            End If

            Dim distance As String = ReadField(row, "DistanceFromCenter").Trim()
            Dim angle As String = ReadField(row, "ModeledAngle").Trim()
            Dim normalizedUnit As String = NormalizeTapAlignUnit(unit)
            Dim distanceText As String = If(String.IsNullOrWhiteSpace(distance), "-", distance & normalizedUnit)
            Dim angleText As String = If(String.IsNullOrWhiteSpace(angle), "-", angle)

            Return $"[{typeLabel}]: [{typeName}] 연결이 중심축에서 벗어 났습니다. ( 커넥터 이격거리:{distanceText}), 각도:{angleText}"
        End Function

        Private Shared Function ResolveUnconnectedCenterAxisItemBase(row As Dictionary(Of String, Object)) As String
            Return If(IsUnconnectedCenterAxisDuct(row), "덕트 중심축 연결", "파이프 중심축 연결")
        End Function

        Private Shared Function ResolveUnconnectedCenterAxisTypeLabel(row As Dictionary(Of String, Object)) As String
            Return If(IsUnconnectedCenterAxisDuct(row), "Duct Type", "Pipe Type")
        End Function

        Private Shared Function IsUnconnectedCenterAxisDuct(row As Dictionary(Of String, Object)) As Boolean
            Dim domain As String = ReadField(row, "Domain")
            Dim hostCategory As String = ReadField(row, "HostCategory")
            Dim category As String = ReadField(row, "Category")
            Return ContainsDuctText(domain) OrElse ContainsDuctText(hostCategory) OrElse ContainsDuctText(category)
        End Function

        Private Shared Function ContainsDuctText(value As String) As Boolean
            Dim text As String = If(value, String.Empty)
            Return text.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   text.IndexOf("덕트", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Sub ClearMultiUnconnectedCache()
            _multiUnconnectedRows = Nothing
            _multiUnconnectedFileSummaries = Nothing
        End Sub

        Private Function GetMultiUnconnectedRowCount() As Integer
            Return If(_multiUnconnectedRows, New List(Of UnconnectedConnectorReviewService.ReviewRow)()).Count
        End Function

        Private Function BuildUnconnectedMultiSummary() As Object
            Dim summaries = If(_multiUnconnectedFileSummaries, New List(Of UnconnectedConnectorReviewService.FileSummary)())
            Dim totalTargets As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TargetElementCount))
            Dim totalConnectors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ConnectorCount))
            Dim totalErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ErrorCount))
            Dim totalFullErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.FullErrorCount))
            Dim totalPartialErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.PartialErrorCount))
            Dim totalOk As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.OkCount))
            Dim centerAxisEnabled As Boolean = summaries.Any(Function(item) item IsNot Nothing AndAlso item.CenterAxisEnabled)
            Dim totalCenterAxisTargets As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.CenterAxisTargetCount))
            Dim totalCenterAxisErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.CenterAxisErrorCount))
            Dim centerAxisTolText As String = "0.5"
            Dim centerAxisUnitText As String = "mm"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Unconnected IsNot Nothing Then
                centerAxisTolText = If(_multiRequest.Unconnected.CenterAxisTol > 0,
                                       _multiRequest.Unconnected.CenterAxisTol.ToString(Globalization.CultureInfo.InvariantCulture),
                                       "0.5")
                centerAxisUnitText = NormalizeTapAlignUnit(_multiRequest.Unconnected.CenterAxisUnit)
                If String.IsNullOrWhiteSpace(centerAxisUnitText) Then centerAxisUnitText = "mm"
            End If

            Dim commonFilterLabel As String = "없음"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(_multiRequest.Common.TargetFilter) OrElse Not String.IsNullOrWhiteSpace(_multiRequest.Common.ExcludeTargetFilter) Then
                    commonFilterLabel = "적용"
                End If
            End If

            Return New With {
                .key = "unconnected",
                .label = "미연결 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"공통 검토대상 필터: {commonFilterLabel}",
                    $"커넥터 보유 객체 수: {totalTargets}개",
                    $"검토 커넥터 수: {totalConnectors}개",
                    $"미연결 오류 객체 수: {totalErrors}개",
                    $"전체 미연결: {totalFullErrors}개",
                    $"일부 미연결: {totalPartialErrors}개",
                    $"중심축 연결 검토: {If(centerAxisEnabled, "사용", "미사용")}",
                    $"중심축 허용 기준: {centerAxisTolText} {centerAxisUnitText}",
                    $"중심축 검토 대상: {totalCenterAxisTargets}개",
                    $"중심축 오류 수: {totalCenterAxisErrors}개",
                    $"정상 객체 수: {totalOk}개",
                    $"엑셀 결과 행 수: {GetMultiUnconnectedRowCount()}행"
                },
                .fileSummaries = BuildUnconnectedFileSummaries()
            }
        End Function

        Private Function BuildUnconnectedFileSummaries() As List(Of Object)
            Dim summaries = If(_multiUnconnectedFileSummaries, New List(Of UnconnectedConnectorReviewService.FileSummary)())
            Dim orderedNames = BuildOrderedMultiFileNames(summaries.Select(Function(item) If(item Is Nothing, "", item.File)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issues As Integer = 0
                Dim near As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = ""

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.ConnectorCount
                    issues = summary.ErrorCount + summary.CenterAxisErrorCount
                    near = summary.PartialErrorCount
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

                If summary IsNot Nothing AndAlso summary.CenterAxisEnabled Then
                    Dim splitReason = $"미연결 오류 {summary.ErrorCount}건 / 중심축 오류 {summary.CenterAxisErrorCount}건"
                    If String.IsNullOrWhiteSpace(reason) Then
                        reason = splitReason
                    ElseIf reason.IndexOf(splitReason, StringComparison.OrdinalIgnoreCase) < 0 Then
                        reason = reason & " / " & splitReason
                    End If
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = total,
                    .issues = issues,
                    .near = near,
                    .unconnectedIssues = If(summary Is Nothing, 0, summary.ErrorCount),
                    .centerAxisIssues = If(summary Is Nothing, 0, summary.CenterAxisErrorCount),
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Sub ExportUnconnected(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiUnconnectedRows, New List(Of UnconnectedConnectorReviewService.ReviewRow)())
            Dim summaries = If(_multiUnconnectedFileSummaries, New List(Of UnconnectedConnectorReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "미연결 검토 결과가 없습니다."})
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
                Dim table = UnconnectedConnectorReviewService.BuildExportTable(fileRows)
                ExcelCore.EnsureNoDataRow(table, UnconnectedConnectorReviewService.BuildEmptyExportMessage(summary))

                Dim sheetName As String = BuildUnconnectedSheetName(fileName)
                sheets.Add(New KeyValuePair(Of String, DataTable)(sheetName, table))

                SetSplitExportIssueCount(fileIssueCounts, sheetName, UnconnectedConnectorReviewService.CountIssueRows(fileRows))
            Next

            If sheets.Count = 0 Then
                Dim table = UnconnectedConnectorReviewService.BuildExportTable(rows)
                Dim summary = summaries.FirstOrDefault()
                ExcelCore.EnsureNoDataRow(table, UnconnectedConnectorReviewService.BuildEmptyExportMessage(summary))
                sheets.Add(New KeyValuePair(Of String, DataTable)("Review", table))
            End If

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "unconnected", "UnconnectedConnectorReview", "Review", sheets, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(
                sheets,
                BuildUnconnectedDefaultExcelName(),
                doAutoFit,
                "hub:multi-progress",
                sheetKeyOverride:="unconnected",
                exportKind:="unconnected",
                exportLocale:=exportLocale)

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("unconnected", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Function BuildUnconnectedDefaultExcelName() As String
            Dim baseName As String = String.Empty
            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing AndAlso _multiRequest.RvtPaths.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiRequest.RvtPaths(0)))
            ElseIf _multiUnconnectedFileSummaries IsNot Nothing AndAlso _multiUnconnectedFileSummaries.Count = 1 Then
                baseName = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(_multiUnconnectedFileSummaries(0).File))
            End If

            If String.IsNullOrWhiteSpace(baseName) Then
                Return $"UnconnectedConnectorReview_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Return $"{baseName}_UnconnectedConnectorReview.xlsx"
        End Function

        Private Function BuildUnconnectedSheetName(fileName As String) As String
            Dim safeName As String = Path.GetFileNameWithoutExtension(GetSafeMultiFileName(fileName))
            If String.IsNullOrWhiteSpace(safeName) Then safeName = "Review"
            Return safeName
        End Function

    End Class

End Namespace
