Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Models
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiParameterMissingOptions
            Public Property Enabled As Boolean
            Public Property ParameterNames As List(Of String) = New List(Of String)()
            Public Property TargetFilter As ElementParameterUpdateSettings = New ElementParameterUpdateSettings()
            Public Property ExcludeTargetFilterMatches As Boolean
            Public Property ExceptionRules As List(Of ParameterMissingReviewService.MissingRule) = New List(Of ParameterMissingReviewService.MissingRule)()
        End Class

        Private Shared _multiParameterMissingRows As List(Of ParameterMissingReviewService.ReviewRow)
        Private Shared _multiParameterMissingFileSummaries As List(Of ParameterMissingReviewService.FileSummary)

        Private Function ParseParameterMissing(fd As Dictionary(Of String, Object)) As MultiParameterMissingOptions
            Dim opt As New MultiParameterMissingOptions()
            Dim obj = GetDictValue(fd, "parametermissing")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.ParameterNames = ParseStringList(d, "parameterNames") _
                .Select(Function(x) SafeStr(x).Trim()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()

            Dim targetFilterObj = GetDictValue(d, "targetFilter")
            opt.TargetFilter = ParseDeliveryCleanerElementUpdate(targetFilterObj)
            If opt.TargetFilter Is Nothing Then opt.TargetFilter = New ElementParameterUpdateSettings()
            opt.TargetFilter.Assignments = New List(Of ElementParameterAssignment)()
            Dim targetFilterDict = ToDict(targetFilterObj)
            Dim targetFilterMode = SafeStr(GetDictValue(targetFilterDict, "mode")).Trim().ToLowerInvariant()
            opt.ExcludeTargetFilterMatches = String.Equals(targetFilterMode, "exclude", StringComparison.OrdinalIgnoreCase)

            Dim rules As New List(Of ParameterMissingReviewService.MissingRule)()
            For Each item In EnumeratePayloadItems(GetDictValue(d, "exceptionRules"))
                Dim itemDict = ToDict(item)
                Dim parameterName As String = SafeStr(GetDictValue(itemDict, "parameterName")).Trim()
                Dim combinationMode = ParseDeliveryCleanerCombinationMode(GetDictValue(itemDict, "combinationMode"), ParameterConditionCombination.Or)
                Dim ruleSettings = ParseDeliveryCleanerElementUpdate(New Dictionary(Of String, Object) From {
                    {"enabled", True},
                    {"combinationMode", combinationMode.ToString()},
                    {"conditions", GetDictValue(itemDict, "conditions")}
                })
                If ruleSettings Is Nothing Then ruleSettings = New ElementParameterUpdateSettings()

                rules.Add(New ParameterMissingReviewService.MissingRule With {
                    .Enabled = SafeBoolObj(GetDictValue(itemDict, "enabled"), True),
                    .ParameterName = parameterName,
                    .CombinationMode = combinationMode,
                    .Conditions = If(ruleSettings.Conditions, New List(Of ElementParameterCondition)())
                })
            Next

            opt.ExceptionRules = rules
            Return opt
        End Function

        Private Sub RunParameterMissingMultiForDocument(doc As Document, safeName As String, basePct As Double)
            If _multiRequest Is Nothing OrElse _multiRequest.ParameterMissing Is Nothing OrElse Not _multiRequest.ParameterMissing.Enabled Then Return

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

            Dim settings As New ParameterMissingReviewService.Settings With {
                .ParameterNames = If(_multiRequest.ParameterMissing.ParameterNames, New List(Of String)()),
                .TargetFilter = If(_multiRequest.ParameterMissing.TargetFilter, New ElementParameterUpdateSettings()),
                .ExcludeTargetFilterMatches = _multiRequest.ParameterMissing.ExcludeTargetFilterMatches,
                .ExceptionRules = If(_multiRequest.ParameterMissing.ExceptionRules, New List(Of ParameterMissingReviewService.MissingRule)()),
                .ExtraParameterNames = commonExtraParamNames
            }
            settings.TargetFilter.Assignments = New List(Of ElementParameterAssignment)()
            ApplyParameterMissingScopeSettings(settings, hasAllowedElementScope, allowedElementIds, commonTargetFilter, commonExcludeTargetFilter)

            Dim result = ParameterMissingReviewService.RunOnDocument(
                doc,
                safeName,
                settings,
                Sub(pct, msg)
                    Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                    ReportMultiProgress(overallPct, "파라미터 누락 검토 실행 중", $"{safeName} - {msg}")
                End Sub)

            If _multiParameterMissingRows Is Nothing Then _multiParameterMissingRows = New List(Of ParameterMissingReviewService.ReviewRow)()
            If result IsNot Nothing AndAlso result.Rows IsNot Nothing Then
                _multiParameterMissingRows.AddRange(result.Rows)
            End If

            If _multiParameterMissingFileSummaries Is Nothing Then _multiParameterMissingFileSummaries = New List(Of ParameterMissingReviewService.FileSummary)()
            If result IsNot Nothing AndAlso result.FileSummaries IsNot Nothing Then
                _multiParameterMissingFileSummaries.AddRange(result.FileSummaries)
            Else
                _multiParameterMissingFileSummaries.Add(New ParameterMissingReviewService.FileSummary With {
                    .File = safeName,
                    .Status = "success",
                    .TargetElementCount = 0,
                    .ParameterCount = 0,
                    .TotalReviewed = 0,
                    .ErrorCount = 0,
                    .IgnoredCount = 0,
                    .OkCount = 0,
                    .TargetConditionCount = 0,
                    .ExceptionRuleCount = 0,
                    .Reason = "검토 결과가 없습니다."
                })
            End If
        End Sub

        Private Shared Sub ApplyParameterMissingScopeSettings(settings As ParameterMissingReviewService.Settings,
                                                              hasAllowedElementScope As Boolean,
                                                              allowedElementIds As List(Of Integer),
                                                              commonTargetFilter As String,
                                                              commonExcludeTargetFilter As String)
            If settings Is Nothing Then Return

            TrySetParameterMissingSettingsProperty(settings, "HasAllowedElementScope", hasAllowedElementScope)
            TrySetParameterMissingSettingsProperty(settings, "AllowedElementIds", If(allowedElementIds, New List(Of Integer)()))
            TrySetParameterMissingSettingsProperty(settings, "CommonTargetFilterText", SafeStr(commonTargetFilter))
            TrySetParameterMissingSettingsProperty(settings, "CommonExcludeTargetFilterText", SafeStr(commonExcludeTargetFilter))
        End Sub

        Private Shared Sub TrySetParameterMissingSettingsProperty(settings As ParameterMissingReviewService.Settings,
                                                                  propertyName As String,
                                                                  value As Object)
            If settings Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return

            Dim prop = settings.GetType().GetProperty(propertyName, Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public)
            If prop Is Nothing OrElse Not prop.CanWrite Then Return

            Try
                prop.SetValue(settings, value, Nothing)
            Catch
            End Try
        End Sub

        Private Sub ClearMultiParameterMissingCache()
            _multiParameterMissingRows = Nothing
            _multiParameterMissingFileSummaries = Nothing
        End Sub

        Private Function BuildParameterMissingMultiSummary() As Object
            Dim summaries = If(_multiParameterMissingFileSummaries, New List(Of ParameterMissingReviewService.FileSummary)())
            Dim totalTargets As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TargetElementCount))
            Dim totalReviewed As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.TotalReviewed))
            Dim totalErrors As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.ErrorCount))
            Dim totalIgnored As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.IgnoredCount))
            Dim totalOk As Integer = summaries.Sum(Function(item) If(item Is Nothing, 0, item.OkCount))

            Dim parameterNames As List(Of String) = New List(Of String)()
            Dim exceptionRuleCount As Integer = 0
            Dim commonFilterLabel As String = "없음"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.ParameterMissing IsNot Nothing Then
                parameterNames = If(_multiRequest.ParameterMissing.ParameterNames, New List(Of String)())
                exceptionRuleCount = CountConfiguredMissingRules(_multiRequest.ParameterMissing.ExceptionRules)
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(_multiRequest.Common.TargetFilter) OrElse Not String.IsNullOrWhiteSpace(_multiRequest.Common.ExcludeTargetFilter) Then
                    commonFilterLabel = "적용"
                End If
            End If

            Return New With {
                .key = "parametermissing",
                .label = "파라미터 누락 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"검토 파라미터 수: {parameterNames.Count}개",
                    $"검토 대상 객체 수: {totalTargets}개",
                    $"검토 셀 수: {totalReviewed}개",
                    $"공통 검토대상 필터: {commonFilterLabel}",
                    $"누락 예외 규칙 수: {exceptionRuleCount}개",
                    $"누락 오류 수: {totalErrors}건",
                    $"예외 제외 수: {totalIgnored}건",
                    $"정상 값 수: {totalOk}개",
                    $"엑셀 결과 행 수: {If(_multiParameterMissingRows, New List(Of ParameterMissingReviewService.ReviewRow)()).Count}행"
                },
                .fileSummaries = BuildParameterMissingFileSummaries()
            }
        End Function

        Private Function BuildParameterMissingFileSummaries() As List(Of Object)
            Dim summaries = If(_multiParameterMissingFileSummaries, New List(Of ParameterMissingReviewService.FileSummary)())
            Dim orderedNames = BuildOrderedMultiFileNames(summaries.Select(Function(item) If(item Is Nothing, "", item.File)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issues As Integer = 0
                Dim nearCount As Integer = 0
                Dim statusText As String = "pending"
                Dim reason As String = ""

                Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                If summary IsNot Nothing Then
                    total = summary.TotalReviewed
                    issues = summary.ErrorCount
                    nearCount = summary.IgnoredCount
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
                    .near = nearCount,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Sub ExportParameterMissing(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiParameterMissingRows, New List(Of ParameterMissingReviewService.ReviewRow)())
            Dim summaries = If(_multiParameterMissingFileSummaries, New List(Of ParameterMissingReviewService.FileSummary)())
            If rows.Count = 0 AndAlso summaries.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "파라미터 누락 검토 결과가 없습니다."})
                Return
            End If

            Dim orderedNames = BuildOrderedMultiFileNames(
                rows.Select(Function(row) If(row Is Nothing, "", row.File)),
                summaries.Select(Function(summary) If(summary Is Nothing, "", summary.File)))

            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim defaultFileName As String
            If orderedNames.Count >= 2 OrElse requestedCount >= 2 Then
                defaultFileName = $"ParameterMissing_Selected {Math.Max(orderedNames.Count, requestedCount)} Files.xlsx"
            Else
                defaultFileName = $"ParameterMissing_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Dim saved As String = ""
            If Not String.IsNullOrWhiteSpace(outputFolder) OrElse orderedNames.Count >= 2 Then
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim perFileRows = rows _
                        .Where(Function(row) row IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(row.File), fileName, StringComparison.OrdinalIgnoreCase)) _
                        .ToList()
                    If perFileRows.Count = 0 Then Continue For

                    Dim baseName As String = fileName
                    Try
                        baseName = IO.Path.GetFileNameWithoutExtension(fileName)
                    Catch
                    End Try
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = fileName

                    Dim table = ParameterMissingReviewService.BuildExportTable(perFileRows)
                    ExcelCore.EnsureNoDataRow(table, "검토 결과가 없습니다.")
                    sheetList.Add(New KeyValuePair(Of String, DataTable)(baseName, table))

                    Dim summary = summaries.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase))
                    SetSplitExportIssueCount(fileIssueCounts, baseName, If(summary Is Nothing, 0, summary.ErrorCount))
                Next

                If sheetList.Count = 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "파라미터 누락 검토 결과가 없습니다."})
                    Return
                End If

                If Not String.IsNullOrWhiteSpace(outputFolder) Then
                    Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "parametermissing", "ParameterMissingReview", "파라미터 누락 검토", sheetList, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                    If savedCount <= 0 Then
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 내보내기가 취소되었습니다."})
                    Else
                        SendSplitExportCompleted(outputFolder, savedCount)
                    End If
                    Return
                End If

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="parametermissing", exportKind:="parametermissing", exportLocale:=exportLocale)
            Else
                Dim table = ParameterMissingReviewService.BuildExportTable(rows)
                ExcelCore.EnsureNoDataRow(table, "검토 결과가 없습니다.")
                saved = ExcelCore.PickAndSaveXlsx("파라미터 누락 검토", table, defaultFileName, doAutoFit, "hub:multi-progress", "parametermissing", exportLocale:=exportLocale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 내보내기가 취소되었습니다."})
            Else
                TryApplyExportStyles("parametermissing", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Shared Function CountConfiguredConditions(settings As ElementParameterUpdateSettings) As Integer
            If settings Is Nothing OrElse settings.Conditions Is Nothing Then Return 0
            Return settings.Conditions _
                .Where(Function(condition) condition IsNot Nothing AndAlso condition.IsConfigured() AndAlso (condition.Enabled OrElse Not String.IsNullOrWhiteSpace(condition.ParameterName))) _
                .Count()
        End Function

        Private Shared Function CountConfiguredMissingRules(rules As IEnumerable(Of ParameterMissingReviewService.MissingRule)) As Integer
            Return (If(rules, Enumerable.Empty(Of ParameterMissingReviewService.MissingRule)())) _
                .Where(Function(rule) rule IsNot Nothing AndAlso rule.HasConfiguredConditions()) _
                .Count()
        End Function

    End Class

End Namespace
