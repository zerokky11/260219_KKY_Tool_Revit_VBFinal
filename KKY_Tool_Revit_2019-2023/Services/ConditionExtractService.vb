Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Models
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application

Namespace Services

    Public NotInheritable Class ConditionExtractService

        Private Sub New()
        End Sub

        Public Class Settings
            Public Property UseActiveDocument As Boolean
            Public Property RvtPaths As List(Of String) = New List(Of String)()
            Public Property OutputFolder As String = String.Empty
            Public Property CloseAllWorksetsOnOpen As Boolean = True
            Public Property ElementParameterUpdate As ElementParameterUpdateSettings
            Public Property ExtractParameterNames As List(Of String) = New List(Of String)()
            Public Property IncludeCoordinates As Boolean
            Public Property IncludeLinearMetrics As Boolean
            Public Property LengthUnit As String = "mm"
            Public Property AreaUnit As String = "mm2"
            Public Property VolumeUnit As String = "mm3"
        End Class

        Public Class RunSummary
            Public Property SuccessCount As Integer
            Public Property FailCount As Integer
            Public Property NoDataCount As Integer
            Public Property TotalCandidateCount As Integer
            Public Property TotalFilteredCount As Integer
            Public Property TotalMatchedCount As Integer
            Public Property TotalExtractedElementCount As Integer
        End Class

        Public Class FileRunResult
            Public Property FilePath As String = String.Empty
            Public Property FileName As String = String.Empty
            Public Property Status As String = String.Empty
            Public Property CandidateCount As Integer
            Public Property FilteredCount As Integer
            Public Property MatchedCount As Integer
            Public Property ExtractedElementCount As Integer
            Public Property WasCentralFile As Boolean
            Public Property UsedLocalFile As Boolean
            Public Property Message As String = String.Empty
        End Class

        Public Class DetailRow
            Public Property FilePath As String = String.Empty
            Public Property FileName As String = String.Empty
            Public Property ElementId As Integer
            Public Property Category As String = String.Empty
            Public Property FamilyName As String = String.Empty
            Public Property TypeName As String = String.Empty
            Public Property Note As String = String.Empty
            Public Property XValue As Double?
            Public Property YValue As Double?
            Public Property ZValue As Double?
            Public Property DirectionX As Double?
            Public Property DirectionY As Double?
            Public Property DirectionZ As Double?
            Public Property LengthValue As Double?
            Public Property ParameterValues As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property ParameterTypeTokens As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        End Class

        Public Class LogEntry
            Public Property Level As String = String.Empty
            Public Property FilePath As String = String.Empty
            Public Property Message As String = String.Empty
        End Class

        Public Class RunResult
            Public Property Ok As Boolean
            Public Property Message As String = String.Empty
            Public Property OutputFolder As String = String.Empty
            Public Property ResultWorkbookPath As String = String.Empty
            Public Property LogTextPath As String = String.Empty
            Public Property Summary As RunSummary = New RunSummary()
            Public Property Files As List(Of FileRunResult) = New List(Of FileRunResult)()
            Public Property Details As List(Of DetailRow) = New List(Of DetailRow)()
            Public Property Logs As List(Of LogEntry) = New List(Of LogEntry)()
        End Class

        Private Class DocumentProcessResult
            Public Property CandidateCount As Integer
            Public Property FilteredCount As Integer
            Public Property MatchedCount As Integer
            Public Property ExtractedElementCount As Integer
            Public Property Message As String = String.Empty
        End Class

        Public Shared Function Run(uiapp As UIApplication,
                                   settings As Settings,
                                   progress As IProgress(Of Object)) As RunResult
            Dim result As New RunResult()
            If uiapp Is Nothing Then
                result.Message = "Revit UIApplication을 찾을 수 없습니다."
                Return result
            End If

            Dim validationMessage = ValidateSettings(uiapp, settings)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                result.Message = validationMessage
                Return result
            End If

            result.OutputFolder = ResolveOutputFolder(uiapp, settings)
            Directory.CreateDirectory(result.OutputFolder)

            Dim parameterNames = GetExtractParameterNames(settings)
            Dim app As RevitApp = uiapp.Application
            Dim totalTargets As Integer = If(settings.UseActiveDocument, 1, settings.RvtPaths.Count)
            ReportProgress(progress, 0, totalTargets, "조건별 객체 속성 추출 준비 중...", result.OutputFolder)

            If settings.UseActiveDocument Then
                Dim activeDoc = uiapp.ActiveUIDocument?.Document
                Dim fileResult = ProcessActiveDocument(activeDoc, settings, parameterNames, result.Logs, result.Details, progress)
                result.Files.Add(fileResult)
            Else
                Dim distinctPaths = settings.RvtPaths _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()

                For i As Integer = 0 To distinctPaths.Count - 1
                    Dim pathText = distinctPaths(i)
                    ReportProgress(progress, i, totalTargets, "파일 준비 중...", pathText)
                    Dim fileResult = ProcessBatchFile(app, pathText, settings, parameterNames, result.Logs, result.Details, progress, i + 1, totalTargets)
                    result.Files.Add(fileResult)
                Next
            End If

            result.Summary = BuildSummary(result.Files)
            result.Ok = result.Files.Any(Function(x) String.Equals(x.Status, "Success", StringComparison.OrdinalIgnoreCase) OrElse
                                                     String.Equals(x.Status, "NoData", StringComparison.OrdinalIgnoreCase))
            If result.Files.Count = 0 Then
                result.Message = "처리할 대상이 없습니다."
            ElseIf result.Summary.FailCount > 0 AndAlso result.Summary.SuccessCount = 0 AndAlso result.Summary.NoDataCount = 0 Then
                result.Message = "모든 대상 처리가 실패했습니다."
            Else
                result.Message = "조건별 객체 속성 추출이 완료되었습니다."
            End If

            SaveArtifacts(result, settings)
            ReportProgress(progress, totalTargets, totalTargets, "완료", "결과 파일 저장을 마쳤습니다.")
            Return result
        End Function

        Public Shared Sub ExportArtifacts(uiapp As UIApplication, settings As Settings, result As RunResult)
            If uiapp Is Nothing Then Throw New ArgumentNullException(NameOf(uiapp))
            If result Is Nothing Then Throw New ArgumentNullException(NameOf(result))

            Dim effectiveSettings = If(settings, New Settings())
            If String.IsNullOrWhiteSpace(result.OutputFolder) Then
                result.OutputFolder = ResolveOutputFolder(uiapp, effectiveSettings)
            End If

            Directory.CreateDirectory(result.OutputFolder)
            SaveArtifacts(result, effectiveSettings)
        End Sub

        Private Shared Function ProcessActiveDocument(doc As Document,
                                                      settings As Settings,
                                                      parameterNames As List(Of String),
                                                      logs As List(Of LogEntry),
                                                      details As List(Of DetailRow),
                                                      progress As IProgress(Of Object)) As FileRunResult
            Dim fileResult As New FileRunResult()
            If doc Is Nothing Then
                fileResult.Status = "Fail"
                fileResult.Message = "활성 문서를 찾을 수 없습니다."
                AddLog(logs, "FAIL", "", fileResult.Message)
                Return fileResult
            End If

            fileResult.FilePath = If(doc.PathName, String.Empty)
            fileResult.FileName = GetDisplayFileName(doc.PathName, doc.Title)
            ReportProgress(progress, 0, 1, "활성 문서 검토 중...", fileResult.FileName)

            Try
                Dim docResult = ProcessDocument(doc, fileResult.FileName, fileResult.FilePath, settings, parameterNames, details, logs)
                ApplyDocumentProcessResult(fileResult, docResult)
                fileResult.Status = If(docResult.ExtractedElementCount > 0, "Success", "NoData")
            Catch ex As Exception
                fileResult.Status = "Fail"
                fileResult.Message = ex.Message
                AddLog(logs, "FAIL", fileResult.FilePath, fileResult.FileName & " 처리 실패: " & ex.Message)
            End Try

            Return fileResult
        End Function

        Private Shared Function ProcessBatchFile(app As RevitApp,
                                                 requestedPath As String,
                                                 settings As Settings,
                                                 parameterNames As List(Of String),
                                                 logs As List(Of LogEntry),
                                                 details As List(Of DetailRow),
                                                 progress As IProgress(Of Object),
                                                 currentIndex As Integer,
                                                 totalCount As Integer) As FileRunResult
            Dim fileResult As New FileRunResult() With {
                .FilePath = requestedPath,
                .FileName = GetDisplayFileName(requestedPath, Path.GetFileName(requestedPath))
            }

            If String.IsNullOrWhiteSpace(requestedPath) OrElse Not File.Exists(requestedPath) Then
                fileResult.Status = "Fail"
                fileResult.Message = "RVT 파일을 찾을 수 없습니다."
                AddLog(logs, "FAIL", requestedPath, fileResult.Message)
                Return fileResult
            End If

            If IsAlreadyOpen(app, requestedPath) Then
                fileResult.Status = "Fail"
                fileResult.Message = "이미 열려 있는 문서입니다."
                AddLog(logs, "FAIL", requestedPath, fileResult.Message)
                Return fileResult
            End If

            Dim doc As Document = Nothing
            Dim fileInfo = TryExtractBasicFileInfo(requestedPath)

            Try
                If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                    fileResult.WasCentralFile = True
                    AddLog(logs, "INFO", requestedPath, "중앙 파일은 Detach + CloseAllWorksets 방식으로 엽니다.")
                End If

                ReportProgress(progress, currentIndex - 1, totalCount, "RVT 열기 중...", fileResult.FileName)
                doc = OpenProjectDocument(app, requestedPath, settings.CloseAllWorksetsOnOpen, fileInfo)
                If doc Is Nothing Then
                    fileResult.Status = "Fail"
                    fileResult.Message = "문서를 열지 못했습니다."
                    AddLog(logs, "FAIL", requestedPath, fileResult.Message)
                    Return fileResult
                End If

                Dim docResult = ProcessDocument(doc, fileResult.FileName, requestedPath, settings, parameterNames, details, logs)
                ApplyDocumentProcessResult(fileResult, docResult)
                fileResult.Status = If(docResult.ExtractedElementCount > 0, "Success", "NoData")
            Catch ex As Exception
                fileResult.Status = "Fail"
                fileResult.Message = ex.Message
                AddLog(logs, "FAIL", requestedPath, fileResult.FileName & " 처리 실패: " & ex.Message)
            Finally
                SafeClose(doc)
            End Try

            Return fileResult
        End Function

        Private Shared Function ProcessDocument(doc As Document,
                                                fileName As String,
                                                filePath As String,
                                                settings As Settings,
                                                parameterNames As List(Of String),
                                                details As List(Of DetailRow),
                                                logs As List(Of LogEntry)) As DocumentProcessResult
            Dim result As New DocumentProcessResult()
            Dim candidates = ModelParameterExtractionService.GetExtractableElements(doc)
            result.CandidateCount = candidates.Count

            Dim filteredElements = candidates.ToList()
            result.FilteredCount = filteredElements.Count

            Dim updateSettings = If(settings.ElementParameterUpdate, New ElementParameterUpdateSettings())
            Dim conditions = GetEnabledConditions(updateSettings)

            For Each element In filteredElements
                If element Is Nothing Then Continue For
                If conditions.Count > 0 AndAlso Not MatchesConditions(doc, element, conditions, updateSettings.CombinationMode) Then
                    Continue For
                End If

                result.MatchedCount += 1

                Dim row As New DetailRow() With {
                    .FilePath = filePath,
                    .FileName = fileName,
                    .ElementId = element.Id.IntegerValue,
                    .Category = ModelParameterExtractionService.GetElementCategoryName(element),
                    .FamilyName = ModelParameterExtractionService.GetElementFamilyName(doc, element),
                    .TypeName = ModelParameterExtractionService.GetElementTypeName(doc, element)
                }

                Dim notes As New List(Of String)()
                If settings.IncludeCoordinates Then
                    Dim point = GetRepresentativePoint(element)
                    If point IsNot Nothing Then
                        row.XValue = ConvertInternalLength(point.X, settings.LengthUnit)
                        row.YValue = ConvertInternalLength(point.Y, settings.LengthUnit)
                        row.ZValue = ConvertInternalLength(point.Z, settings.LengthUnit)
                    Else
                        notes.Add("좌표 없음")
                    End If
                End If

                If settings.IncludeLinearMetrics Then
                    Dim direction As XYZ = Nothing
                    Dim lengthValue As Double = 0.0R
                    If TryGetLinearMetrics(element, direction, lengthValue) Then
                        row.DirectionX = Math.Round(direction.X, 6)
                        row.DirectionY = Math.Round(direction.Y, 6)
                        row.DirectionZ = Math.Round(direction.Z, 6)
                        row.LengthValue = ConvertInternalLength(lengthValue, settings.LengthUnit)
                    Else
                        notes.Add("선형 정보 없음")
                    End If
                End If

                For Each parameterName In parameterNames
                    Dim info = ModelParameterExtractionService.GetElementParameterValueInfo(doc, element, parameterName)
                    row.ParameterTypeTokens(parameterName) = If(info?.DataTypeToken, String.Empty)
                    row.ParameterValues(parameterName) = FormatExtractedParameterValue(info, settings)
                Next

                row.Note = String.Join(" / ", notes.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
                details.Add(row)
                result.ExtractedElementCount += 1
            Next

            If result.ExtractedElementCount > 0 Then
                result.Message = result.ExtractedElementCount.ToString(CultureInfo.InvariantCulture) & "개 객체 추출"
                AddLog(logs, "INFO", filePath, fileName & " " & result.Message)
            Else
                result.Message = "조건에 맞는 객체가 없습니다."
                AddLog(logs, "INFO", filePath, fileName & " 추출 대상 없음")
            End If

            Return result
        End Function

        Private Shared Function MatchesConditions(doc As Document,
                                                  element As Element,
                                                  conditions As List(Of ElementParameterCondition),
                                                  combinationMode As ParameterConditionCombination) As Boolean
            If conditions.Count = 0 Then Return True

            Dim results As New List(Of Boolean)()
            For Each condition In conditions
                results.Add(EvaluateCondition(doc, element, condition))
            Next

            If combinationMode = ParameterConditionCombination.Or Then
                Return results.Any(Function(x) x)
            End If

            Return results.All(Function(x) x)
        End Function

        Private Shared Function EvaluateCondition(doc As Document,
                                                  element As Element,
                                                  condition As ElementParameterCondition) As Boolean
            If condition Is Nothing Then Return True

            Dim hasParameter = ModelParameterExtractionService.HasElementParameter(doc, element, condition.ParameterName)
            If Not hasParameter Then Return False

            Dim actualText = ModelParameterExtractionService.GetElementParameterValue(doc, element, condition.ParameterName)
            Select Case condition.Operator
                Case FilterRuleOperator.HasValue
                    Return Not String.IsNullOrWhiteSpace(actualText)
                Case FilterRuleOperator.HasNoValue
                    Return String.IsNullOrWhiteSpace(actualText)
            End Select

            Dim expectedText = If(condition.Value, String.Empty)
            Dim numericActual As Double
            Dim numericExpected As Double
            Dim hasNumericActual = TryParseNumber(actualText, numericActual)
            Dim hasNumericExpected = TryParseNumber(expectedText, numericExpected)

            If hasNumericActual AndAlso hasNumericExpected Then
                Select Case condition.Operator
                    Case FilterRuleOperator.Equals : Return Math.Abs(numericActual - numericExpected) < 0.000001R
                    Case FilterRuleOperator.NotEquals : Return Math.Abs(numericActual - numericExpected) >= 0.000001R
                    Case FilterRuleOperator.Greater : Return numericActual > numericExpected
                    Case FilterRuleOperator.GreaterOrEqual : Return numericActual >= numericExpected
                    Case FilterRuleOperator.Less : Return numericActual < numericExpected
                    Case FilterRuleOperator.LessOrEqual : Return numericActual <= numericExpected
                End Select
            End If

            Dim left = actualText.Trim()
            Dim right = expectedText.Trim()
            Select Case condition.Operator
                Case FilterRuleOperator.Equals
                    Return String.Equals(left, right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.NotEquals
                    Return Not String.Equals(left, right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.Contains
                    Return left.IndexOf(right, StringComparison.OrdinalIgnoreCase) >= 0
                Case FilterRuleOperator.NotContains
                    Return left.IndexOf(right, StringComparison.OrdinalIgnoreCase) < 0
                Case FilterRuleOperator.BeginsWith
                    Return left.StartsWith(right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.NotBeginsWith
                    Return Not left.StartsWith(right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.EndsWith
                    Return left.EndsWith(right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.NotEndsWith
                    Return Not left.EndsWith(right, StringComparison.OrdinalIgnoreCase)
                Case FilterRuleOperator.Greater
                    Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase) > 0
                Case FilterRuleOperator.GreaterOrEqual
                    Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase) >= 0
                Case FilterRuleOperator.Less
                    Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase) < 0
                Case FilterRuleOperator.LessOrEqual
                    Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase) <= 0
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function GetEnabledConditions(settings As ElementParameterUpdateSettings) As List(Of ElementParameterCondition)
            If settings Is Nothing OrElse settings.Conditions Is Nothing Then Return New List(Of ElementParameterCondition)()
            Return settings.Conditions _
                .Where(Function(x) x IsNot Nothing AndAlso x.IsConfigured() AndAlso (x.Enabled OrElse Not String.IsNullOrWhiteSpace(x.ParameterName))) _
                .Select(Function(x) x.Clone()) _
                .ToList()
        End Function

        Private Shared Function GetExtractParameterNames(settings As Settings) As List(Of String)
            If settings Is Nothing OrElse settings.ExtractParameterNames Is Nothing Then Return New List(Of String)()
            Return settings.ExtractParameterNames _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Select(Function(x) x.Trim()) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
        End Function

        Private Shared Function ValidateSettings(uiapp As UIApplication, settings As Settings) As String
            If settings Is Nothing Then Return "설정을 읽지 못했습니다."
            If settings.ElementParameterUpdate Is Nothing Then settings.ElementParameterUpdate = New ElementParameterUpdateSettings()

            Dim hasValueExtraction = GetExtractParameterNames(settings).Count > 0
            If Not hasValueExtraction AndAlso Not settings.IncludeCoordinates AndAlso Not settings.IncludeLinearMetrics Then
                Return "추출 파라미터 또는 좌표/선형 옵션을 1개 이상 선택하세요."
            End If

            If settings.UseActiveDocument Then
                Dim doc = uiapp.ActiveUIDocument?.Document
                If doc Is Nothing Then Return "활성 문서를 찾을 수 없습니다."
                If doc.IsFamilyDocument Then Return "패밀리 문서는 검토 대상에서 제외됩니다."
                Return String.Empty
            End If

            Dim targetPaths = If(settings.RvtPaths, New List(Of String)()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
            If targetPaths.Count = 0 Then Return "RVT 파일을 하나 이상 추가하세요."

            Return String.Empty
        End Function

        Private Shared Function ResolveOutputFolder(uiapp As UIApplication, settings As Settings) As String
            Dim outputFolder = If(settings.OutputFolder, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(outputFolder) Then Return outputFolder

            If settings.UseActiveDocument Then
                Dim doc = uiapp.ActiveUIDocument?.Document
                If doc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(doc.PathName) Then
                    Return Path.GetDirectoryName(doc.PathName)
                End If
            ElseIf settings.RvtPaths IsNot Nothing Then
                Dim firstPath = settings.RvtPaths.FirstOrDefault(Function(x) Not String.IsNullOrWhiteSpace(x))
                If Not String.IsNullOrWhiteSpace(firstPath) AndAlso File.Exists(firstPath) Then
                    Return Path.GetDirectoryName(firstPath)
                End If
            End If

            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Function

        Private Shared Function BuildSummary(fileResults As IEnumerable(Of FileRunResult)) As RunSummary
            Dim results = If(fileResults, Enumerable.Empty(Of FileRunResult)()).Where(Function(x) x IsNot Nothing).ToList()
            Return New RunSummary With {
                .SuccessCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "Success", StringComparison.OrdinalIgnoreCase)),
                .FailCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "Fail", StringComparison.OrdinalIgnoreCase)),
                .NoDataCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "NoData", StringComparison.OrdinalIgnoreCase)),
                .TotalCandidateCount = results.Sum(Function(x) x.CandidateCount),
                .TotalFilteredCount = results.Sum(Function(x) x.FilteredCount),
                .TotalMatchedCount = results.Sum(Function(x) x.MatchedCount),
                .TotalExtractedElementCount = results.Sum(Function(x) x.ExtractedElementCount)
            }
        End Function

        Private Shared Sub SaveArtifacts(result As RunResult, settings As Settings)
            If result Is Nothing Then Return

            Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            Dim workbookPath = Path.Combine(result.OutputFolder, "ConditionExtractResult_" & timeStamp & ".xlsx")
            Dim logPath = Path.Combine(result.OutputFolder, "ConditionExtractLog_" & timeStamp & ".txt")
            Dim parameterNames = GetExtractParameterNames(settings)

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable)) From {
                New KeyValuePair(Of String, DataTable)("Summary", BuildSummaryTable(result.Files)),
                New KeyValuePair(Of String, DataTable)("Detail", BuildDetailTable(result.Details, parameterNames, settings)),
                New KeyValuePair(Of String, DataTable)("Logs", BuildLogTable(result.Logs))
            }

            ExcelCore.SaveXlsxMulti(workbookPath, sheets, autoFit:=True)
            File.WriteAllLines(logPath, BuildLogLines(result.Logs), New UTF8Encoding(True))

            result.ResultWorkbookPath = workbookPath
            result.LogTextPath = logPath
        End Sub

        Private Shared Function BuildSummaryTable(fileResults As IEnumerable(Of FileRunResult)) As DataTable
            Dim table As New DataTable("Summary")
            table.Columns.Add("FileName", GetType(String))
            table.Columns.Add("Status", GetType(String))
            table.Columns.Add("CandidateCount", GetType(Integer))
            table.Columns.Add("FilteredCount", GetType(Integer))
            table.Columns.Add("MatchedCount", GetType(Integer))
            table.Columns.Add("ExtractedElementCount", GetType(Integer))
            table.Columns.Add("CentralToLocal", GetType(String))
            table.Columns.Add("Message", GetType(String))

            For Each item In If(fileResults, Enumerable.Empty(Of FileRunResult)())
                If item Is Nothing Then Continue For
                table.Rows.Add(item.FileName, item.Status, item.CandidateCount, item.FilteredCount, item.MatchedCount, item.ExtractedElementCount, If(item.UsedLocalFile, "Y", "N"), item.Message)
            Next

            Return table
        End Function

        Private Shared Function BuildDetailTable(detailRows As IEnumerable(Of DetailRow),
                                                 parameterNames As IEnumerable(Of String),
                                                 settings As Settings) As DataTable
            Dim table As New DataTable("Detail")
            table.Columns.Add("FileName", GetType(String))
            table.Columns.Add("ElementId", GetType(Integer))
            table.Columns.Add("Category", GetType(String))
            table.Columns.Add("FamilyName", GetType(String))
            table.Columns.Add("TypeName", GetType(String))
            table.Columns.Add("Note", GetType(String))

            Dim lengthUnitLabel = GetLengthUnitLabel(settings)
            Dim areaUnitLabel = GetAreaUnitLabel(settings)
            Dim volumeUnitLabel = GetVolumeUnitLabel(settings)

            If settings IsNot Nothing AndAlso settings.IncludeCoordinates Then
                table.Columns.Add("X(" & lengthUnitLabel & ")", GetType(Double))
                table.Columns.Add("Y(" & lengthUnitLabel & ")", GetType(Double))
                table.Columns.Add("Z(" & lengthUnitLabel & ")", GetType(Double))
            End If

            If settings IsNot Nothing AndAlso settings.IncludeLinearMetrics Then
                table.Columns.Add("DirectionX", GetType(Double))
                table.Columns.Add("DirectionY", GetType(Double))
                table.Columns.Add("DirectionZ", GetType(Double))
                table.Columns.Add("Length(" & lengthUnitLabel & ")", GetType(Double))
            End If

            Dim distinctParamNames = If(parameterNames, Enumerable.Empty(Of String)()).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            Dim parameterColumnMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each name In distinctParamNames
                Dim token = ResolveParameterDataTypeToken(detailRows, name)
                Dim columnName = BuildParameterColumnName(name, token, lengthUnitLabel, areaUnitLabel, volumeUnitLabel)
                table.Columns.Add(columnName, GetType(String))
                parameterColumnMap(name) = columnName
            Next

            For Each rowItem In If(detailRows, Enumerable.Empty(Of DetailRow)())
                If rowItem Is Nothing Then Continue For
                Dim row = table.NewRow()
                row("FileName") = rowItem.FileName
                row("ElementId") = rowItem.ElementId
                row("Category") = rowItem.Category
                row("FamilyName") = rowItem.FamilyName
                row("TypeName") = rowItem.TypeName
                row("Note") = rowItem.Note

                If table.Columns.Contains("X(" & lengthUnitLabel & ")") Then row("X(" & lengthUnitLabel & ")") = If(rowItem.XValue.HasValue, CType(rowItem.XValue.Value, Object), DBNull.Value)
                If table.Columns.Contains("Y(" & lengthUnitLabel & ")") Then row("Y(" & lengthUnitLabel & ")") = If(rowItem.YValue.HasValue, CType(rowItem.YValue.Value, Object), DBNull.Value)
                If table.Columns.Contains("Z(" & lengthUnitLabel & ")") Then row("Z(" & lengthUnitLabel & ")") = If(rowItem.ZValue.HasValue, CType(rowItem.ZValue.Value, Object), DBNull.Value)
                If table.Columns.Contains("DirectionX") Then row("DirectionX") = If(rowItem.DirectionX.HasValue, CType(rowItem.DirectionX.Value, Object), DBNull.Value)
                If table.Columns.Contains("DirectionY") Then row("DirectionY") = If(rowItem.DirectionY.HasValue, CType(rowItem.DirectionY.Value, Object), DBNull.Value)
                If table.Columns.Contains("DirectionZ") Then row("DirectionZ") = If(rowItem.DirectionZ.HasValue, CType(rowItem.DirectionZ.Value, Object), DBNull.Value)
                If table.Columns.Contains("Length(" & lengthUnitLabel & ")") Then row("Length(" & lengthUnitLabel & ")") = If(rowItem.LengthValue.HasValue, CType(rowItem.LengthValue.Value, Object), DBNull.Value)

                For Each name In distinctParamNames
                    Dim columnName = parameterColumnMap(name)
                    row(columnName) = If(rowItem.ParameterValues IsNot Nothing AndAlso rowItem.ParameterValues.ContainsKey(name), rowItem.ParameterValues(name), String.Empty)
                Next

                table.Rows.Add(row)
            Next

            Return table
        End Function

        Private Shared Function BuildLogTable(logs As IEnumerable(Of LogEntry)) As DataTable
            Dim table As New DataTable("Logs")
            table.Columns.Add("Level", GetType(String))
            table.Columns.Add("FilePath", GetType(String))
            table.Columns.Add("Message", GetType(String))

            For Each logItem In If(logs, Enumerable.Empty(Of LogEntry)())
                If logItem Is Nothing Then Continue For
                table.Rows.Add(logItem.Level, logItem.FilePath, logItem.Message)
            Next

            Return table
        End Function

        Private Shared Function BuildLogLines(logs As IEnumerable(Of LogEntry)) As IEnumerable(Of String)
            Return If(logs, Enumerable.Empty(Of LogEntry)()) _
                .Select(Function(x) "[" & x.Level & "] " & If(x.FilePath, "") & " :: " & If(x.Message, "")) _
                .ToList()
        End Function

        Private Shared Sub ApplyDocumentProcessResult(fileResult As FileRunResult, docResult As DocumentProcessResult)
            fileResult.CandidateCount = docResult.CandidateCount
            fileResult.FilteredCount = docResult.FilteredCount
            fileResult.MatchedCount = docResult.MatchedCount
            fileResult.ExtractedElementCount = docResult.ExtractedElementCount
            fileResult.Message = docResult.Message
        End Sub

        Private Shared Sub AddLog(logs As List(Of LogEntry), level As String, filePath As String, message As String)
            If logs Is Nothing OrElse String.IsNullOrWhiteSpace(message) Then Return
            logs.Add(New LogEntry With {
                .Level = If(level, "INFO"),
                .FilePath = If(filePath, String.Empty),
                .Message = message
            })
        End Sub

        Private Shared Sub ReportProgress(progress As IProgress(Of Object),
                                          current As Integer,
                                          total As Integer,
                                          message As String,
                                          detail As String)
            If progress Is Nothing Then Return

            Dim percent As Double = 0.0R
            If total > 0 Then
                percent = Math.Max(0.0R, Math.Min(100.0R, (CDbl(current) / CDbl(total)) * 100.0R))
            End If

            progress.Report(New With {
                .title = "조건별 객체 대상 속성 추출",
                .message = If(message, String.Empty),
                .detail = If(detail, String.Empty),
                .current = current,
                .total = total,
                .percent = percent
            })
        End Sub

        Private Shared Function GetRepresentativePoint(element As Element) As XYZ
            If element Is Nothing Then Return Nothing

            Dim location = element.Location
            If TypeOf location Is LocationPoint Then
                Return DirectCast(location, LocationPoint).Point
            End If

            If TypeOf location Is LocationCurve Then
                Dim curve = DirectCast(location, LocationCurve).Curve
                If curve IsNot Nothing Then
                    Try
                        Dim startPoint = curve.GetEndPoint(0)
                        Dim endPoint = curve.GetEndPoint(1)
                        Return New XYZ((startPoint.X + endPoint.X) / 2.0R,
                                       (startPoint.Y + endPoint.Y) / 2.0R,
                                       (startPoint.Z + endPoint.Z) / 2.0R)
                    Catch
                    End Try

                    Try
                        Return curve.Evaluate(0.5R, True)
                    Catch
                    End Try
                End If
            End If

            Try
                Dim bbox = element.BoundingBox(Nothing)
                If bbox IsNot Nothing AndAlso bbox.Min IsNot Nothing AndAlso bbox.Max IsNot Nothing Then
                    Return New XYZ((bbox.Min.X + bbox.Max.X) / 2.0R,
                                   (bbox.Min.Y + bbox.Max.Y) / 2.0R,
                                   (bbox.Min.Z + bbox.Max.Z) / 2.0R)
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function TryGetLinearMetrics(element As Element,
                                                    ByRef direction As XYZ,
                                                    ByRef lengthValue As Double) As Boolean
            direction = Nothing
            lengthValue = 0.0R
            If element Is Nothing Then Return False

            Dim locationCurve = TryCast(element.Location, LocationCurve)
            If locationCurve Is Nothing OrElse locationCurve.Curve Is Nothing Then Return False

            Try
                Dim curve = locationCurve.Curve
                Dim startPoint = curve.GetEndPoint(0)
                Dim endPoint = curve.GetEndPoint(1)
                Dim vector = endPoint - startPoint
                If vector Is Nothing OrElse vector.GetLength() <= 0.000001R Then Return False

                direction = vector.Normalize()
                lengthValue = curve.Length
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ResolveParameterDataTypeToken(detailRows As IEnumerable(Of DetailRow), parameterName As String) As String
            For Each rowItem In If(detailRows, Enumerable.Empty(Of DetailRow)())
                If rowItem?.ParameterTypeTokens Is Nothing Then Continue For
                If Not rowItem.ParameterTypeTokens.ContainsKey(parameterName) Then Continue For

                Dim token = If(rowItem.ParameterTypeTokens(parameterName), String.Empty)
                If Not String.IsNullOrWhiteSpace(token) Then Return token
            Next

            Return String.Empty
        End Function

        Private Shared Function BuildParameterColumnName(parameterName As String,
                                                         dataTypeToken As String,
                                                         lengthUnitLabel As String,
                                                         areaUnitLabel As String,
                                                         volumeUnitLabel As String) As String
            If IsLengthDataTypeToken(dataTypeToken) Then
                Return parameterName & " (" & lengthUnitLabel & ")"
            End If

            If IsAreaDataTypeToken(dataTypeToken) Then
                Return parameterName & " (" & areaUnitLabel & ")"
            End If

            If IsVolumeDataTypeToken(dataTypeToken) Then
                Return parameterName & " (" & volumeUnitLabel & ")"
            End If

            Return parameterName
        End Function

        Private Shared Function FormatExtractedParameterValue(info As ModelParameterExtractionService.ElementParameterValueInfo,
                                                              settings As Settings) As String
            If info Is Nothing OrElse Not info.HasParameter Then Return String.Empty

            Dim rawText = If(info.ValueText, String.Empty)
            If Not info.InternalDoubleValue.HasValue Then Return rawText

            If IsLengthDataTypeToken(info.DataTypeToken) Then
                Return FormatConvertedNumber(ConvertInternalLength(info.InternalDoubleValue.Value, settings.LengthUnit))
            End If

            If IsAreaDataTypeToken(info.DataTypeToken) Then
                Return FormatConvertedNumber(ConvertInternalArea(info.InternalDoubleValue.Value, settings.AreaUnit))
            End If

            If IsVolumeDataTypeToken(info.DataTypeToken) Then
                Return FormatConvertedNumber(ConvertInternalVolume(info.InternalDoubleValue.Value, settings.VolumeUnit))
            End If

            Return rawText
        End Function

        Private Shared Function FormatConvertedNumber(value As Double) As String
            Return value.ToString("0.###", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function ConvertInternalLength(internalLength As Double, unitName As String) As Double
            Dim normalized = NormalizeLengthUnit(unitName)
#If REVIT2021 Or REVIT2023 Or REVIT2025 Then
            Select Case normalized
                Case "inch"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, UnitTypeId.Inches), 3)
                Case "ft"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, UnitTypeId.Feet), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, UnitTypeId.Millimeters), 3)
            End Select
#Else
            Select Case normalized
                Case "inch"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, DisplayUnitType.DUT_DECIMAL_INCHES), 3)
                Case "ft"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, DisplayUnitType.DUT_DECIMAL_FEET), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalLength, DisplayUnitType.DUT_MILLIMETERS), 3)
            End Select
#End If
        End Function

        Private Shared Function ConvertInternalArea(internalArea As Double, unitName As String) As Double
            Dim normalized = NormalizeAreaUnit(unitName)
#If REVIT2021 Or REVIT2023 Or REVIT2025 Then
            Select Case normalized
                Case "in2"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, UnitTypeId.SquareInches), 3)
                Case "ft2"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, UnitTypeId.SquareFeet), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, UnitTypeId.SquareMillimeters), 3)
            End Select
#Else
            Select Case normalized
                Case "in2"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, DisplayUnitType.DUT_SQUARE_INCHES), 3)
                Case "ft2"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, DisplayUnitType.DUT_SQUARE_FEET), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalArea, DisplayUnitType.DUT_SQUARE_MILLIMETERS), 3)
            End Select
#End If
        End Function

        Private Shared Function ConvertInternalVolume(internalVolume As Double, unitName As String) As Double
            Dim normalized = NormalizeVolumeUnit(unitName)
#If REVIT2021 Or REVIT2023 Or REVIT2025 Then
            Select Case normalized
                Case "in3"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, UnitTypeId.CubicInches), 3)
                Case "ft3"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, UnitTypeId.CubicFeet), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, UnitTypeId.CubicMillimeters), 3)
            End Select
#Else
            Select Case normalized
                Case "in3"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, DisplayUnitType.DUT_CUBIC_INCHES), 3)
                Case "ft3"
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, DisplayUnitType.DUT_CUBIC_FEET), 3)
                Case Else
                    Return Math.Round(UnitUtils.ConvertFromInternalUnits(internalVolume, DisplayUnitType.DUT_CUBIC_MILLIMETERS), 3)
            End Select
#End If
        End Function

        Private Shared Function GetLengthUnitLabel(settings As Settings) As String
            Return NormalizeLengthUnit(If(settings?.LengthUnit, "mm"))
        End Function

        Private Shared Function GetAreaUnitLabel(settings As Settings) As String
            Select Case NormalizeAreaUnit(If(settings?.AreaUnit, "mm2"))
                Case "in2" : Return "inch^2"
                Case "ft2" : Return "ft^2"
                Case Else : Return "mm^2"
            End Select
        End Function

        Private Shared Function GetVolumeUnitLabel(settings As Settings) As String
            Select Case NormalizeVolumeUnit(If(settings?.VolumeUnit, "mm3"))
                Case "in3" : Return "inch^3"
                Case "ft3" : Return "ft^3"
                Case Else : Return "mm^3"
            End Select
        End Function

        Private Shared Function NormalizeLengthUnit(unitName As String) As String
            Dim raw = If(unitName, String.Empty).Trim().ToLowerInvariant()
            Select Case raw
                Case "inch", "in", "inches"
                    Return "inch"
                Case "ft", "foot", "feet"
                    Return "ft"
                Case Else
                    Return "mm"
            End Select
        End Function

        Private Shared Function NormalizeAreaUnit(unitName As String) As String
            Dim raw = If(unitName, String.Empty).Trim().ToLowerInvariant()
            Select Case raw
                Case "in2", "inch2", "in^2", "inch^2", "sqin", "squareinch", "squareinches"
                    Return "in2"
                Case "ft2", "ft^2", "sqft", "squarefoot", "squarefeet"
                    Return "ft2"
                Case Else
                    Return "mm2"
            End Select
        End Function

        Private Shared Function NormalizeVolumeUnit(unitName As String) As String
            Dim raw = If(unitName, String.Empty).Trim().ToLowerInvariant()
            Select Case raw
                Case "in3", "inch3", "in^3", "inch^3", "cuin", "cubicinch", "cubicinches"
                    Return "in3"
                Case "ft3", "ft^3", "cuft", "cubicfoot", "cubicfeet"
                    Return "ft3"
                Case Else
                    Return "mm3"
            End Select
        End Function

        Private Shared Function IsLengthDataTypeToken(dataTypeToken As String) As Boolean
            Dim normalized = NormalizeDataTypeToken(dataTypeToken)
            Return normalized.Contains("length")
        End Function

        Private Shared Function IsAreaDataTypeToken(dataTypeToken As String) As Boolean
            Dim normalized = NormalizeDataTypeToken(dataTypeToken)
            Return normalized.Contains("area")
        End Function

        Private Shared Function IsVolumeDataTypeToken(dataTypeToken As String) As Boolean
            Dim normalized = NormalizeDataTypeToken(dataTypeToken)
            Return normalized.Contains("volume")
        End Function

        Private Shared Function NormalizeDataTypeToken(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            Dim builder As New StringBuilder()
            For Each ch In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) Then builder.Append(ch)
            Next

            Return builder.ToString()
        End Function

        Private Shared Function TryParseNumber(text As String, ByRef value As Double) As Boolean
            Dim raw = If(text, String.Empty).Trim()
            If Double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, value) Then Return True
            If Double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, value) Then Return True
            Return False
        End Function

        Private Shared Function GetDisplayFileName(filePath As String, fallback As String) As String
            If Not String.IsNullOrWhiteSpace(filePath) Then
                Try
                    Return Path.GetFileName(filePath)
                Catch
                End Try
            End If
            Return If(fallback, String.Empty)
        End Function

        Private Shared Function OpenProjectDocument(app As RevitApp,
                                                    userVisiblePath As String,
                                                    closeAllWorksets As Boolean,
                                                    fileInfo As BasicFileInfo) As Document
            Dim modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)
            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                openOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            End If

            If closeAllWorksets Then
                Dim worksetConfig As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                openOpts.SetOpenWorksetsConfiguration(worksetConfig)
            End If

            Return app.OpenDocumentFile(modelPath, openOpts)
        End Function

        Private Shared Function TryExtractBasicFileInfo(pathText As String) As BasicFileInfo
            Try
                Return BasicFileInfo.Extract(pathText)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function IsAlreadyOpen(app As RevitApp, userVisiblePath As String) As Boolean
            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For
                    If String.Equals(doc.PathName, userVisiblePath, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
            End Try
            Return False
        End Function

        Private Shared Sub SafeClose(doc As Document)
            If doc Is Nothing Then Return
            Try
                doc.Close(False)
            Catch
            End Try
        End Sub

        Private Shared Sub TryDeleteFile(pathText As String)
            If String.IsNullOrWhiteSpace(pathText) OrElse Not File.Exists(pathText) Then Return
            Try
                File.Delete(pathText)
            Catch
            End Try
        End Sub

    End Class

End Namespace
