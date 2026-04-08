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

    Public NotInheritable Class ParameterModifierService

        Private Sub New()
        End Sub

        Public Class Settings
            Public Property UseActiveDocument As Boolean
            Public Property RvtPaths As List(Of String) = New List(Of String)()
            Public Property OutputFolder As String = String.Empty
            Public Property CloseAllWorksetsOnOpen As Boolean = True
            Public Property SynchronizeAfterProcessing As Boolean = True
            Public Property SyncComment As String = String.Empty
            Public Property FilterProfile As ViewFilterProfile
            Public Property ElementParameterUpdate As ElementParameterUpdateSettings
        End Class

        Public Class RunSummary
            Public Property SuccessCount As Integer
            Public Property FailCount As Integer
            Public Property NoChangeCount As Integer
            Public Property TotalCandidateCount As Integer
            Public Property TotalFilteredCount As Integer
            Public Property TotalMatchedCount As Integer
            Public Property TotalUpdatedElementCount As Integer
            Public Property TotalUpdatedParameterCount As Integer
        End Class

        Public Class FileRunResult
            Public Property FilePath As String = String.Empty
            Public Property FileName As String = String.Empty
            Public Property Status As String = String.Empty
            Public Property CandidateCount As Integer
            Public Property FilteredCount As Integer
            Public Property MatchedCount As Integer
            Public Property UpdatedElementCount As Integer
            Public Property UpdatedParameterCount As Integer
            Public Property SynchronizeRequested As Boolean
            Public Property SynchronizePerformed As Boolean
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
            Public Property Status As String = String.Empty
            Public Property Note As String = String.Empty
            Public Property ParameterValues As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
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
            Public Property UpdatedElementCount As Integer
            Public Property UpdatedParameterCount As Integer
            Public Property HasChanges As Boolean
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

            Dim assignmentNames = GetAssignmentParameterNames(settings)
            Dim app As RevitApp = uiapp.Application
            Dim totalTargets As Integer = If(settings.UseActiveDocument, 1, settings.RvtPaths.Count)
            ReportProgress(progress, 0, totalTargets, "파라미터 수정기 준비 중...", result.OutputFolder)

            If settings.UseActiveDocument Then
                Dim activeDoc = uiapp.ActiveUIDocument?.Document
                Dim fileResult = ProcessActiveDocument(activeDoc, settings, assignmentNames, result.Logs, result.Details, progress)
                result.Files.Add(fileResult)
            Else
                Dim distinctPaths = settings.RvtPaths _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()

                For i As Integer = 0 To distinctPaths.Count - 1
                    Dim pathText = distinctPaths(i)
                    ReportProgress(progress, i, totalTargets, "파일 준비 중...", pathText)
                    Dim fileResult = ProcessBatchFile(app, pathText, settings, assignmentNames, result.Logs, result.Details, progress, i + 1, totalTargets)
                    result.Files.Add(fileResult)
                Next
            End If

            result.Summary = BuildSummary(result.Files)
            result.Ok = result.Files.Any() AndAlso result.Files.Any(Function(x) String.Equals(x.Status, "Success", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(x.Status, "NoChange", StringComparison.OrdinalIgnoreCase))
            If result.Files.Count = 0 Then
                result.Message = "처리할 대상이 없습니다."
            ElseIf result.Summary.FailCount > 0 AndAlso result.Summary.SuccessCount = 0 AndAlso result.Summary.NoChangeCount = 0 Then
                result.Message = "모든 대상 처리에 실패했습니다."
            Else
                result.Message = "파라미터 수정 작업이 완료되었습니다."
            End If

            ReportProgress(progress, totalTargets, totalTargets, "완료", "결과 창을 준비하고 있습니다.")
            Return result
        End Function

        Public Shared Sub ExportArtifacts(uiapp As UIApplication, settings As Settings, result As RunResult)
            If uiapp Is Nothing Then Throw New ArgumentNullException(NameOf(uiapp))
            If result Is Nothing Then Throw New ArgumentNullException(NameOf(result))

            Dim effectiveSettings = If(settings, New Settings())
            If effectiveSettings.ElementParameterUpdate Is Nothing Then
                effectiveSettings.ElementParameterUpdate = New ElementParameterUpdateSettings()
            End If

            Dim hasWorkbook = Not String.IsNullOrWhiteSpace(result.ResultWorkbookPath) AndAlso File.Exists(result.ResultWorkbookPath)
            Dim hasLog = Not String.IsNullOrWhiteSpace(result.LogTextPath) AndAlso File.Exists(result.LogTextPath)
            If hasWorkbook AndAlso hasLog Then Return

            If String.IsNullOrWhiteSpace(result.OutputFolder) Then
                result.OutputFolder = ResolveOutputFolder(uiapp, effectiveSettings)
            End If
            Directory.CreateDirectory(result.OutputFolder)

            Dim assignmentNames = GetAssignmentParameterNames(effectiveSettings)
            SaveArtifacts(result, assignmentNames)
        End Sub

        Private Shared Function ProcessActiveDocument(doc As Document,
                                                      settings As Settings,
                                                      assignmentNames As List(Of String),
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
            fileResult.SynchronizeRequested = settings.SynchronizeAfterProcessing
            ReportProgress(progress, 0, 1, "활성 문서 처리 중...", fileResult.FileName)

            Try
                Dim docResult = ProcessDocument(doc, fileResult.FileName, fileResult.FilePath, settings, assignmentNames, details, logs)
                ApplyDocumentProcessResult(fileResult, docResult)

                If docResult.HasChanges AndAlso settings.SynchronizeAfterProcessing Then
                    If doc.IsWorkshared Then
                        Dim syncError As String = String.Empty
                        If SyncWithCentral(doc, settings.SyncComment, syncError) Then
                            fileResult.SynchronizePerformed = True
                            fileResult.Message = AppendMessage(fileResult.Message, "동기화 완료")
                            AddLog(logs, "OK", fileResult.FilePath, fileResult.FileName & " 동기화 완료")
                        Else
                            fileResult.Status = "Fail"
                            fileResult.Message = AppendMessage(fileResult.Message, "동기화 실패: " & syncError)
                            AddLog(logs, "FAIL", fileResult.FilePath, fileResult.FileName & " 동기화 실패: " & syncError)
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(doc.PathName) Then
                        doc.Save()
                        fileResult.SynchronizePerformed = True
                        fileResult.Message = AppendMessage(fileResult.Message, "저장 완료")
                        AddLog(logs, "OK", fileResult.FilePath, fileResult.FileName & " 저장 완료")
                    End If
                End If

                If String.IsNullOrWhiteSpace(fileResult.Status) Then
                    fileResult.Status = If(docResult.HasChanges, "Success", "NoChange")
                End If
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
                                                 assignmentNames As List(Of String),
                                                 logs As List(Of LogEntry),
                                                 details As List(Of DetailRow),
                                                 progress As IProgress(Of Object),
                                                 currentIndex As Integer,
                                                 totalCount As Integer) As FileRunResult
            Dim fileResult As New FileRunResult() With {
                .FilePath = requestedPath,
                .FileName = GetDisplayFileName(requestedPath, Path.GetFileName(requestedPath)),
                .SynchronizeRequested = True
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

            Dim openPath = requestedPath
            Dim createdLocal = False
            Dim doc As Document = Nothing

            Try
                Dim fileInfo = TryExtractBasicFileInfo(requestedPath)
                If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                    openPath = CreateNewLocalPath(requestedPath)
                    fileResult.WasCentralFile = True
                    fileResult.UsedLocalFile = True
                    createdLocal = True
                    AddLog(logs, "INFO", requestedPath, "중앙파일을 로컬파일로 생성했습니다: " & openPath)
                End If

                ReportProgress(progress, currentIndex - 1, totalCount, "RVT 열기 중...", fileResult.FileName)
                doc = OpenProjectDocument(app, openPath, settings.CloseAllWorksetsOnOpen)
                If doc Is Nothing Then
                    fileResult.Status = "Fail"
                    fileResult.Message = "문서를 열지 못했습니다."
                    AddLog(logs, "FAIL", requestedPath, fileResult.Message)
                    Return fileResult
                End If

                Dim docResult = ProcessDocument(doc, fileResult.FileName, requestedPath, settings, assignmentNames, details, logs)
                ApplyDocumentProcessResult(fileResult, docResult)

                If docResult.HasChanges Then
                    If doc.IsWorkshared Then
                        Dim syncError As String = String.Empty
                        If SyncWithCentral(doc, settings.SyncComment, syncError) Then
                            fileResult.SynchronizePerformed = True
                            fileResult.Message = AppendMessage(fileResult.Message, "동기화 완료")
                            AddLog(logs, "OK", requestedPath, fileResult.FileName & " 동기화 완료")
                        Else
                            fileResult.Status = "Fail"
                            fileResult.Message = AppendMessage(fileResult.Message, "동기화 실패: " & syncError)
                            AddLog(logs, "FAIL", requestedPath, fileResult.FileName & " 동기화 실패: " & syncError)
                        End If
                    Else
                        doc.Save()
                        fileResult.SynchronizePerformed = True
                        fileResult.Message = AppendMessage(fileResult.Message, "저장 완료")
                        AddLog(logs, "OK", requestedPath, fileResult.FileName & " 저장 완료")
                    End If
                End If

                If String.IsNullOrWhiteSpace(fileResult.Status) Then
                    fileResult.Status = If(docResult.HasChanges, "Success", "NoChange")
                End If
            Catch ex As Exception
                fileResult.Status = "Fail"
                fileResult.Message = ex.Message
                AddLog(logs, "FAIL", requestedPath, fileResult.FileName & " 처리 실패: " & ex.Message)
            Finally
                SafeClose(doc)
                If createdLocal Then
                    TryDeleteFile(openPath)
                End If
            End Try

            Return fileResult
        End Function

        Private Shared Function ProcessDocument(doc As Document,
                                                fileName As String,
                                                filePath As String,
                                                settings As Settings,
                                                assignmentNames As List(Of String),
                                                details As List(Of DetailRow),
                                                logs As List(Of LogEntry)) As DocumentProcessResult
            Dim result As New DocumentProcessResult()
            Dim candidates = ModelParameterExtractionService.GetExtractableElements(doc)
            result.CandidateCount = candidates.Count

            Dim filteredElements = candidates.ToList()
            If settings.FilterProfile IsNot Nothing AndAlso settings.FilterProfile.IsConfigured() Then
                Dim candidateIds = filteredElements.Select(Function(x) x.Id).ToList()
                Dim matchedIds = RevitViewFilterProfileService.GetMatchingElementIds(doc, settings.FilterProfile, candidateIds, Nothing)
                Dim matchedSet As New HashSet(Of Integer)(matchedIds.Select(Function(x) x.IntegerValue))
                filteredElements = filteredElements.Where(Function(x) matchedSet.Contains(x.Id.IntegerValue)).ToList()
            End If
            result.FilteredCount = filteredElements.Count

            Dim updateSettings = settings.ElementParameterUpdate
            Dim assignments = GetEnabledAssignments(updateSettings)
            Dim conditions = GetEnabledConditions(updateSettings)
            result.MatchedCount = 0

            Using tx As New Transaction(doc, "KKY Parameter Modifier")
                tx.Start()

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

                    Dim updatedCountForElement As Integer = 0
                    Dim notes As New List(Of String)()
                    For Each assignment In assignments
                        Dim parameter = FindParameterOnElementOrType(doc, element, assignment.ParameterName)
                        If parameter Is Nothing Then
                            notes.Add(assignment.ParameterName & ": 파라미터 없음")
                            Continue For
                        End If

                        If parameter.IsReadOnly Then
                            notes.Add(assignment.ParameterName & ": 읽기 전용")
                            Continue For
                        End If

                        Dim updateMessage As String = String.Empty
                        Dim changed = TrySetParameterValue(doc, parameter, assignment.Value, updateMessage)
                        If changed Then
                            updatedCountForElement += 1
                            result.UpdatedParameterCount += 1
                        ElseIf Not String.IsNullOrWhiteSpace(updateMessage) Then
                            notes.Add(assignment.ParameterName & ": " & updateMessage)
                        End If
                    Next

                    If updatedCountForElement > 0 Then
                        result.UpdatedElementCount += 1
                        result.HasChanges = True
                    End If

                    row.Status = If(updatedCountForElement > 0, "Updated", "Skipped")
                    row.Note = String.Join(" / ", notes.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))

                    For Each parameterName In assignmentNames
                        row.ParameterValues(parameterName) = ModelParameterExtractionService.GetElementParameterValue(doc, element, parameterName)
                    Next

                    If updatedCountForElement > 0 OrElse notes.Count > 0 Then
                        details.Add(row)
                    End If
                Next

                If result.HasChanges Then
                    tx.Commit()
                    result.Message = result.UpdatedElementCount.ToString() & "개 객체, " & result.UpdatedParameterCount.ToString() & "개 파라미터 입력"
                    AddLog(logs, "INFO", filePath, fileName & " " & result.Message)
                Else
                    tx.RollBack()
                    result.Message = "변경된 값이 없습니다."
                    AddLog(logs, "INFO", filePath, fileName & " 변경 없음")
                End If
            End Using

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

        Private Shared Function TrySetParameterValue(doc As Document,
                                                     parameter As Parameter,
                                                     valueText As String,
                                                     ByRef message As String) As Boolean
            message = String.Empty
            If parameter Is Nothing Then
                message = "파라미터를 찾을 수 없습니다."
                Return False
            End If

            Try
                Select Case parameter.StorageType
                    Case StorageType.String
                        Dim beforeText = parameter.AsString()
                        Dim afterText = If(valueText, String.Empty)
                        If String.Equals(beforeText, afterText, StringComparison.Ordinal) Then
                            message = "이미 같은 값"
                            Return False
                        End If
                        parameter.Set(afterText)
                        Return True

                    Case StorageType.Integer
                        Dim intValue As Integer
                        If Not TryParseInteger(valueText, intValue) Then
                            message = "정수 값으로 변환할 수 없습니다."
                            Return False
                        End If
                        If parameter.AsInteger() = intValue Then
                            message = "이미 같은 값"
                            Return False
                        End If
                        parameter.Set(intValue)
                        Return True

                    Case StorageType.Double
                        Dim beforeDouble = parameter.AsDouble()
                        Dim setByString As Boolean = False
                        Try
                            setByString = parameter.SetValueString(If(valueText, String.Empty))
                        Catch
                            setByString = False
                        End Try
                        If setByString Then
                            If Math.Abs(parameter.AsDouble() - beforeDouble) < 0.000001R Then
                                message = "이미 같은 값"
                                Return False
                            End If
                            Return True
                        End If

                        Dim doubleValue As Double
                        If Not TryParseNumber(valueText, doubleValue) Then
                            message = "숫자 값으로 변환할 수 없습니다."
                            Return False
                        End If
                        If Math.Abs(beforeDouble - doubleValue) < 0.000001R Then
                            message = "이미 같은 값"
                            Return False
                        End If
                        parameter.Set(doubleValue)
                        Return True

                    Case StorageType.ElementId
                        Dim intValue As Integer
                        If Not Integer.TryParse(If(valueText, String.Empty).Trim(), intValue) Then
                            message = "ElementId 정수값이 필요합니다."
                            Return False
                        End If
                        Dim beforeId = parameter.AsElementId()
                        If beforeId IsNot Nothing AndAlso beforeId.IntegerValue = intValue Then
                            message = "이미 같은 값"
                            Return False
                        End If
                        parameter.Set(New ElementId(intValue))
                        Return True

                    Case Else
                        message = "지원하지 않는 파라미터 형식입니다."
                        Return False
                End Select
            Catch ex As Exception
                message = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function FindParameterOnElementOrType(doc As Document, element As Element, parameterName As String) As Parameter
            If doc Is Nothing OrElse element Is Nothing OrElse String.IsNullOrWhiteSpace(parameterName) Then Return Nothing

            Dim instanceParam = FindParameterByName(element, parameterName)
            If instanceParam IsNot Nothing Then Return instanceParam

            Dim typeId = element.GetTypeId()
            If typeId Is Nothing OrElse typeId = ElementId.InvalidElementId Then Return Nothing

            Dim typeElement = TryCast(doc.GetElement(typeId), Element)
            Return FindParameterByName(typeElement, parameterName)
        End Function

        Private Shared Function FindParameterByName(owner As Element, parameterName As String) As Parameter
            If owner Is Nothing OrElse String.IsNullOrWhiteSpace(parameterName) Then Return Nothing

            Try
                Dim direct = owner.LookupParameter(parameterName)
                If IsMatchingParameter(direct, parameterName) Then Return direct
            Catch
            End Try

            Try
                For Each parameter As Parameter In owner.Parameters
                    If IsMatchingParameter(parameter, parameterName) Then Return parameter
                Next
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function IsMatchingParameter(parameter As Parameter, parameterName As String) As Boolean
            If parameter Is Nothing OrElse parameter.Definition Is Nothing Then Return False
            Dim actualName = If(parameter.Definition.Name, String.Empty).Trim()
            Return String.Equals(actualName, parameterName.Trim(), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetParameterValueText(doc As Document, parameter As Parameter) As String
            If parameter Is Nothing Then Return String.Empty

            Try
                Dim formatted = parameter.AsValueString()
                If Not String.IsNullOrWhiteSpace(formatted) Then Return formatted
            Catch
            End Try

            Try
                Select Case parameter.StorageType
                    Case StorageType.String
                        Return If(parameter.AsString(), String.Empty)
                    Case StorageType.Integer
                        Return parameter.AsInteger().ToString(CultureInfo.InvariantCulture)
                    Case StorageType.Double
                        Return parameter.AsDouble().ToString(CultureInfo.InvariantCulture)
                    Case StorageType.ElementId
                        Dim elementId = parameter.AsElementId()
                        If elementId Is Nothing OrElse elementId = ElementId.InvalidElementId Then Return String.Empty
                        Dim refElement = doc.GetElement(elementId)
                        If refElement IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(refElement.Name) Then
                            Return refElement.Name
                        End If
                        Return elementId.IntegerValue.ToString(CultureInfo.InvariantCulture)
                    Case Else
                        Return String.Empty
                End Select
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function GetEnabledAssignments(settings As ElementParameterUpdateSettings) As List(Of ElementParameterAssignment)
            If settings Is Nothing OrElse settings.Assignments Is Nothing Then Return New List(Of ElementParameterAssignment)()
            Return settings.Assignments _
                .Where(Function(x) x IsNot Nothing AndAlso x.IsConfigured() AndAlso (x.Enabled OrElse Not String.IsNullOrWhiteSpace(x.ParameterName))) _
                .Select(Function(x) x.Clone()) _
                .ToList()
        End Function

        Private Shared Function GetEnabledConditions(settings As ElementParameterUpdateSettings) As List(Of ElementParameterCondition)
            If settings Is Nothing OrElse settings.Conditions Is Nothing Then Return New List(Of ElementParameterCondition)()
            Return settings.Conditions _
                .Where(Function(x) x IsNot Nothing AndAlso x.IsConfigured() AndAlso (x.Enabled OrElse Not String.IsNullOrWhiteSpace(x.ParameterName))) _
                .Select(Function(x) x.Clone()) _
                .ToList()
        End Function

        Private Shared Function GetAssignmentParameterNames(settings As Settings) As List(Of String)
            If settings?.ElementParameterUpdate?.Assignments Is Nothing Then Return New List(Of String)()
            Return settings.ElementParameterUpdate.Assignments _
                .Where(Function(x) x IsNot Nothing AndAlso x.IsConfigured()) _
                .Select(Function(x) x.ParameterName.Trim()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
        End Function

        Private Shared Function ValidateSettings(uiapp As UIApplication, settings As Settings) As String
            If settings Is Nothing Then Return "설정을 읽지 못했습니다."
            If settings.ElementParameterUpdate Is Nothing Then Return "파라미터 입력 설정이 없습니다."

            Dim assignments = GetEnabledAssignments(settings.ElementParameterUpdate)
            If assignments.Count = 0 Then
                Return "입력할 파라미터를 하나 이상 지정하세요."
            End If

            If settings.UseActiveDocument Then
                Dim doc = uiapp.ActiveUIDocument?.Document
                If doc Is Nothing Then Return "활성 문서를 찾을 수 없습니다."
                If doc.IsFamilyDocument Then Return "패밀리 문서에는 적용할 수 없습니다."
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
                    Return Path.Combine(Path.GetDirectoryName(doc.PathName), "KKY_ParameterModifier")
                End If
            ElseIf settings.RvtPaths IsNot Nothing Then
                Dim firstPath = settings.RvtPaths.FirstOrDefault(Function(x) Not String.IsNullOrWhiteSpace(x))
                If Not String.IsNullOrWhiteSpace(firstPath) AndAlso File.Exists(firstPath) Then
                    Return Path.Combine(Path.GetDirectoryName(firstPath), "KKY_ParameterModifier")
                End If
            End If

            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "KKY_Tool_Revit", "ParameterModifier")
        End Function

        Private Shared Function BuildSummary(fileResults As IEnumerable(Of FileRunResult)) As RunSummary
            Dim results = If(fileResults, Enumerable.Empty(Of FileRunResult)()).Where(Function(x) x IsNot Nothing).ToList()
            Return New RunSummary With {
                .SuccessCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "Success", StringComparison.OrdinalIgnoreCase)),
                .FailCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "Fail", StringComparison.OrdinalIgnoreCase)),
                .NoChangeCount = Enumerable.Count(results, Function(x) String.Equals(x.Status, "NoChange", StringComparison.OrdinalIgnoreCase)),
                .TotalCandidateCount = results.Sum(Function(x) x.CandidateCount),
                .TotalFilteredCount = results.Sum(Function(x) x.FilteredCount),
                .TotalMatchedCount = results.Sum(Function(x) x.MatchedCount),
                .TotalUpdatedElementCount = results.Sum(Function(x) x.UpdatedElementCount),
                .TotalUpdatedParameterCount = results.Sum(Function(x) x.UpdatedParameterCount)
            }
        End Function

        Private Shared Sub SaveArtifacts(result As RunResult, assignmentNames As List(Of String))
            If result Is Nothing Then Return
            Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim workbookPath = Path.Combine(result.OutputFolder, "ParameterModifierResult_" & timeStamp & ".xlsx")
            Dim logPath = Path.Combine(result.OutputFolder, "ParameterModifierLog_" & timeStamp & ".txt")

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable)) From {
                New KeyValuePair(Of String, DataTable)("Summary", BuildSummaryTable(result.Files)),
                New KeyValuePair(Of String, DataTable)("Detail", BuildDetailTable(result.Details, assignmentNames)),
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
            table.Columns.Add("UpdatedElementCount", GetType(Integer))
            table.Columns.Add("UpdatedParameterCount", GetType(Integer))
            table.Columns.Add("SyncPerformed", GetType(String))
            table.Columns.Add("CentralToLocal", GetType(String))
            table.Columns.Add("Message", GetType(String))

            For Each item In If(fileResults, Enumerable.Empty(Of FileRunResult)())
                If item Is Nothing Then Continue For
                table.Rows.Add(item.FileName, item.Status, item.CandidateCount, item.FilteredCount, item.MatchedCount, item.UpdatedElementCount, item.UpdatedParameterCount, If(item.SynchronizePerformed, "Y", "N"), If(item.UsedLocalFile, "Y", "N"), item.Message)
            Next

            Return table
        End Function

        Private Shared Function BuildDetailTable(detailRows As IEnumerable(Of DetailRow), assignmentNames As IEnumerable(Of String)) As DataTable
            Dim table As New DataTable("Detail")
            table.Columns.Add("FileName", GetType(String))
            table.Columns.Add("ElementId", GetType(Integer))
            table.Columns.Add("Category", GetType(String))
            table.Columns.Add("FamilyName", GetType(String))
            table.Columns.Add("TypeName", GetType(String))
            table.Columns.Add("Status", GetType(String))
            table.Columns.Add("Note", GetType(String))

            Dim paramNames = If(assignmentNames, Enumerable.Empty(Of String)()).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            For Each name In paramNames
                table.Columns.Add(name, GetType(String))
            Next

            For Each rowItem In If(detailRows, Enumerable.Empty(Of DetailRow)())
                If rowItem Is Nothing Then Continue For
                Dim row = table.NewRow()
                row("FileName") = rowItem.FileName
                row("ElementId") = rowItem.ElementId
                row("Category") = rowItem.Category
                row("FamilyName") = rowItem.FamilyName
                row("TypeName") = rowItem.TypeName
                row("Status") = rowItem.Status
                row("Note") = rowItem.Note
                For Each name In paramNames
                    row(name) = If(rowItem.ParameterValues IsNot Nothing AndAlso rowItem.ParameterValues.ContainsKey(name), rowItem.ParameterValues(name), String.Empty)
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
            fileResult.UpdatedElementCount = docResult.UpdatedElementCount
            fileResult.UpdatedParameterCount = docResult.UpdatedParameterCount
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

        Private Shared Function AppendMessage(currentMessage As String, nextMessage As String) As String
            If String.IsNullOrWhiteSpace(currentMessage) Then Return If(nextMessage, String.Empty)
            If String.IsNullOrWhiteSpace(nextMessage) Then Return currentMessage
            Return currentMessage & " / " & nextMessage
        End Function

        Private Shared Sub ReportProgress(progress As IProgress(Of Object),
                                          current As Integer,
                                          total As Integer,
                                          message As String,
                                          detail As String)
            If progress Is Nothing Then Return

            Dim percent As Double = 0
            If total > 0 Then
                percent = Math.Max(0, Math.Min(100, (CDbl(current) / CDbl(total)) * 100.0R))
            End If

            progress.Report(New With {
                .title = "파라미터 수정기",
                .message = If(message, String.Empty),
                .detail = If(detail, String.Empty),
                .current = current,
                .total = total,
                .percent = percent
            })
        End Sub

        Private Shared Function TryParseInteger(text As String, ByRef value As Integer) As Boolean
            Dim raw = If(text, String.Empty).Trim()
            If Integer.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, value) Then Return True
            If Integer.TryParse(raw, value) Then Return True

            Select Case raw.ToLowerInvariant()
                Case "true", "yes", "y", "on"
                    value = 1
                    Return True
                Case "false", "no", "n", "off"
                    value = 0
                    Return True
            End Select

            Return False
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
                                                    closeAllWorksets As Boolean) As Document
            Dim modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)
            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If closeAllWorksets Then
                Dim worksetConfig As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                openOpts.SetOpenWorksetsConfiguration(worksetConfig)
            End If

            Return app.OpenDocumentFile(modelPath, openOpts)
        End Function

        Private Shared Function SyncWithCentral(doc As Document,
                                                comment As String,
                                                ByRef err As String) As Boolean
            err = String.Empty
            If doc Is Nothing OrElse Not doc.IsWorkshared Then
                err = "Workshared 문서가 아닙니다."
                Return False
            End If

            Try
                Dim twc As New TransactWithCentralOptions()
                Dim swc As New SynchronizeWithCentralOptions()
                swc.Comment = If(comment, String.Empty)
                Try
                    Dim relinquish As New RelinquishOptions(True)
                    swc.SetRelinquishOptions(relinquish)
                Catch
                End Try

                doc.SynchronizeWithCentral(twc, swc)
                Return True
            Catch ex As Exception
                err = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function CreateNewLocalPath(centralPath As String) As String
            Dim localRoot = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "ParameterModifier", DateTime.Now.ToString("yyyyMMdd"))
            Directory.CreateDirectory(localRoot)

            Dim fileName = Path.GetFileNameWithoutExtension(centralPath) & "_" & Environment.UserName & "_" & DateTime.Now.ToString("HHmmssfff") & ".rvt"
            Dim localPath = Path.Combine(localRoot, fileName)

            Dim sourcePath = ModelPathUtils.ConvertUserVisiblePathToModelPath(centralPath)
            Dim targetPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath)
            WorksharingUtils.CreateNewLocal(sourcePath, targetPath)
            Return localPath
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
