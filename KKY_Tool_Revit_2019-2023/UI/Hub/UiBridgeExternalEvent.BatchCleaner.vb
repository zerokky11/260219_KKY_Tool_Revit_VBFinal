Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Models
Imports KKY_Tool_Revit.Services
Imports Microsoft.VisualBasic.FileIO
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared ReadOnly _deliveryCleanerLock As New Object()
        Private Shared ReadOnly _deliveryCleanerLogs As New List(Of String)()
        Private Shared _deliveryCleanerSettings As BatchCleanSettings
        Private Shared _deliveryCleanerSession As BatchPrepareSession
        Private Shared _deliveryCleanerExtractParamsCsv As String = String.Empty
        Private Shared _deliveryCleanerLastLogExportPath As String = String.Empty

        Private Sub HandleDeliveryCleanerInit(app As UIApplication, payload As Object)
            SendToWeb("deliverycleaner:init", BuildDeliveryCleanerStatePayload())
        End Sub

        Private Sub HandleDeliveryCleanerPickRvts(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                SendToWeb("deliverycleaner:rvts-picked", New With {.ok = True, .paths = dlg.FileNames})
            End Using
        End Sub

        Private Sub HandleDeliveryCleanerBrowseOutputFolder(app As UIApplication, payload As Object)
            Using dlg As New WinForms.FolderBrowserDialog()
                dlg.Description = "정리 결과 폴더 선택"
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                SendToWeb("deliverycleaner:output-folder-picked", New With {.ok = True, .path = dlg.SelectedPath})
            End Using
        End Sub

        Private Sub HandleDeliveryCleanerFilterImport(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "XML (*.xml)|*.xml"
                dlg.Title = "View Filter XML 불러오기"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                Dim profile = RevitViewFilterProfileService.LoadFromXml(dlg.FileName)
                SendToWeb("deliverycleaner:filter-loaded", New With {
                    .ok = True,
                    .profile = SerializeFilterProfile(profile),
                    .source = dlg.FileName
                })
            End Using
        End Sub

        Private Sub HandleDeliveryCleanerFilterSave(app As UIApplication, payload As Object)
            Dim profile = ParseDeliveryCleanerFilterProfile(GetProp(ParsePayloadDict(payload), "filterProfile"))
            If profile Is Nothing OrElse Not profile.IsConfigured() Then
                SendToWeb("deliverycleaner:error", New With {.message = "저장할 필터 설정이 올바르지 않습니다."})
                Return
            End If

            Using dlg As New WinForms.SaveFileDialog()
                dlg.Filter = "XML (*.xml)|*.xml"
                dlg.Title = "View Filter XML 저장"
                dlg.FileName = If(String.IsNullOrWhiteSpace(profile.FilterName), "ViewFilterProfile.xml", profile.FilterName & ".xml")
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                RevitViewFilterProfileService.SaveToXml(profile, dlg.FileName)
                SendToWeb("deliverycleaner:filter-saved", New With {.ok = True, .path = dlg.FileName})
            End Using
        End Sub

        Private Sub HandleDeliveryCleanerFilterDocList(app As UIApplication, payload As Object)
            Dim doc = app.ActiveUIDocument?.Document
            If doc Is Nothing Then
                SendToWeb("deliverycleaner:filter-doc-list", New With {
                    .ok = False,
                    .message = "현재 열려 있는 문서를 찾을 수 없습니다.",
                    .items = New List(Of Object)()
                })
                Return
            End If

            Dim items = New FilteredElementCollector(doc) _
                .OfClass(GetType(ParameterFilterElement)) _
                .Cast(Of ParameterFilterElement)() _
                .OrderBy(Function(x) x.Name) _
                .Select(Function(x) New With {.id = x.Id.IntegerValue, .name = x.Name}) _
                .ToList()

            SendToWeb("deliverycleaner:filter-doc-list", New With {.ok = True, .items = items, .docTitle = doc.Title})
        End Sub

        Private Sub HandleDeliveryCleanerFilterDocExtract(app As UIApplication, payload As Object)
            Dim doc = app.ActiveUIDocument?.Document
            If doc Is Nothing Then
                SendToWeb("deliverycleaner:error", New With {.message = "현재 열려 있는 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim pd = ParsePayloadDict(payload)
            Dim filterIdInt = SafeIntObj(GetProp(pd, "filterId"), Integer.MinValue)
            If filterIdInt = Integer.MinValue Then
                SendToWeb("deliverycleaner:error", New With {.message = "추출할 필터를 선택하세요."})
                Return
            End If

            Dim filterEl = TryCast(doc.GetElement(New ElementId(filterIdInt)), ParameterFilterElement)
            If filterEl Is Nothing Then
                SendToWeb("deliverycleaner:error", New With {.message = "선택한 필터를 찾을 수 없습니다."})
                Return
            End If

            Dim profile = RevitViewFilterProfileService.ExtractProfileFromFilter(doc, filterEl.Id)
            SendToWeb("deliverycleaner:filter-loaded", New With {
                .ok = True,
                .profile = SerializeFilterProfile(profile),
                .source = doc.Title
            })
        End Sub

        Private Sub HandleDeliveryCleanerRun(app As UIApplication, payload As Object)
            Dim settings = ParseDeliveryCleanerSettings(payload)
            Dim validationMessage = ValidateDeliveryCleanerSettings(settings)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                SendToWeb("deliverycleaner:error", New With {.message = validationMessage})
                Return
            End If

            SyncLock _deliveryCleanerLock
                _deliveryCleanerLogs.Clear()
            End SyncLock

            Dim logger As Action(Of String) = Sub(line) AppendDeliveryCleanerLog(line)
            AppendDeliveryCleanerLog("정리 시작")

            Try
                Dim session = BatchCleanService.CleanAndSave(app, settings, logger)
                session.DesignOptionAuditCsvPath = ConvertDeliveryCleanerCsvToXlsx(session.DesignOptionAuditCsvPath, "Design Option 검토")
                SyncLock _deliveryCleanerLock
                    _deliveryCleanerSettings = settings.Clone()
                    _deliveryCleanerSession = session
                End SyncLock

                SendToWeb("deliverycleaner:run-done", New With {
                    .ok = True,
                    .state = BuildDeliveryCleanerStatePayload(),
                    .summary = BuildDeliveryCleanerRunSummary(session)
                })
            Catch ex As Exception
                AppendDeliveryCleanerLog("정리 중 오류: " & ex.Message)
                SendToWeb("deliverycleaner:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleDeliveryCleanerVerify(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim settings = ParseDeliveryCleanerSettings(payload)
            Dim targetPaths = ResolveDeliveryCleanerTargetPaths(ParseStringList(pd, "filePaths"))
            If targetPaths.Count = 0 Then
                SendToWeb("deliverycleaner:error", New With {.message = "검토할 RVT 파일이 없습니다."})
                Return
            End If

            Dim outputFolder = ResolveDeliveryCleanerOutputFolder(settings, targetPaths)
            Dim logger As Action(Of String) = Sub(line) AppendDeliveryCleanerLog(line)

            Try
                Dim csvPath = VerificationService.VerifyPaths(app, targetPaths, outputFolder, settings, logger)
                Dim exportPath = ConvertDeliveryCleanerCsvToXlsx(csvPath, "정리 결과 검토")
                SyncLock _deliveryCleanerLock
                    If _deliveryCleanerSession Is Nothing Then _deliveryCleanerSession = New BatchPrepareSession()
                    _deliveryCleanerSession.VerificationCsvPath = exportPath
                    If String.IsNullOrWhiteSpace(_deliveryCleanerSession.OutputFolder) Then _deliveryCleanerSession.OutputFolder = outputFolder
                    If _deliveryCleanerSettings Is Nothing Then _deliveryCleanerSettings = settings.Clone()
                End SyncLock

                SendToWeb("deliverycleaner:verify-done", New With {.ok = True, .path = exportPath, .state = BuildDeliveryCleanerStatePayload()})
            Catch ex As Exception
                AppendDeliveryCleanerLog("검토 중 오류: " & ex.Message)
                SendToWeb("deliverycleaner:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleDeliveryCleanerExtract(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim settings = ParseDeliveryCleanerSettings(payload)
            Dim targetPaths = ResolveDeliveryCleanerTargetPaths(ParseStringList(pd, "filePaths"))
            If targetPaths.Count = 0 Then
                SendToWeb("deliverycleaner:error", New With {.message = "속성값을 추출할 RVT 파일이 없습니다."})
                Return
            End If

            Dim parameterNamesCsv = Convert.ToString(GetProp(pd, "extractParameterNamesCsv"))
            If String.IsNullOrWhiteSpace(parameterNamesCsv) Then parameterNamesCsv = _deliveryCleanerExtractParamsCsv
            If String.IsNullOrWhiteSpace(parameterNamesCsv) Then
                SendToWeb("deliverycleaner:error", New With {.message = "추출할 파라미터 이름을 입력하세요."})
                Return
            End If

            Dim outputFolder = ResolveDeliveryCleanerOutputFolder(settings, targetPaths)
            Dim logger As Action(Of String) = Sub(line) AppendDeliveryCleanerLog(line)

            Try
                Dim csvPath = ModelParameterExtractionService.ExportModelParameters(app, targetPaths, outputFolder, parameterNamesCsv, logger)
                Dim exportPath = ConvertDeliveryCleanerCsvToXlsx(csvPath, "속성값 추출")
                SyncLock _deliveryCleanerLock
                    _deliveryCleanerExtractParamsCsv = parameterNamesCsv
                    If _deliveryCleanerSession Is Nothing Then _deliveryCleanerSession = New BatchPrepareSession()
                    If String.IsNullOrWhiteSpace(_deliveryCleanerSession.OutputFolder) Then _deliveryCleanerSession.OutputFolder = outputFolder
                    If _deliveryCleanerSettings Is Nothing Then _deliveryCleanerSettings = settings.Clone()
                End SyncLock

                SendToWeb("deliverycleaner:extract-done", New With {
                    .ok = True,
                    .path = exportPath,
                    .parameterNamesCsv = parameterNamesCsv,
                    .state = BuildDeliveryCleanerStatePayload()
                })
            Catch ex As Exception
                AppendDeliveryCleanerLog("속성값 추출 중 오류: " & ex.Message)
                SendToWeb("deliverycleaner:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleDeliveryCleanerPurge(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim settings = ParseDeliveryCleanerSettings(payload)
            Dim targetPaths = ResolveDeliveryCleanerTargetPaths(ParseStringList(pd, "filePaths"))
            If targetPaths.Count = 0 Then
                SendToWeb("deliverycleaner:error", New With {.message = "Purge 대상 파일이 없습니다."})
                Return
            End If

            If targetPaths.Any(Function(path) String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path)) Then
                SendToWeb("deliverycleaner:error", New With {.message = "Purge 대상 파일 경로를 다시 확인하세요."})
                Return
            End If

            If PurgeUiBatchService.IsRunning Then
                SendToWeb("deliverycleaner:error", New With {.message = "이미 Purge 일괄처리가 실행 중입니다."})
                Return
            End If

            Dim session = EnsureDeliveryCleanerSession(targetPaths, settings)
            Dim logger As Action(Of String) = Sub(line) AppendDeliveryCleanerLog(line)

            Try
                Dim started = PurgeUiBatchService.Start(app, session, 5, logger)
                If Not started Then
                    SendToWeb("deliverycleaner:error", New With {.message = "Purge를 시작하지 못했습니다."})
                    Return
                End If

                SendToWeb("deliverycleaner:purge-started", New With {
                    .ok = True,
                    .state = BuildDeliveryCleanerStatePayload(),
                    .snapshot = SerializePurgeSnapshot(PurgeUiBatchService.GetProgressSnapshot())
                })
            Catch ex As Exception
                AppendDeliveryCleanerLog("Purge 시작 실패: " & ex.Message)
                SendToWeb("deliverycleaner:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleDeliveryCleanerPurgeStatus(app As UIApplication, payload As Object)
            SendToWeb("deliverycleaner:purge-status", New With {.ok = True, .snapshot = SerializePurgeSnapshot(PurgeUiBatchService.GetProgressSnapshot())})
        End Sub

        Private Sub HandleDeliveryCleanerExportLog(app As UIApplication, payload As Object)
            Dim logs As List(Of String) = Nothing
            Dim preferredFolder As String = String.Empty

            SyncLock _deliveryCleanerLock
                logs = New List(Of String)(_deliveryCleanerLogs)
                If _deliveryCleanerSession IsNot Nothing Then preferredFolder = _deliveryCleanerSession.OutputFolder
                If String.IsNullOrWhiteSpace(preferredFolder) AndAlso _deliveryCleanerSettings IsNot Nothing Then preferredFolder = _deliveryCleanerSettings.OutputFolder
            End SyncLock

            If logs.Count = 0 Then
                SendToWeb("deliverycleaner:error", New With {.message = "저장할 실행 로그가 없습니다."})
                Return
            End If

            Dim pd = ParsePayloadDict(payload)
            Dim outputFolder = NormalizeDeliveryCleanerText(GetProp(pd, "outputFolder"))
            If String.IsNullOrWhiteSpace(outputFolder) Then outputFolder = preferredFolder
            If String.IsNullOrWhiteSpace(outputFolder) Then
                outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "KKY Tool Revit", "DeliveryCleanerLogs")
            End If

            Try
                Directory.CreateDirectory(outputFolder)
                Dim exportPath = Path.Combine(outputFolder, $"RVT_정리_납품용_로그_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
                SaveDeliveryCleanerLogsWorkbook(exportPath, logs)

                SyncLock _deliveryCleanerLock
                    _deliveryCleanerLastLogExportPath = exportPath
                End SyncLock

                SendToWeb("deliverycleaner:log-exported", New With {
                    .ok = True,
                    .path = exportPath,
                    .state = BuildDeliveryCleanerStatePayload()
                })
            Catch ex As Exception
                SendToWeb("deliverycleaner:error", New With {.message = "로그 엑셀 저장 중 오류가 발생했습니다: " & ex.Message})
            End Try
        End Sub

        Private Sub HandleDeliveryCleanerOpenFolder(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim pathText = Convert.ToString(GetProp(pd, "path"))
            If String.IsNullOrWhiteSpace(pathText) Then
                SyncLock _deliveryCleanerLock
                    If _deliveryCleanerSession IsNot Nothing Then pathText = _deliveryCleanerSession.OutputFolder
                    If String.IsNullOrWhiteSpace(pathText) AndAlso _deliveryCleanerSettings IsNot Nothing Then pathText = _deliveryCleanerSettings.OutputFolder
                End SyncLock
            End If

            If String.IsNullOrWhiteSpace(pathText) Then
                SendToWeb("deliverycleaner:error", New With {.message = "열 폴더 경로가 없습니다."})
                Return
            End If

            Dim targetPath = pathText
            If File.Exists(targetPath) Then targetPath = Path.GetDirectoryName(targetPath)
            If String.IsNullOrWhiteSpace(targetPath) OrElse Not Directory.Exists(targetPath) Then
                SendToWeb("deliverycleaner:error", New With {.message = "폴더 경로를 찾을 수 없습니다."})
                Return
            End If

            Dim psi As New ProcessStartInfo("explorer.exe", """" & targetPath & """")
            psi.UseShellExecute = True
            Process.Start(psi)
            SendToWeb("deliverycleaner:folder-opened", New With {.ok = True, .path = targetPath})
        End Sub

        Private Shared Function BuildDeliveryCleanerStatePayload() As Object
            Dim settings As BatchCleanSettings = Nothing
            Dim session As BatchPrepareSession = Nothing
            Dim logs As List(Of String) = Nothing
            Dim extractCsv As String = String.Empty
            Dim lastLogExportPath As String = String.Empty

            SyncLock _deliveryCleanerLock
                settings = If(_deliveryCleanerSettings IsNot Nothing, _deliveryCleanerSettings.Clone(), Nothing)
                session = _deliveryCleanerSession
                logs = New List(Of String)(_deliveryCleanerLogs)
                extractCsv = _deliveryCleanerExtractParamsCsv
                lastLogExportPath = _deliveryCleanerLastLogExportPath
            End SyncLock

            Return New With {
                .settings = SerializeDeliveryCleanerSettings(settings),
                .session = SerializeDeliveryCleanerSession(session),
                .logs = logs,
                .extractParameterNamesCsv = extractCsv,
                .lastLogExportPath = lastLogExportPath,
                .purge = SerializePurgeSnapshot(PurgeUiBatchService.GetProgressSnapshot())
            }
        End Function

        Private Shared Function SerializeDeliveryCleanerSettings(settings As BatchCleanSettings) As Object
            If settings Is Nothing Then Return Nothing

            Return New With {
                .filePaths = If(settings.FilePaths, New List(Of String)()),
                .outputFolder = settings.OutputFolder,
                .target3DViewName = settings.Target3DViewName,
                .viewParameters = If(settings.ViewParameters, New List(Of ViewParameterAssignment)()) _
                    .Select(Function(x) New With {
                        .enabled = If(x IsNot Nothing, x.Enabled, False),
                        .parameterName = If(x IsNot Nothing, x.ParameterName, String.Empty),
                        .parameterValue = If(x IsNot Nothing, x.ParameterValue, String.Empty)
                    }).ToList(),
                .useFilter = settings.UseFilter,
                .applyFilterInitially = settings.ApplyFilterInitially,
                .autoEnableFilterIfEmpty = settings.AutoEnableFilterIfEmpty,
                .filterProfile = SerializeFilterProfile(settings.FilterProfile),
                .elementParameterUpdate = SerializeElementParameterUpdateSettings(settings.ElementParameterUpdate)
            }
        End Function

        Private Shared Function SerializeDeliveryCleanerSession(session As BatchPrepareSession) As Object
            If session Is Nothing Then Return Nothing

            Return New With {
                .outputFolder = session.OutputFolder,
                .cleanedOutputPaths = If(session.CleanedOutputPaths, New List(Of String)()),
                .verificationCsvPath = session.VerificationCsvPath,
                .designOptionAuditCsvPath = session.DesignOptionAuditCsvPath,
                .results = If(session.Results, New List(Of BatchCleanResult)()) _
                    .Select(Function(x) New With {
                        .sourcePath = x.SourcePath,
                        .outputPath = x.OutputPath,
                        .success = x.Success,
                        .message = x.Message
                    }).ToList()
            }
        End Function

        Private Shared Function BuildDeliveryCleanerRunSummary(session As BatchPrepareSession) As Object
            Dim results = If(session?.Results, New List(Of BatchCleanResult)())
            Dim successCount = results.Where(Function(x) x IsNot Nothing AndAlso x.Success).Count()
            Dim failCount = results.Count - successCount
            Dim cleanedCount = If(session?.CleanedOutputPaths, New List(Of String)()).Count

            Return New With {
                .successCount = successCount,
                .failCount = failCount,
                .cleanedCount = cleanedCount
            }
        End Function

        Private Shared Function SerializeFilterProfile(profile As ViewFilterProfile) As Object
            If profile Is Nothing Then Return Nothing

            Return New With {
                .filterName = profile.FilterName,
                .categoriesCsv = profile.CategoriesCsv,
                .parameterToken = profile.ParameterToken,
                .operatorName = profile.Operator.ToString(),
                .ruleValue = profile.RuleValue,
                .filterDefinitionXml = profile.FilterDefinitionXml,
                .structureSummary = profile.StructureSummary
            }
        End Function

        Private Shared Function SerializeElementParameterUpdateSettings(settings As ElementParameterUpdateSettings) As Object
            If settings Is Nothing Then Return Nothing

            Return New With {
                .enabled = settings.Enabled,
                .combinationMode = settings.CombinationMode.ToString(),
                .summary = settings.BuildSummary(),
                .conditions = If(settings.Conditions, New List(Of ElementParameterCondition)()) _
                    .Select(Function(x) New With {
                        .enabled = If(x IsNot Nothing, x.Enabled, False),
                        .parameterName = If(x IsNot Nothing, x.ParameterName, String.Empty),
                        .operatorName = If(x IsNot Nothing, x.Operator.ToString(), FilterRuleOperator.Equals.ToString()),
                        .value = If(x IsNot Nothing, x.Value, String.Empty)
                    }).ToList(),
                .assignments = If(settings.Assignments, New List(Of ElementParameterAssignment)()) _
                    .Select(Function(x) New With {
                        .enabled = If(x IsNot Nothing, x.Enabled, False),
                        .parameterName = If(x IsNot Nothing, x.ParameterName, String.Empty),
                        .value = If(x IsNot Nothing, x.Value, String.Empty)
                    }).ToList()
            }
        End Function

        Private Shared Function SerializePurgeSnapshot(snapshot As PurgeBatchProgressSnapshot) As Object
            If snapshot Is Nothing Then
                Return New With {
                    .isRunning = False,
                    .isCompleted = False,
                    .isFaulted = False,
                    .currentFileIndex = 0,
                    .totalFiles = 0,
                    .currentIteration = 0,
                    .totalIterations = 0,
                    .currentFileName = "",
                    .stateName = "",
                    .message = ""
                }
            End If

            Return New With {
                .isRunning = snapshot.IsRunning,
                .isCompleted = snapshot.IsCompleted,
                .isFaulted = snapshot.IsFaulted,
                .currentFileIndex = snapshot.CurrentFileIndex,
                .totalFiles = snapshot.TotalFiles,
                .currentIteration = snapshot.CurrentIteration,
                .totalIterations = snapshot.TotalIterations,
                .currentFileName = snapshot.CurrentFileName,
                .stateName = snapshot.StateName,
                .message = snapshot.Message
            }
        End Function

        Private Shared Function ParseDeliveryCleanerSettings(payload As Object) As BatchCleanSettings
            Dim pd = ParsePayloadDict(payload)
            Dim settings As New BatchCleanSettings()

            settings.FilePaths = ParseStringList(pd, "filePaths")
            settings.OutputFolder = NormalizeDeliveryCleanerText(GetProp(pd, "outputFolder"))
            settings.Target3DViewName = NormalizeDeliveryCleanerText(GetProp(pd, "target3DViewName"))
            settings.UseFilter = SafeBoolObj(GetProp(pd, "useFilter"), False)
            settings.ApplyFilterInitially = SafeBoolObj(GetProp(pd, "applyFilterInitially"), True)
            settings.AutoEnableFilterIfEmpty = SafeBoolObj(GetProp(pd, "autoEnableFilterIfEmpty"), False)
            settings.ViewParameters = ParseDeliveryCleanerViewParameters(GetProp(pd, "viewParameters"))
            settings.FilterProfile = ParseDeliveryCleanerFilterProfile(GetProp(pd, "filterProfile"))
            settings.ElementParameterUpdate = ParseDeliveryCleanerElementUpdate(GetProp(pd, "elementParameterUpdate"))

            Return settings
        End Function

        Private Shared Function ParseDeliveryCleanerViewParameters(raw As Object) As List(Of ViewParameterAssignment)
            Dim list As New List(Of ViewParameterAssignment)()
            Dim enumerable = TryCast(raw, IEnumerable)
            If enumerable Is Nothing OrElse TypeOf raw Is String Then Return list

            For Each item As Object In enumerable
                Dim d = ParsePayloadDict(item)
                list.Add(New ViewParameterAssignment With {
                    .Enabled = SafeBoolObj(GetProp(d, "enabled"), False),
                    .ParameterName = NormalizeDeliveryCleanerText(GetProp(d, "parameterName")),
                    .ParameterValue = NormalizeDeliveryCleanerText(GetProp(d, "parameterValue"))
                })
            Next

            Return list
        End Function

        Private Shared Function ParseDeliveryCleanerFilterProfile(raw As Object) As ViewFilterProfile
            If raw Is Nothing Then Return Nothing

            Dim d = ParsePayloadDict(raw)
            Dim profile As New ViewFilterProfile()
            profile.FilterName = NormalizeDeliveryCleanerText(GetProp(d, "filterName"))
            profile.CategoriesCsv = NormalizeDeliveryCleanerText(GetProp(d, "categoriesCsv"))
            profile.ParameterToken = NormalizeDeliveryCleanerText(GetProp(d, "parameterToken"))
            profile.RuleValue = Convert.ToString(GetProp(d, "ruleValue"))
            profile.FilterDefinitionXml = NormalizeDeliveryCleanerText(GetProp(d, "filterDefinitionXml"))
            profile.StructureSummary = NormalizeDeliveryCleanerText(GetProp(d, "structureSummary"))

            Dim operatorName = NormalizeDeliveryCleanerText(GetProp(d, "operatorName"))
            Dim op As FilterRuleOperator
            If [Enum].TryParse(operatorName, True, op) Then
                profile.Operator = op
            Else
                profile.Operator = FilterRuleOperator.Equals
            End If

            Return profile
        End Function

        Private Shared Function ParseDeliveryCleanerElementUpdate(raw As Object) As ElementParameterUpdateSettings
            Dim result As New ElementParameterUpdateSettings()
            If raw Is Nothing Then Return result

            Dim d = ParsePayloadDict(raw)
            result.Enabled = SafeBoolObj(GetProp(d, "enabled"), False)

            Dim comboName = NormalizeDeliveryCleanerText(GetProp(d, "combinationMode"))
            Dim combo As ParameterConditionCombination
            If [Enum].TryParse(comboName, True, combo) Then
                result.CombinationMode = combo
            End If

            Dim conditionsRaw = TryCast(GetProp(d, "conditions"), IEnumerable)
            If conditionsRaw IsNot Nothing Then
                For Each item As Object In conditionsRaw
                    Dim itemDict = ParsePayloadDict(item)
                    Dim operatorName = NormalizeDeliveryCleanerText(GetProp(itemDict, "operatorName"))
                    Dim op As FilterRuleOperator
                    If Not [Enum].TryParse(operatorName, True, op) Then op = FilterRuleOperator.Equals

                    result.Conditions.Add(New ElementParameterCondition With {
                        .Enabled = SafeBoolObj(GetProp(itemDict, "enabled"), False),
                        .ParameterName = NormalizeDeliveryCleanerText(GetProp(itemDict, "parameterName")),
                        .Operator = op,
                        .Value = Convert.ToString(GetProp(itemDict, "value"))
                    })
                Next
            End If

            Dim assignmentsRaw = TryCast(GetProp(d, "assignments"), IEnumerable)
            If assignmentsRaw IsNot Nothing Then
                For Each item As Object In assignmentsRaw
                    Dim itemDict = ParsePayloadDict(item)
                    result.Assignments.Add(New ElementParameterAssignment With {
                        .Enabled = SafeBoolObj(GetProp(itemDict, "enabled"), False),
                        .ParameterName = NormalizeDeliveryCleanerText(GetProp(itemDict, "parameterName")),
                        .Value = Convert.ToString(GetProp(itemDict, "value"))
                    })
                Next
            End If

            Return result
        End Function

        Private Shared Function ValidateDeliveryCleanerSettings(settings As BatchCleanSettings) As String
            If settings Is Nothing Then Return "정리 설정을 읽지 못했습니다."
            If settings.FilePaths Is Nothing OrElse settings.FilePaths.Count = 0 Then Return "RVT 파일을 하나 이상 추가하세요."
            If String.IsNullOrWhiteSpace(settings.OutputFolder) Then Return "정리 결과 폴더를 지정하세요."

            Try
                Directory.CreateDirectory(settings.OutputFolder)
            Catch ex As Exception
                Return "정리 결과 폴더를 만들 수 없습니다: " & ex.Message
            End Try

            If settings.UseFilter AndAlso (settings.FilterProfile Is Nothing OrElse Not settings.FilterProfile.IsConfigured()) Then
                Return "필터 사용이 켜져 있으면 XML 가져오기 또는 문서 추출로 필터를 준비해야 합니다."
            End If

            If settings.ElementParameterUpdate IsNot Nothing AndAlso settings.ElementParameterUpdate.Enabled AndAlso Not settings.ElementParameterUpdate.IsConfigured() Then
                Return "객체 파라미터 입력을 사용하려면 조건과 입력 파라미터를 함께 지정해야 합니다."
            End If

            Return String.Empty
        End Function

        Private Shared Function ResolveDeliveryCleanerTargetPaths(payloadPaths As List(Of String)) As List(Of String)
            Dim normalized = If(payloadPaths, New List(Of String)()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x) AndAlso File.Exists(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
            If normalized.Count > 0 Then Return normalized

            SyncLock _deliveryCleanerLock
                If _deliveryCleanerSession IsNot Nothing AndAlso _deliveryCleanerSession.CleanedOutputPaths IsNot Nothing Then
                    Return _deliveryCleanerSession.CleanedOutputPaths _
                        .Where(Function(x) Not String.IsNullOrWhiteSpace(x) AndAlso File.Exists(x)) _
                        .Distinct(StringComparer.OrdinalIgnoreCase) _
                        .ToList()
                End If

                If _deliveryCleanerSettings IsNot Nothing AndAlso _deliveryCleanerSettings.FilePaths IsNot Nothing Then
                    Return _deliveryCleanerSettings.FilePaths _
                        .Where(Function(x) Not String.IsNullOrWhiteSpace(x) AndAlso File.Exists(x)) _
                        .Distinct(StringComparer.OrdinalIgnoreCase) _
                        .ToList()
                End If
            End SyncLock

            Return New List(Of String)()
        End Function

        Private Shared Function ResolveDeliveryCleanerOutputFolder(settings As BatchCleanSettings, targetPaths As IList(Of String)) As String
            Dim pathText = NormalizeDeliveryCleanerText(settings?.OutputFolder)
            If String.IsNullOrWhiteSpace(pathText) Then
                SyncLock _deliveryCleanerLock
                    If _deliveryCleanerSession IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(_deliveryCleanerSession.OutputFolder) Then
                        pathText = _deliveryCleanerSession.OutputFolder
                    ElseIf _deliveryCleanerSettings IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(_deliveryCleanerSettings.OutputFolder) Then
                        pathText = _deliveryCleanerSettings.OutputFolder
                    End If
                End SyncLock
            End If

            If String.IsNullOrWhiteSpace(pathText) AndAlso targetPaths IsNot Nothing AndAlso targetPaths.Count > 0 Then
                pathText = Path.GetDirectoryName(targetPaths(0))
            End If

            If Not String.IsNullOrWhiteSpace(pathText) Then Directory.CreateDirectory(pathText)
            Return If(pathText, String.Empty)
        End Function

        Private Shared Function EnsureDeliveryCleanerSession(targetPaths As IList(Of String), settings As BatchCleanSettings) As BatchPrepareSession
            SyncLock _deliveryCleanerLock
                Dim normalizedTargets = If(targetPaths, New List(Of String)()) _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Distinct(StringComparer.OrdinalIgnoreCase) _
                    .ToList()

                If _deliveryCleanerSession Is Nothing Then _deliveryCleanerSession = New BatchPrepareSession()
                _deliveryCleanerSession.OutputFolder = ResolveDeliveryCleanerOutputFolder(settings, normalizedTargets)
                _deliveryCleanerSession.CleanedOutputPaths = normalizedTargets
                If settings IsNot Nothing Then _deliveryCleanerSettings = settings.Clone()
                Return _deliveryCleanerSession
            End SyncLock
        End Function

        Private Shared Function NormalizeDeliveryCleanerText(value As Object) As String
            Dim text = Convert.ToString(value)
            Return If(text, String.Empty).Trim()
        End Function

        Private Shared Function ConvertDeliveryCleanerCsvToXlsx(csvPath As String, sheetName As String) As String
            If String.IsNullOrWhiteSpace(csvPath) OrElse Not File.Exists(csvPath) Then Return csvPath
            If String.Equals(Path.GetExtension(csvPath), ".xlsx", StringComparison.OrdinalIgnoreCase) Then Return csvPath

            Dim table = LoadDeliveryCleanerCsvAsDataTable(csvPath, sheetName)
            Dim xlsxPath = Path.ChangeExtension(csvPath, ".xlsx")
            ExcelCore.SaveXlsx(xlsxPath, sheetName, table, autoFit:=True)

            Try
                File.Delete(csvPath)
            Catch
            End Try

            Return xlsxPath
        End Function

        Private Shared Function LoadDeliveryCleanerCsvAsDataTable(csvPath As String, tableName As String) As DataTable
            Dim table As New DataTable(If(String.IsNullOrWhiteSpace(tableName), "Sheet1", tableName))

            Using parser As New TextFieldParser(csvPath, Encoding.UTF8)
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                parser.TrimWhiteSpace = False

                If parser.EndOfData Then
                    table.Columns.Add("Message", GetType(String))
                    Return table
                End If

                Dim headers = parser.ReadFields()
                If headers Is Nothing OrElse headers.Length = 0 Then
                    table.Columns.Add("Message", GetType(String))
                    Return table
                End If

                For i = 0 To headers.Length - 1
                    Dim name = headers(i)
                    If String.IsNullOrWhiteSpace(name) Then name = $"Column{i + 1}"
                    Dim uniqueName = name
                    Dim suffix = 2
                    While table.Columns.Contains(uniqueName)
                        uniqueName = $"{name}_{suffix}"
                        suffix += 1
                    End While
                    table.Columns.Add(uniqueName, GetType(String))
                Next

                While Not parser.EndOfData
                    Dim fields = parser.ReadFields()
                    If fields Is Nothing Then Continue While

                    Dim row = table.NewRow()
                    For i = 0 To table.Columns.Count - 1
                        row(i) = If(i < fields.Length, fields(i), String.Empty)
                    Next
                    table.Rows.Add(row)
                End While
            End Using

            Return table
        End Function

        Private Shared Sub SaveDeliveryCleanerLogsWorkbook(filePath As String, logs As IList(Of String))
            Using workbook As IWorkbook = New XSSFWorkbook()
                Dim sheet = workbook.CreateSheet("실행 로그")
                Dim headerRow = sheet.CreateRow(0)
                headerRow.CreateCell(0).SetCellValue("No")
                headerRow.CreateCell(1).SetCellValue("기록")

                For i = 0 To logs.Count - 1
                    Dim row = sheet.CreateRow(i + 1)
                    row.CreateCell(0).SetCellValue(i + 1)
                    row.CreateCell(1).SetCellValue(If(logs(i), String.Empty))
                Next

                sheet.SetColumnWidth(0, 12 * 256)
                sheet.SetColumnWidth(1, 120 * 256)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    workbook.Write(fs)
                End Using
            End Using
        End Sub

        Private Shared Sub AppendDeliveryCleanerLog(message As String)
            If String.IsNullOrWhiteSpace(message) Then Return
            Dim stamped = $"[{DateTime.Now:HH:mm:ss}] {message}"

            SyncLock _deliveryCleanerLock
                _deliveryCleanerLogs.Add(stamped)
                If _deliveryCleanerLogs.Count > 2000 Then
                    _deliveryCleanerLogs.RemoveRange(0, _deliveryCleanerLogs.Count - 2000)
                End If
            End SyncLock

            SendToWeb("deliverycleaner:log", New With {.message = stamped})
        End Sub

    End Class

End Namespace
