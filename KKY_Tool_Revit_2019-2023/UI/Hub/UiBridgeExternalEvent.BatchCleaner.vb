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
        Private Const DeliveryCleanerProgressChannel As String = "deliverycleaner:progress"

        Private NotInheritable Class DeliveryCleanerLogSummaryRow
            Public Property FileName As String
            Public Property WorkType As String
            Public Property ItemName As String
            Public Property Status As String
            Public Property Detail As String
        End Class

        Private NotInheritable Class DeliveryCleanerStepTracker
            Public Property FileName As String
            Public Property WorkType As String
            Public Property ItemName As String
            Public Property SuccessDetail As String
            Public Property FailureDetail As String
        End Class

        Private NotInheritable Class DeliveryCleanerProgressContext
            Public Property Title As String
            Public Property Mode As String
            Public Property TotalFiles As Integer
            Public Property CompletedFiles As Integer
            Public Property CurrentFileIndex As Integer
            Public Property CurrentFileName As String
            Public Property StepIndex As Integer
            Public Property TotalSteps As Integer
            Public Property LastPercent As Double
        End Class

        Private Shared Function CountExistingDeliveryCleanerFiles(paths As IEnumerable(Of String)) As Integer
            Return If(paths, Enumerable.Empty(Of String)()) _
                .Where(Function(path) Not String.IsNullOrWhiteSpace(path) AndAlso File.Exists(path)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .Count()
        End Function

        Private Shared Function ClampDeliveryCleanerPercent(value As Double) As Double
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then Return 0.0R
            Return Math.Max(0.0R, Math.Min(100.0R, value))
        End Function

        Private Shared Sub SendDeliveryCleanerProgress(context As DeliveryCleanerProgressContext,
                                                       message As String,
                                                       detail As String,
                                                       Optional fileProgress As Double? = Nothing,
                                                       Optional fixedPercent As Double? = Nothing,
                                                       Optional complete As Boolean = False)
            If context Is Nothing Then Return

            Dim totalFiles As Integer = Math.Max(1, context.TotalFiles)
            Dim percent As Double

            If complete Then
                percent = 100.0R
            ElseIf fixedPercent.HasValue Then
                percent = fixedPercent.Value
            Else
                Dim completed As Integer = Math.Max(0, Math.Min(context.CompletedFiles, totalFiles))
                Dim activeProgress As Double = 0.0R
                If fileProgress.HasValue AndAlso completed < totalFiles Then
                    activeProgress = Math.Max(0.0R, Math.Min(1.0R, fileProgress.Value))
                End If
                percent = ((CDbl(completed) + activeProgress) / CDbl(totalFiles)) * 100.0R
            End If

            percent = ClampDeliveryCleanerPercent(percent)
            If Not complete Then
                percent = Math.Max(context.LastPercent, percent)
            End If
            context.LastPercent = percent

            SendToWeb(DeliveryCleanerProgressChannel, New With {
                .title = If(context.Title, "Delivery Cleaner"),
                .mode = If(context.Mode, String.Empty),
                .message = If(message, String.Empty),
                .detail = If(detail, String.Empty),
                .percent = percent,
                .complete = complete,
                .currentFile = If(context.CurrentFileName, String.Empty),
                .currentFileIndex = Math.Max(0, context.CurrentFileIndex),
                .totalFiles = Math.Max(0, context.TotalFiles)
            })
        End Sub

        Private Shared Sub UpdateDeliveryCleanerProgressFromLog(context As DeliveryCleanerProgressContext, rawLine As String)
            If context Is Nothing Then Return

            Dim line = NormalizeDeliveryCleanerLogLine(rawLine)
            If String.IsNullOrWhiteSpace(line) Then Return

            Select Case If(context.Mode, String.Empty).Trim().ToLowerInvariant()
                Case "run"
                    If line.StartsWith("정리 시작: ", StringComparison.OrdinalIgnoreCase) OrElse
                       line.StartsWith("저장 시작: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CurrentFileIndex = Math.Min(Math.Max(1, context.CompletedFiles + 1), Math.Max(1, context.TotalFiles))
                        context.CurrentFileName = GetFileNameOnly(ExtractValueAfterColon(line))
                        context.StepIndex = 0
                        SendDeliveryCleanerProgress(context, "파일 준비 중...", context.CurrentFileName, fileProgress:=0.05R)
                        Return
                    End If

                    If line.StartsWith("[STEP] ", StringComparison.OrdinalIgnoreCase) Then
                        context.StepIndex = Math.Min(context.StepIndex + 1, Math.Max(1, context.TotalSteps))
                        Dim fileProgress = Math.Min(0.9R, CDbl(context.StepIndex) / CDbl(Math.Max(1, context.TotalSteps)))
                        SendDeliveryCleanerProgress(context, line.Substring(7).Trim(), context.CurrentFileName, fileProgress:=fileProgress)
                        Return
                    End If

                    If line.StartsWith("정리 및 저장 완료: ", StringComparison.OrdinalIgnoreCase) OrElse
                       line.StartsWith("저장 완료: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CompletedFiles = Math.Min(context.TotalFiles, context.CompletedFiles + 1)
                        context.CurrentFileIndex = context.CompletedFiles
                        If String.IsNullOrWhiteSpace(context.CurrentFileName) Then
                            context.CurrentFileName = GetFileNameOnly(ExtractValueAfterColon(line))
                        End If
                        context.StepIndex = context.TotalSteps
                        SendDeliveryCleanerProgress(context, "파일 처리 완료", context.CurrentFileName)
                        Return
                    End If

                    If line.StartsWith("실패: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CompletedFiles = Math.Min(context.TotalFiles, context.CompletedFiles + 1)
                        context.CurrentFileIndex = context.CompletedFiles
                        SendDeliveryCleanerProgress(context, "파일 처리 실패", ExtractValueAfterColon(line))
                        Return
                    End If

                    If line.StartsWith("Design Option CSV 저장: ", StringComparison.OrdinalIgnoreCase) Then
                        SendDeliveryCleanerProgress(context, "Design Option 리포트 저장 중...", GetFileNameOnly(ExtractValueAfterColon(line)), fixedPercent:=98.0R)
                        Return
                    End If

                    If line.StartsWith("Design Option CSV 저장 실패: ", StringComparison.OrdinalIgnoreCase) Then
                        SendDeliveryCleanerProgress(context, "Design Option 리포트 저장 실패", ExtractValueAfterColon(line), fixedPercent:=98.0R)
                        Return
                    End If

                Case "verify"
                    If line.StartsWith("검토 파일 열기: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CurrentFileIndex = Math.Min(Math.Max(1, context.CompletedFiles + 1), Math.Max(1, context.TotalFiles))
                        context.CurrentFileName = GetFileNameOnly(ExtractValueAfterColon(line))
                        SendDeliveryCleanerProgress(context, "검토 중...", context.CurrentFileName, fileProgress:=0.15R)
                        Return
                    End If

                    If line.StartsWith("검토 완료: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CompletedFiles = Math.Min(context.TotalFiles, context.CompletedFiles + 1)
                        context.CurrentFileIndex = context.CompletedFiles
                        SendDeliveryCleanerProgress(context, "검토 완료", If(String.IsNullOrWhiteSpace(context.CurrentFileName), line, context.CurrentFileName))
                        Return
                    End If

                    If line.StartsWith("검토 CSV 저장: ", StringComparison.OrdinalIgnoreCase) Then
                        SendDeliveryCleanerProgress(context, "검토 결과 저장 중...", GetFileNameOnly(ExtractValueAfterColon(line)), fixedPercent:=98.0R)
                        Return
                    End If

                Case "extract"
                    If line.StartsWith("속성값 추출 파일 열기: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CurrentFileIndex = Math.Min(Math.Max(1, context.CompletedFiles + 1), Math.Max(1, context.TotalFiles))
                        context.CurrentFileName = GetFileNameOnly(ExtractValueAfterColon(line))
                        SendDeliveryCleanerProgress(context, "속성값 추출 중...", context.CurrentFileName, fileProgress:=0.15R)
                        Return
                    End If

                    If line.StartsWith("속성값 추출 완료: ", StringComparison.OrdinalIgnoreCase) Then
                        context.CompletedFiles = Math.Min(context.TotalFiles, context.CompletedFiles + 1)
                        context.CurrentFileIndex = context.CompletedFiles
                        SendDeliveryCleanerProgress(context, "속성값 추출 완료", If(String.IsNullOrWhiteSpace(context.CurrentFileName), line, context.CurrentFileName))
                        Return
                    End If

                    If line.StartsWith("속성값 추출 CSV 저장: ", StringComparison.OrdinalIgnoreCase) Then
                        SendDeliveryCleanerProgress(context, "속성값 추출 결과 저장 중...", GetFileNameOnly(ExtractValueAfterColon(line)), fixedPercent:=98.0R)
                        Return
                    End If
            End Select
        End Sub

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

            Dim progress As New DeliveryCleanerProgressContext With {
                .Title = "Delivery Cleaner",
                .Mode = "run",
                .TotalFiles = CountExistingDeliveryCleanerFiles(settings.FilePaths),
                .TotalSteps = 12
            }
            Dim logger As Action(Of String) =
                Sub(line)
                    AppendDeliveryCleanerLog(line)
                    UpdateDeliveryCleanerProgressFromLog(progress, line)
                End Sub
            SendDeliveryCleanerProgress(progress, "Preparing...", "", fixedPercent:=0.0R)
            AppendDeliveryCleanerLog("정리 시작")

            Try
                Dim session = BatchCleanService.CleanAndSave(app, settings, logger)
                session.DesignOptionAuditCsvPath = ConvertDeliveryCleanerCsvToXlsx(session.DesignOptionAuditCsvPath, "Design Option 검토")
                SyncLock _deliveryCleanerLock
                    _deliveryCleanerSettings = settings.Clone()
                    _deliveryCleanerSession = session
                End SyncLock

                SendDeliveryCleanerProgress(progress, "Completed", "", complete:=True)
                SendToWeb("deliverycleaner:run-done", New With {
                    .ok = True,
                    .state = BuildDeliveryCleanerStatePayload(),
                    .summary = BuildDeliveryCleanerRunSummary(session)
                })
            Catch ex As Exception
                AppendDeliveryCleanerLog("정리 중 오류: " & ex.Message)
                SendDeliveryCleanerProgress(progress, "Failed", ex.Message, fixedPercent:=progress.LastPercent)
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
            Dim progress As New DeliveryCleanerProgressContext With {
                .Title = "Verification",
                .Mode = "verify",
                .TotalFiles = CountExistingDeliveryCleanerFiles(targetPaths)
            }
            Dim logger As Action(Of String) =
                Sub(line)
                    AppendDeliveryCleanerLog(line)
                    UpdateDeliveryCleanerProgressFromLog(progress, line)
                End Sub
            SendDeliveryCleanerProgress(progress, "Preparing...", "", fixedPercent:=0.0R)

            Try
                Dim csvPath = VerificationService.VerifyPaths(app, targetPaths, outputFolder, settings, logger)
                Dim exportPath = ConvertDeliveryCleanerCsvToXlsx(csvPath, "정리 결과 검토")
                SyncLock _deliveryCleanerLock
                    If _deliveryCleanerSession Is Nothing Then _deliveryCleanerSession = New BatchPrepareSession()
                    _deliveryCleanerSession.VerificationCsvPath = exportPath
                    If String.IsNullOrWhiteSpace(_deliveryCleanerSession.OutputFolder) Then _deliveryCleanerSession.OutputFolder = outputFolder
                    If _deliveryCleanerSettings Is Nothing Then _deliveryCleanerSettings = settings.Clone()
                End SyncLock

                SendDeliveryCleanerProgress(progress, "Completed", exportPath, complete:=True)
                SendToWeb("deliverycleaner:verify-done", New With {.ok = True, .path = exportPath, .state = BuildDeliveryCleanerStatePayload()})
            Catch ex As Exception
                AppendDeliveryCleanerLog("검토 중 오류: " & ex.Message)
                SendDeliveryCleanerProgress(progress, "Failed", ex.Message, fixedPercent:=progress.LastPercent)
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
            Dim progress As New DeliveryCleanerProgressContext With {
                .Title = "Parameter Extraction",
                .Mode = "extract",
                .TotalFiles = CountExistingDeliveryCleanerFiles(targetPaths)
            }
            Dim logger As Action(Of String) =
                Sub(line)
                    AppendDeliveryCleanerLog(line)
                    UpdateDeliveryCleanerProgressFromLog(progress, line)
                End Sub
            SendDeliveryCleanerProgress(progress, "Preparing...", "", fixedPercent:=0.0R)

            Try
                Dim csvPath = ModelParameterExtractionService.ExportModelParameters(app, targetPaths, outputFolder, parameterNamesCsv, logger)
                Dim exportPath = ConvertDeliveryCleanerCsvToXlsx(csvPath, "속성값 추출")
                SyncLock _deliveryCleanerLock
                    _deliveryCleanerExtractParamsCsv = parameterNamesCsv
                    If _deliveryCleanerSession Is Nothing Then _deliveryCleanerSession = New BatchPrepareSession()
                    If String.IsNullOrWhiteSpace(_deliveryCleanerSession.OutputFolder) Then _deliveryCleanerSession.OutputFolder = outputFolder
                    If _deliveryCleanerSettings Is Nothing Then _deliveryCleanerSettings = settings.Clone()
                End SyncLock

                SendDeliveryCleanerProgress(progress, "Completed", exportPath, complete:=True)
                SendToWeb("deliverycleaner:extract-done", New With {
                    .ok = True,
                    .path = exportPath,
                    .parameterNamesCsv = parameterNamesCsv,
                    .state = BuildDeliveryCleanerStatePayload()
                })
            Catch ex As Exception
                AppendDeliveryCleanerLog("속성값 추출 중 오류: " & ex.Message)
                SendDeliveryCleanerProgress(progress, "Failed", ex.Message, fixedPercent:=progress.LastPercent)
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
                Dim progress As New DeliveryCleanerProgressContext With {
                    .Title = "Log Export",
                    .Mode = "log",
                    .TotalFiles = 1
                }
                SendDeliveryCleanerProgress(progress, "Preparing summary...", "", fixedPercent:=0.1R)
                SaveDeliveryCleanerLogSummaryWorkbook(exportPath, logs)
                SendDeliveryCleanerProgress(progress, "Completed", exportPath, complete:=True)

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

        Private Shared Function ConvertDeliveryCleanerCsvToXlsx(csvPath As String,
                                                                sheetName As String,
                                                                Optional progressKey As String = DeliveryCleanerProgressChannel) As String
            If String.IsNullOrWhiteSpace(csvPath) OrElse Not File.Exists(csvPath) Then Return csvPath
            If String.Equals(Path.GetExtension(csvPath), ".xlsx", StringComparison.OrdinalIgnoreCase) Then Return csvPath

            Dim table = LoadDeliveryCleanerCsvAsDataTable(csvPath, sheetName)
            TrimDeliveryCleanerExportTable(table, sheetName)
            Dim xlsxPath = Path.ChangeExtension(csvPath, ".xlsx")
            ExcelCore.SaveStyledSimple(xlsxPath, sheetName, table, GetDeliveryCleanerGroupHeader(table), autoFit:=True, progressKey:=progressKey)
            EnsureDeliveryCleanerWorkbookBorders(xlsxPath)

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

        Private Shared Sub TrimDeliveryCleanerExportTable(table As DataTable, sheetName As String)
            If table Is Nothing Then Return

            If String.Equals(sheetName, "Design Option 검토", StringComparison.OrdinalIgnoreCase) Then
                If table.Columns.Contains("SourcePath") Then
                    table.Columns.Remove("SourcePath")
                End If
            End If
        End Sub

        Private Shared Function GetDeliveryCleanerGroupHeader(table As DataTable) As String
            If table Is Nothing Then Return Nothing
            If table.Columns.Contains("파일명") Then Return "파일명"
            If table.Columns.Contains("FileName") Then Return "FileName"
            Return Nothing
        End Function

        Private Shared Sub EnsureDeliveryCleanerWorkbookBorders(xlsxPath As String)
            If String.IsNullOrWhiteSpace(xlsxPath) OrElse Not File.Exists(xlsxPath) Then Return

            Dim workbook As XSSFWorkbook = Nothing

            Using readStream As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                workbook = New XSSFWorkbook(readStream)
            End Using

            If workbook Is Nothing Then Return

            Try
                For i As Integer = 0 To workbook.NumberOfSheets - 1
                    Dim sheet = workbook.GetSheetAt(i)
                    If sheet Is Nothing Then Continue For
                    ApplyThinBordersToExistingCells(workbook, sheet)
                Next

                Using writeStream As New FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    workbook.Write(writeStream)
                End Using
            Finally
                workbook.Close()
            End Try
        End Sub

        Private Shared Sub ApplyThinBordersToExistingCells(workbook As IWorkbook, sheet As ISheet)
            If workbook Is Nothing OrElse sheet Is Nothing Then Return

            Dim styleCache As New Dictionary(Of Integer, ICellStyle)()

            For rowIndex As Integer = 0 To sheet.LastRowNum
                Dim row = sheet.GetRow(rowIndex)
                If row Is Nothing OrElse row.LastCellNum <= 0 Then Continue For

                For colIndex As Integer = 0 To row.LastCellNum - 1
                    Dim cell = row.GetCell(colIndex)
                    If cell Is Nothing Then Continue For

                    Dim baseStyle = cell.CellStyle
                    Dim cacheKey As Integer = If(baseStyle Is Nothing, -1, CInt(baseStyle.Index))
                    Dim borderedStyle As ICellStyle = Nothing

                    If Not styleCache.TryGetValue(cacheKey, borderedStyle) Then
                        borderedStyle = workbook.CreateCellStyle()
                        If baseStyle IsNot Nothing Then borderedStyle.CloneStyleFrom(baseStyle)
                        borderedStyle.BorderBottom = BorderStyle.Thin
                        borderedStyle.BorderTop = BorderStyle.Thin
                        borderedStyle.BorderLeft = BorderStyle.Thin
                        borderedStyle.BorderRight = BorderStyle.Thin
                        styleCache(cacheKey) = borderedStyle
                    End If

                    cell.CellStyle = borderedStyle
                Next
            Next
        End Sub

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

        Private Shared Sub SaveDeliveryCleanerLogSummaryWorkbook(filePath As String, logs As IList(Of String))
            Dim table = BuildDeliveryCleanerLogSummaryTable(logs)
            ExcelCore.SaveStyledSimple(filePath, "실행 로그 요약", table, "파일명", autoFit:=True)
        End Sub

        Private Shared Function BuildDeliveryCleanerLogSummaryTable(logs As IList(Of String)) As DataTable
            Dim table As New DataTable("실행 로그 요약")
            table.Columns.Add("파일명", GetType(String))
            table.Columns.Add("작업", GetType(String))
            table.Columns.Add("항목", GetType(String))
            table.Columns.Add("상태", GetType(String))
            table.Columns.Add("상세", GetType(String))

            For Each item In BuildDeliveryCleanerLogSummaryRows(logs)
                table.Rows.Add(
                    If(item.FileName, String.Empty),
                    If(item.WorkType, String.Empty),
                    If(item.ItemName, String.Empty),
                    If(item.Status, String.Empty),
                    If(item.Detail, String.Empty))
            Next

            If table.Rows.Count = 0 Then
                table.Rows.Add("", "실행 로그", "기록 없음", "O", "요약할 로그가 없습니다.")
            End If

            Return table
        End Function

        Private Shared Function BuildDeliveryCleanerLogSummaryRows(logs As IList(Of String)) As List(Of DeliveryCleanerLogSummaryRow)
            Dim result As New List(Of DeliveryCleanerLogSummaryRow)()
            Dim currentCleanFile As String = String.Empty
            Dim currentVerifyFile As String = String.Empty
            Dim currentExtractFile As String = String.Empty
            Dim currentPurgeFile As String = String.Empty
            Dim currentStep As DeliveryCleanerStepTracker = Nothing

            For Each rawLine In If(logs, Array.Empty(Of String)())
                Dim line = NormalizeDeliveryCleanerLogLine(rawLine)
                If String.IsNullOrWhiteSpace(line) Then Continue For

                If line.StartsWith("정리 시작: ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("저장 시작: ", StringComparison.OrdinalIgnoreCase) Then
                    FinalizeDeliveryCleanerStep(result, currentStep)
                    currentStep = Nothing
                    currentCleanFile = GetFileNameOnly(ExtractValueAfterColon(line))
                    Continue For
                End If

                If line.StartsWith("[STEP] ", StringComparison.OrdinalIgnoreCase) Then
                    FinalizeDeliveryCleanerStep(result, currentStep)
                    currentStep = New DeliveryCleanerStepTracker With {
                        .FileName = currentCleanFile,
                        .WorkType = "정리",
                        .ItemName = line.Substring(7).Trim()
                    }
                    Continue For
                End If

                If line.StartsWith("실패: ", StringComparison.OrdinalIgnoreCase) Then
                    Dim detail = line.Substring(4).Trim()
                    If currentStep IsNot Nothing AndAlso String.IsNullOrWhiteSpace(currentStep.FailureDetail) Then
                        currentStep.FailureDetail = detail
                    End If
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentCleanFile,
                        .WorkType = "정리",
                        .ItemName = "정리 실행",
                        .Status = "X",
                        .Detail = detail
                    })
                    Continue For
                End If

                If line.StartsWith("정리 및 저장 완료: ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("저장 완료: ", StringComparison.OrdinalIgnoreCase) Then
                    FinalizeDeliveryCleanerStep(result, currentStep)
                    currentStep = Nothing
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = GetFileNameOnly(ExtractValueAfterColon(line)),
                        .WorkType = "정리",
                        .ItemName = "저장",
                        .Status = "O",
                        .Detail = "정리 결과 저장 완료"
                    })
                    Continue For
                End If

                If line.StartsWith("검토 파일 열기: ", StringComparison.OrdinalIgnoreCase) Then
                    currentVerifyFile = GetFileNameOnly(ExtractValueAfterColon(line))
                    Continue For
                End If

                If line.StartsWith("검토 완료: ", StringComparison.OrdinalIgnoreCase) Then
                    Dim detail = ExtractValueAfterColon(line)
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentVerifyFile,
                        .WorkType = "정리 결과 검토",
                        .ItemName = "정리 결과 검토",
                        .Status = If(detail.IndexOf("CHECK", StringComparison.OrdinalIgnoreCase) >= 0, "X", "O"),
                        .Detail = detail
                    })
                    Continue For
                End If

                If line.StartsWith("검토 CSV 저장: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = "",
                        .WorkType = "정리 결과 검토",
                        .ItemName = "엑셀 저장",
                        .Status = "O",
                        .Detail = GetFileNameOnly(ExtractValueAfterColon(line))
                    })
                    Continue For
                End If

                If line.StartsWith("검토 중 오류: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentVerifyFile,
                        .WorkType = "정리 결과 검토",
                        .ItemName = "정리 결과 검토",
                        .Status = "X",
                        .Detail = ExtractValueAfterColon(line)
                    })
                    Continue For
                End If

                If line.StartsWith("속성값 추출 파일 열기: ", StringComparison.OrdinalIgnoreCase) Then
                    currentExtractFile = GetFileNameOnly(ExtractValueAfterColon(line))
                    Continue For
                End If

                If line.StartsWith("속성값 추출 완료: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentExtractFile,
                        .WorkType = "속성값 추출",
                        .ItemName = "속성값 추출",
                        .Status = "O",
                        .Detail = ExtractValueAfterColon(line)
                    })
                    Continue For
                End If

                If line.StartsWith("속성값 추출 CSV 저장: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = "",
                        .WorkType = "속성값 추출",
                        .ItemName = "엑셀 저장",
                        .Status = "O",
                        .Detail = GetFileNameOnly(ExtractValueAfterColon(line))
                    })
                    Continue For
                End If

                If line.StartsWith("속성값 추출 중 오류: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentExtractFile,
                        .WorkType = "속성값 추출",
                        .ItemName = "속성값 추출",
                        .Status = "X",
                        .Detail = ExtractValueAfterColon(line)
                    })
                    Continue For
                End If

                If line.StartsWith("Purge 실행: ", StringComparison.OrdinalIgnoreCase) Then
                    currentPurgeFile = GetFileNameOnly(ExtractValueAfterColon(line))
                    Continue For
                End If

                If line.StartsWith("Purge 후 저장 완료: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = GetFileNameOnly(ExtractValueAfterColon(line)),
                        .WorkType = "Purge",
                        .ItemName = "Purge 일괄처리",
                        .Status = "O",
                        .Detail = "Purge 후 저장 완료"
                    })
                    Continue For
                End If

                If line.StartsWith("Purge 시작 실패: ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("Purge 상태 처리 오류: ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("퍼지 대상 활성 열기 실패: ", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New DeliveryCleanerLogSummaryRow With {
                        .FileName = currentPurgeFile,
                        .WorkType = "Purge",
                        .ItemName = "Purge 일괄처리",
                        .Status = "X",
                        .Detail = ExtractValueAfterColon(line)
                    })
                    Continue For
                End If

                If currentStep IsNot Nothing Then
                    If IsDeliveryCleanerFailureLine(line) Then
                        If String.IsNullOrWhiteSpace(currentStep.FailureDetail) Then
                            currentStep.FailureDetail = line
                        End If
                    ElseIf String.IsNullOrWhiteSpace(currentStep.SuccessDetail) AndAlso IsDeliveryCleanerUsefulSuccessLine(line) Then
                        currentStep.SuccessDetail = line
                    End If
                End If
            Next

            FinalizeDeliveryCleanerStep(result, currentStep)
            Return result
        End Function

        Private Shared Sub FinalizeDeliveryCleanerStep(rows As IList(Of DeliveryCleanerLogSummaryRow), tracker As DeliveryCleanerStepTracker)
            If rows Is Nothing OrElse tracker Is Nothing OrElse String.IsNullOrWhiteSpace(tracker.ItemName) Then Return

            Dim detail = If(String.IsNullOrWhiteSpace(tracker.FailureDetail), tracker.SuccessDetail, tracker.FailureDetail)
            If String.IsNullOrWhiteSpace(detail) Then
                detail = If(String.IsNullOrWhiteSpace(tracker.FailureDetail), "완료", "실패")
            End If

            rows.Add(New DeliveryCleanerLogSummaryRow With {
                .FileName = tracker.FileName,
                .WorkType = tracker.WorkType,
                .ItemName = tracker.ItemName,
                .Status = If(String.IsNullOrWhiteSpace(tracker.FailureDetail), "O", "X"),
                .Detail = detail
            })
        End Sub

        Private Shared Function NormalizeDeliveryCleanerLogLine(rawLine As String) As String
            Dim text = If(rawLine, String.Empty).Trim()
            If text.StartsWith("[") Then
                Dim endIndex = text.IndexOf("]"c)
                If endIndex >= 0 AndAlso endIndex < text.Length - 1 Then
                    text = text.Substring(endIndex + 1).Trim()
                End If
            End If
            Return text
        End Function

        Private Shared Function IsDeliveryCleanerFailureLine(line As String) As Boolean
            If String.IsNullOrWhiteSpace(line) Then Return False

            Dim failureTokens = {
                "실패",
                "오류",
                "충돌",
                "유효하지 않습니다",
                "찾지 못",
                "남아 있습니다"
            }

            Return failureTokens.Any(Function(token) line.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

        Private Shared Function IsDeliveryCleanerUsefulSuccessLine(line As String) As Boolean
            If String.IsNullOrWhiteSpace(line) Then Return False
            If line.StartsWith("Detach + ", StringComparison.OrdinalIgnoreCase) Then Return True
            If line.StartsWith("모든 워크셋 닫기 일반 열기로 재시도", StringComparison.OrdinalIgnoreCase) Then Return True
            If line.Contains("삭제 결과") Then Return True
            If line.Contains("단계 완료") Then Return True
            If line.Contains("설정 완료") Then Return True
            If line.Contains("설정 적용 완료") Then Return True
            If line.Contains("대상 3D 뷰 생성") Then Return True
            If line.Contains("삭제된 뷰/템플릿") Then Return True
            If line.Contains("건너뜀") Then Return True
            If line.Contains("일괄 입력:") Then Return True
            If line.Contains("선삭제:") Then Return True
            If line.Contains("CSV 저장") Then Return True
            Return False
        End Function

        Private Shared Function ExtractValueAfterColon(line As String) As String
            If String.IsNullOrWhiteSpace(line) Then Return String.Empty
            Dim index = line.IndexOf(": ")
            If index >= 0 AndAlso index < line.Length - 2 Then
                Return line.Substring(index + 2).Trim()
            End If

            index = line.IndexOf(":")
            If index >= 0 AndAlso index < line.Length - 1 Then
                Return line.Substring(index + 1).Trim()
            End If

            Return line.Trim()
        End Function

        Private Shared Function GetFileNameOnly(value As String) As String
            Dim text = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return String.Empty

            Dim slashIndex = Math.Max(text.LastIndexOf("\"c), text.LastIndexOf("/"c))
            If slashIndex >= 0 AndAlso slashIndex < text.Length - 1 Then
                Return text.Substring(slashIndex + 1).Trim()
            End If

            Return text
        End Function

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
