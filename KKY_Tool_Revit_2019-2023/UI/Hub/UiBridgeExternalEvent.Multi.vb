Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports Drawing = System.Drawing
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports WinForms = System.Windows.Forms
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services
Imports KKY_Tool_Revit.Exports

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Class MultiCommonOptions
            Public Property ExtraParams As String = String.Empty
            Public Property TargetFilter As String = String.Empty
            Public Property ExcludeEndDummy As Boolean
            Public Property IncludePointXY As Boolean
            Public Property IncludeLinearMetrics As Boolean
        End Class

        ' === commonoptions:get ===
        Private Sub HandleCommonOptionsGet(app As UIApplication, payload As Object)
            Try
                Dim stored = HubCommonOptionsStorageService.Load()
                SendToWeb("commonoptions:loaded", New With {
                    .extraParamsText = stored.ExtraParamsText,
                    .targetFilterText = stored.TargetFilterText,
                    .excludeEndDummy = stored.ExcludeEndDummy,
                    .includePointXY = stored.IncludePointXY,
                    .includeLinearMetrics = stored.IncludeLinearMetrics
                })
            Catch ex As Exception
                SendToWeb("commonoptions:loaded", New With {
                    .extraParamsText = "",
                    .targetFilterText = "",
                    .excludeEndDummy = False,
                    .includePointXY = False,
                    .includeLinearMetrics = False,
                    .errorMessage = ex.Message
                })
            End Try
        End Sub

        ' === commonoptions:save ===
        Private Sub HandleCommonOptionsSave(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim extraText As String = Convert.ToString(GetProp(pd, "extraParamsText"))
            Dim filterText As String = Convert.ToString(GetProp(pd, "targetFilterText"))
            Dim excludeEndDummy As Boolean = SafeBoolObj(GetProp(pd, "excludeEndDummy"), False)
            Dim includePointXY As Boolean = SafeBoolObj(GetProp(pd, "includePointXY"), False)
            Dim includeLinearMetrics As Boolean = SafeBoolObj(GetProp(pd, "includeLinearMetrics"), False)
            Dim options As New HubCommonOptionsStorageService.HubCommonOptions() With {
                .ExtraParamsText = If(extraText, String.Empty),
                .TargetFilterText = If(filterText, String.Empty),
                .ExcludeEndDummy = excludeEndDummy,
                .IncludePointXY = includePointXY,
                .IncludeLinearMetrics = includeLinearMetrics
            }

            Dim ok = HubCommonOptionsStorageService.Save(options)
            SendToWeb("commonoptions:saved", New With {.ok = ok})
        End Sub

        Private Class MultiConnectorOptions
            Public Property Enabled As Boolean
            Public Property Tol As Double = 1.0R
            Public Property Unit As String = "inch"
            Public Property Param As String = "Comments"
            Public Property IncludePointXY As Boolean
            Public Property IncludeLinearMetrics As Boolean
        End Class

        Private Class MultiPmsOptions
            Public Property Enabled As Boolean
            Public Property NdRound As Integer = 3
            Public Property TolMm As Double = 0.01R
            Public Property ClassMatch As Boolean
        End Class

        Private Class MultiGuidOptions
            Public Property Enabled As Boolean
            Public Property IncludeFamily As Boolean
            Public Property IncludeAnnotation As Boolean
        End Class

        Private Class MultiFamilyLinkOptions
            Public Property Enabled As Boolean
            Public Property Targets As List(Of FamilyLinkTargetParam) = New List(Of FamilyLinkTargetParam)()
        End Class

        Private Class MultiPointsOptions
            Public Property Enabled As Boolean
            Public Property Unit As String = "ft"
        End Class

        Private Class MultiLinkWorksetOptions
            Public Property Enabled As Boolean
            Public Property ApplyDefaultWorksetOnly As Boolean = True
            Public Property UseSyncComment As Boolean
            Public Property SyncComment As String = String.Empty
        End Class

        Private Class MultiRunRequest
            Public Property Common As MultiCommonOptions = New MultiCommonOptions()
            Public Property Connector As MultiConnectorOptions = New MultiConnectorOptions()
            Public Property FloorInfo As MultiFloorInfoOptions = New MultiFloorInfoOptions()
            Public Property Pms As MultiPmsOptions = New MultiPmsOptions()
            Public Property Guid As MultiGuidOptions = New MultiGuidOptions()
            Public Property FamilyLink As MultiFamilyLinkOptions = New MultiFamilyLinkOptions()
            Public Property Points As MultiPointsOptions = New MultiPointsOptions()
            Public Property LinkWorkset As MultiLinkWorksetOptions = New MultiLinkWorksetOptions()
            Public Property UseActiveDocument As Boolean = False
            Public Property RvtPaths As List(Of String) = New List(Of String)()
        End Class

        Private Class MultiRunItem
            Public Property File As String = ""
            Public Property Status As String = ""
            Public Property Reason As String = ""
            Public Property Phase As String = ""
            Public Property ElapsedMs As Long
        End Class

        Private Shared ReadOnly _multiLock As New Object()
        Private Shared _multiQueue As List(Of String)
        Private Shared _multiTotal As Integer
        Private Shared _multiIndex As Integer
        Private Shared _multiActive As Boolean
        Private Shared _multiPending As Boolean
        Private Shared _multiBusy As Boolean
        Private Shared _multiRequest As MultiRunRequest
        Private Shared _multiApp As UIApplication
        Private Shared _multiIdlingBound As Boolean
        Private Shared _activeLinkWorksetReopenPending As Boolean
        Private Shared _activeLinkWorksetReopenQueued As Boolean
        Private Shared _activeLinkWorksetReopenPath As String = String.Empty
        Private Shared _activeLinkWorksetReopenName As String = String.Empty

        Private Shared _multiConnectorRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiConnectorExtras As List(Of String)
        Private Shared _multiPmsClassRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiPmsSizeRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiPmsRoutingRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiGuidProject As DataTable
        Private Shared _multiGuidFamilyDetail As DataTable
        Private Shared _multiGuidFamilyIndex As DataTable
        Private Shared _multiFamilyLinkRows As List(Of FamilyLinkAuditRow)
        Private Shared _multiPointRows As List(Of ExportPointsService.Row)
        Private Shared _multiLinkWorksetRows As List(Of LinkWorksetAuditRow)
        Private Shared _multiRunItems As List(Of MultiRunItem)

        ' === hub:pick-rvt ===
        ' payload: none
        ' response: hub:rvt-picked { paths:[string] }
        Private Sub HandleMultiPickRvt()
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                Dim files As String() = dlg.FileNames
                SendToWebAfterDialog("hub:rvt-picked", New With {.paths = files})
            End Using
        End Sub

        ' === hub:multi-clear ===
        ' payload: { key?: string }
        Private Sub HandleMultiClear(payload As Object)
            Dim key As String = TryCast(GetProp(payload, "key"), String)
            If String.IsNullOrWhiteSpace(key) Then
                ResetMultiCaches()
                Return
            End If
            Select Case key.ToLowerInvariant()
                Case "connector"
                    _multiConnectorRows = Nothing
                    _multiConnectorExtras = Nothing
                Case "floorinfo"
                    ClearMultiFloorInfoCache()
                Case "pms"
                    _multiPmsClassRows = Nothing
                    _multiPmsSizeRows = Nothing
                    _multiPmsRoutingRows = Nothing
                Case "guid"
                    _multiGuidProject = Nothing
                    _multiGuidFamilyDetail = Nothing
                    _multiGuidFamilyIndex = Nothing
                Case "familylink"
                    _multiFamilyLinkRows = Nothing
                Case "points"
                    _multiPointRows = Nothing
                Case "linkworkset"
                    _multiLinkWorksetRows = Nothing
            End Select
        End Sub

        ' === hub:multi-run ===
        ' payload:
        '  { rvtPaths:[], commonOptions:{extraParams,targetFilter,excludeEndDummy},
        '    features:{connector,pms,guid,familylink,points} }
        ' response:
        '  hub:multi-progress {percent,message,detail}
        '  hub:multi-done { summary:{key:{rows}} }
        Private Sub HandleMultiRun(app As UIApplication, payload As Object)
            Dim req As MultiRunRequest = ParseMultiRequest(payload)
            If req Is Nothing Then
                SendToWeb("hub:multi-error", New With {.message = "요청 정보가 올바르지 않습니다."})
                Return
            End If
            If Not AnyFeatureEnabled(req) Then
                SendToWeb("hub:multi-error", New With {.message = "선택된 기능이 없습니다."})
                Return
            End If

            If ShouldOfferLegacyManageLinksSwitch(app, req) Then
                Dim choice = ConfirmLegacyManageLinksSwitch(app)
                If choice = TaskDialogResult.Yes Then
                    Dim switchMessage As String = ""
                    If TryEnableLegacyManageLinksAndRestart(app, switchMessage) Then
                        SendToWeb("host:info", New With {.message = switchMessage})
                        SendToWeb("hub:multi-canceled", New With {.message = switchMessage})
                        Return
                    End If

                    SendToWeb("host:warn", New With {.message = switchMessage})
                    SendToWeb("hub:multi-error", New With {.message = switchMessage})
                    Return
                End If

                If choice = TaskDialogResult.No Then
                    SendToWeb("host:info", New With {.message = "[linkworkset] Legacy Manage Links 전환을 취소해 기능 실행을 중단했습니다."})
                    SendToWeb("hub:multi-canceled", New With {.message = "기능 실행을 취소했습니다."})
                    Return
                End If
            End If

            ' 현재 활성 문서(열려있는 파일)로 즉시 검토
            If req.UseActiveDocument Then
                Dim uidoc = app.ActiveUIDocument
                Dim doc As Document = Nothing
                Try
                    If uidoc IsNot Nothing Then doc = uidoc.Document
                Catch
                    doc = Nothing
                End Try
                If doc Is Nothing Then
                    SendToWeb("hub:multi-error", New With {.message = "현재 활성 문서를 찾을 수 없습니다."})
                    Return
                End If

                Dim safeName As String = ""
                Dim docPath As String = ""
                Try
                    safeName = doc.Title
                Catch
                    safeName = ""
                End Try
                Try
                    docPath = doc.PathName
                Catch
                    docPath = ""
                End Try
                If String.IsNullOrWhiteSpace(docPath) Then docPath = safeName

                If ShouldWarnActiveLinkWorksetRefresh(req) Then
                    If Not ConfirmActiveLinkWorksetRefresh(doc, safeName) Then
                        SendToWeb("host:info", New With {.message = "[linkworkset] 활성 문서 실행 취소"})
                        SendToWeb("hub:multi-canceled", New With {.message = "기능 실행을 취소했습니다."})
                        Return
                    End If
                End If

                req.RvtPaths = New List(Of String) From {docPath}
                Dim started = Date.Now
                ExecuteActiveMultiRun(app, req, docPath, safeName, started, False)
                Return
            End If

            If req.RvtPaths.Count = 0 Then
                SendToWeb("hub:multi-error", New With {.message = "검토할 RVT 파일이 없습니다."})
                Return
            End If

            SyncLock _multiLock
                _multiRequest = req
                _multiQueue = New List(Of String)(req.RvtPaths)
                _multiTotal = req.RvtPaths.Count
                _multiIndex = 0
                _multiActive = True
                _multiPending = True
                _multiBusy = False
                _multiApp = app
                _multiRunItems = New List(Of MultiRunItem)()
                ResetMultiCaches()
                If Not _multiIdlingBound Then
                    AddHandler app.Idling, AddressOf HandleMultiIdling
                    _multiIdlingBound = True
                End If
            End SyncLock

            ReportMultiProgress(0.0R, "배치 검토 시작", $"{req.RvtPaths.Count}개 파일 준비")
        End Sub

        Private Sub HandleActiveLinkWorksetAfterUiPrime(app As UIApplication,
                                                        req As MultiRunRequest,
                                                        docPath As String,
                                                        safeName As String,
                                                        started As Date,
                                                        ok As Boolean,
                                                        message As String)
            If Not ok Then
                SendToWeb("host:warn", New With {.message = $"[linkworkset-ui] 프라이밍 실패: {message}"})
            Else
                SendToWeb("host:info", New With {.message = $"[linkworkset-ui] 프라이밍 완료 | {safeName}"})
            End If

            ExecuteActiveMultiRun(app, req, docPath, safeName, started, ok)
        End Sub

        Private Sub ExecuteActiveMultiRun(app As UIApplication,
                                          req As MultiRunRequest,
                                          docPath As String,
                                          safeName As String,
                                          started As Date,
                                          linkWorksetUiPrimed As Boolean)
            Dim uidoc = app.ActiveUIDocument
            Dim doc As Document = Nothing
            Try
                If uidoc IsNot Nothing Then doc = uidoc.Document
            Catch
                doc = Nothing
            End Try
            If doc Is Nothing Then
                SendToWeb("hub:multi-error", New With {.message = "현재 활성 문서를 찾을 수 없습니다."})
                Return
            End If

            SyncLock _multiLock
                _multiRequest = req
                _multiQueue = Nothing
                _multiTotal = 1
                _multiIndex = 1
                _multiActive = False
                _multiPending = False
                _multiBusy = False
                _multiApp = app
                _multiRunItems = New List(Of MultiRunItem)()
                ResetMultiCaches()
            End SyncLock

            Dim saveNeeded As Boolean = False
            ReportMultiProgress(0.0R, "현재 파일 검토 시작", safeName)
            Try
                saveNeeded = RunMultiForDocument(app, doc, docPath, safeName, 0.0R, linkWorksetUiPrimed)
                If saveNeeded Then
                    If ShouldWarnActiveLinkWorksetRefresh(req) Then
                        If Not PersistActiveDocumentForLinkWorkset(doc, safeName, ResolveLinkWorksetSyncComment(req)) Then
                            Throw New InvalidOperationException("활성 문서를 저장 또는 동기화하지 못했습니다.")
                        End If
                        ScheduleActiveLinkWorksetReopen(docPath, safeName)
                    Else
                        Try
                            doc.Save()
                            SendToWeb("host:info", New With {.message = $"[linkworkset] 활성 문서 저장 완료 | {safeName}"})
                        Catch exSave As Exception
                            SendToWeb("host:warn", New With {.message = $"활성 문서 저장 실패: {exSave.Message}"})
                        End Try
                    End If
                End If
                AppendMultiRunItem(safeName, "success", "", "DONE", started)
            Catch ex As Exception
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}")
                AppendMultiRunItem(safeName, "failed", ex.Message, "RUN", started)
                SendToWeb("hub:multi-error", New With {.message = ex.Message})
                Return
            End Try

            FinishMultiRun()
            If saveNeeded AndAlso ShouldWarnActiveLinkWorksetRefresh(req) Then
                PostActiveLinkWorksetCloseCommand(app, safeName)
            End If
        End Sub

        Private Function ShouldUseForegroundBatchLinkWorkset(req As MultiRunRequest) As Boolean
            Return False
        End Function

        Private Sub StartForegroundBatchLinkWorkset(app As UIApplication, req As MultiRunRequest)
            Dim requestedPath As String = NormalizeMultiPath(req.RvtPaths(0))
            Dim safeName As String = Path.GetFileName(requestedPath)
            Dim started = Date.Now

            SyncLock _multiLock
                _multiRequest = req
                _multiQueue = Nothing
                _multiTotal = 1
                _multiIndex = 1
                _multiActive = False
                _multiPending = False
                _multiBusy = False
                _multiApp = app
                _multiRunItems = New List(Of MultiRunItem)()
                ResetMultiCaches()
            End SyncLock

            If Not File.Exists(requestedPath) Then
                SendToWeb("host:warn", New With {.message = $"[multi-open] 파일 경로 확인 실패 | path={requestedPath}"})
                SendToWeb("hub:multi-error", New With {.message = "파일을 찾을 수 없습니다."})
                AppendMultiRunItem(safeName, "skipped", "파일을 찾을 수 없습니다.", "OPEN", started)
                FinishMultiRun()
                Return
            End If

            ReportMultiProgress(0.0R, "전면 배치 준비 중", safeName)

            Dim openPath As String = requestedPath
            Dim createdLocal As Boolean = False
            Try
                Dim fileInfo = TryExtractMultiBasicFileInfo(requestedPath)
                If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                    openPath = CreateMultiNewLocalPath(requestedPath)
                    createdLocal = True
                    SendToWeb("host:info", New With {
                        .message = $"[multi-open] 중앙파일을 임시 로컬로 생성 | central={requestedPath} | local={openPath}"
                    })
                End If

                Dim mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(openPath)
                Dim activated = app.OpenAndActivateDocument(mp, BuildOpenOptions(mp, False), False)
                Dim doc As Document = Nothing
                Try
                    If activated IsNot Nothing Then doc = activated.Document
                Catch
                    doc = Nothing
                End Try

                If doc Is Nothing Then
                    Throw New InvalidOperationException("전면 배치 문서를 활성화하지 못했습니다.")
                End If

                ReportMultiProgress(0.0R, "파일 열기 완료", safeName)

                Dim linkPrimeCount As Integer = CountTopLevelLinkTypes(doc)
                If req.LinkWorkset IsNot Nothing AndAlso req.LinkWorkset.Enabled AndAlso linkPrimeCount > 0 Then
                    SendToWeb("host:info", New With {.message = $"[linkworkset-ui] 배치 전면 프라이밍 시작 | {safeName} | links={linkPrimeCount}"})
                    If Not Global.KKY_Tool_Revit.Services.LinkWorksetUiPrimeService.Start(
                        app,
                        doc,
                        openPath,
                        safeName,
                        linkPrimeCount,
                        Sub(message)
                            SendToWeb("host:info", New With {.message = "[linkworkset-ui] " & message})
                        End Sub,
                        Sub(ok, message)
                            Enqueue(Sub(app2) _self.HandleForegroundBatchLinkWorksetAfterUiPrime(app2, req, requestedPath, openPath, safeName, started, createdLocal, ok, message))
                        End Sub) Then
                        SendToWeb("host:warn", New With {.message = "[linkworkset-ui] 배치 전면 프라이밍 시작 실패, 기본 흐름으로 계속합니다."})
                        HandleForegroundBatchLinkWorksetAfterUiPrime(app, req, requestedPath, openPath, safeName, started, createdLocal, False, "start-failed")
                    End If
                Else
                    HandleForegroundBatchLinkWorksetAfterUiPrime(app, req, requestedPath, openPath, safeName, started, createdLocal, False, "no-links")
                End If
            Catch ex As Exception
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}")
                AppendMultiRunItem(safeName, "failed", ex.Message, "OPEN", started)
                SendToWeb("hub:multi-error", New With {.message = ex.Message})
                If createdLocal Then
                    TryDeleteMultiTempFile(openPath)
                End If
                FinishMultiRun()
            End Try
        End Sub

        Private Sub HandleForegroundBatchLinkWorksetAfterUiPrime(app As UIApplication,
                                                                 req As MultiRunRequest,
                                                                 requestedPath As String,
                                                                 openPath As String,
                                                                 safeName As String,
                                                                 started As Date,
                                                                 createdLocal As Boolean,
                                                                 uiPrimed As Boolean,
                                                                 message As String)
            If uiPrimed Then
                SendToWeb("host:info", New With {.message = $"[linkworkset-ui] 배치 전면 프라이밍 완료 | {safeName}"})
            ElseIf Not String.IsNullOrWhiteSpace(message) Then
                SendToWeb("host:warn", New With {.message = $"[linkworkset-ui] 배치 전면 프라이밍 미완료 | {safeName} | {message}"})
            End If

            Dim uidoc = app.ActiveUIDocument
            Dim doc As Document = Nothing
            Try
                If uidoc IsNot Nothing Then doc = uidoc.Document
            Catch
                doc = Nothing
            End Try

            If doc Is Nothing Then
                AppendMultiRunItem(safeName, "failed", "전면 활성 문서를 찾을 수 없습니다.", "RUN", started)
                SendToWeb("hub:multi-error", New With {.message = "전면 활성 문서를 찾을 수 없습니다."})
                FinishMultiRun()
                Return
            End If

            Dim saveNeeded As Boolean = False
            Try
                ReportMultiProgress(0.0R, "링크 기본 웍셋 점검 중", safeName)
                saveNeeded = RunMultiForDocument(app, doc, requestedPath, safeName, 0.0R, uiPrimed)
                If saveNeeded Then
                    If createdLocal Then
                        ReportMultiProgress(0.0R, "중앙파일 동기화 중", safeName)
                        Dim syncError As String = ""
                        If Not TrySynchronizeMultiLocalToCentral(doc, ResolveLinkWorksetSyncComment(req), syncError) Then
                            Throw New InvalidOperationException("중앙파일 동기화 실패: " & syncError)
                        End If
                        ReportMultiProgress(0.0R, "중앙파일 동기화 완료", safeName)
                    Else
                        ReportMultiProgress(0.0R, "파일 저장 중", safeName)
                        doc.Save()
                        ReportMultiProgress(0.0R, "파일 저장 완료", safeName)
                    End If
                End If
                AppendMultiRunItem(safeName, "success", "", "DONE", started)
            Catch ex As Exception
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}")
                AppendMultiRunItem(safeName, "failed", ex.Message, "RUN", started)
                SendToWeb("host:warn", New With {.message = $"파일 처리 실패: {safeName} - {ex.Message}"})
            Finally
                TryCloseForegroundBatchDocument(app, safeName)
                If createdLocal Then
                    TryDeleteMultiTempFile(openPath)
                End If
            End Try

            FinishMultiRun()
        End Sub

        Private Shared Sub TryCloseForegroundBatchDocument(app As UIApplication, safeName As String)
            If app Is Nothing Then Return

            Dim uidoc As UIDocument = Nothing
            Try
                uidoc = app.ActiveUIDocument
            Catch
                uidoc = Nothing
            End Try
            If uidoc Is Nothing Then Return

            Try
                Dim closeMethod = uidoc.GetType().GetMethod("SaveAndClose", Type.EmptyTypes)
                If closeMethod IsNot Nothing Then
                    closeMethod.Invoke(uidoc, Nothing)
                    SendToWeb("host:info", New With {.message = $"[multi-open] 전면 배치 문서 닫기 완료 | {safeName}"})
                    Return
                End If
            Catch ex As Exception
                SendToWeb("host:warn", New With {.message = $"[multi-open] 전면 배치 문서 SaveAndClose 실패 | {safeName} | {ex.Message}"})
            End Try

            Try
                Dim cmdId As RevitCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Close)
                app.PostCommand(cmdId)
                SendToWeb("host:info", New With {.message = $"[multi-open] 전면 배치 문서 닫기 요청 | {safeName}"})
            Catch ex As Exception
                SendToWeb("host:warn", New With {.message = $"[multi-open] 전면 배치 문서 닫기 실패 | {safeName} | {ex.Message}"})
            End Try
        End Sub

        Private Sub HandleMultiIdling(sender As Object, e As Autodesk.Revit.UI.Events.IdlingEventArgs)
            Dim shouldRun As Boolean = False
            SyncLock _multiLock
                shouldRun = _multiActive AndAlso _multiPending AndAlso Not _multiBusy
                If shouldRun Then
                    _multiPending = False
                    _multiBusy = True
                End If
            End SyncLock
            If shouldRun Then
                ProcessMultiNext(_multiApp)
            End If
        End Sub

        Friend Shared Sub NotifyActiveLinkWorksetDocumentClosed()
            Dim shouldQueue As Boolean = False
            SyncLock _multiLock
                shouldQueue = _activeLinkWorksetReopenPending AndAlso Not _activeLinkWorksetReopenQueued AndAlso Not String.IsNullOrWhiteSpace(_activeLinkWorksetReopenPath)
                If shouldQueue Then
                    _activeLinkWorksetReopenQueued = True
                End If
            End SyncLock

            If shouldQueue Then
                Enqueue(Sub(app) _self.HandlePendingActiveLinkWorksetReopen(app))
            End If
        End Sub

        Private Sub ProcessMultiNext(app As UIApplication)
            Dim filePath As String = Nothing
            SyncLock _multiLock
                If _multiQueue IsNot Nothing AndAlso _multiIndex < _multiQueue.Count Then
                    filePath = _multiQueue(_multiIndex)
                    _multiIndex += 1
                End If
            End SyncLock

            If String.IsNullOrWhiteSpace(filePath) Then
                FinishMultiRun()
                Return
            End If

            filePath = NormalizeMultiPath(filePath)
            Dim safeName As String = System.IO.Path.GetFileName(filePath)
            Dim basePct As Double = If(_multiTotal > 0, CDbl(_multiIndex - 1) / CDbl(_multiTotal), 0.0R)
            ReportMultiProgress(basePct * 100.0R, "파일 여는 중", safeName)

            Dim doc As Document = Nothing
            Dim requestedPath As String = filePath
            Dim openPath As String = filePath
            Dim createdLocal As Boolean = False
            Dim fileInfo As BasicFileInfo = Nothing
            Dim phase As String = "OPEN"
            Dim started = Date.Now
            Try
                If Not System.IO.File.Exists(requestedPath) Then
                    Dim dirPath As String = ""
                    Dim dirExists As Boolean = False
                    Try
                        dirPath = System.IO.Path.GetDirectoryName(requestedPath)
                        dirExists = (Not String.IsNullOrWhiteSpace(dirPath) AndAlso System.IO.Directory.Exists(dirPath))
                    Catch
                    End Try

                    SendToWeb("host:warn", New With {
                        .message = $"[multi-open] 파일 경로 확인 실패 | path={requestedPath} | len={If(requestedPath, String.Empty).Length} | dir={dirPath} | dirExists={If(dirExists, "Y", "N")}"
                    })
                    ReportMultiProgress(basePct * 100.0R, "파일을 찾을 수 없습니다.", safeName)
                    AppendMultiConnectorError(safeName, "파일을 찾을 수 없습니다.")
                    AppendMultiRunItem(safeName, "skipped", "파일을 찾을 수 없습니다.", "OPEN", started)
                    GoTo NextItem
                End If

                fileInfo = TryExtractMultiBasicFileInfo(requestedPath)
                If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                    openPath = CreateMultiNewLocalPath(requestedPath)
                    createdLocal = True
                    SendToWeb("host:info", New With {
                        .message = $"[multi-open] 중앙파일을 임시 로컬로 생성 | central={requestedPath} | local={openPath}"
                    })
                End If

                Dim mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(openPath)

                Dim preferConnectorWorksets As Boolean = False
                Try
                    preferConnectorWorksets = (_multiRequest IsNot Nothing AndAlso _multiRequest.Connector IsNot Nothing AndAlso _multiRequest.Connector.Enabled)
                Catch
                    preferConnectorWorksets = False
                End Try

                doc = app.Application.OpenDocumentFile(mp, BuildOpenOptions(mp, preferConnectorWorksets))
                ReportMultiProgress(basePct * 100.0R, "파일 열기 완료", safeName)
                phase = "RUN"

                Dim rowStartIndex As Integer = If(_multiLinkWorksetRows, New List(Of LinkWorksetAuditRow)()).Count
                Dim saveNeeded = RunMultiForDocument(app, doc, requestedPath, safeName, basePct, False)

                If ShouldRetryLinkWorksetWithHostWorksets(_multiRequest) Then
                    Dim retryWorksetNames = CollectRetryHostWorksetNames(GetMultiLinkRowsSince(rowStartIndex))
                    If retryWorksetNames.Count > 0 Then
                        SendToWeb("host:info", New With {
                            .message = $"[linkworkset] 닫힌 host workset 재시도 | {safeName} | worksets={String.Join(", ", retryWorksetNames)}"
                        })

                        Try
                            doc.Close(False)
                        Catch
                        End Try
                        doc = Nothing

                        TrimMultiLinkWorksetRows(rowStartIndex)

                        Dim retryOptions = BuildOpenOptions(mp, preferConnectorWorksets, retryWorksetNames)
                        doc = app.Application.OpenDocumentFile(mp, retryOptions)
                        ReportMultiProgress(basePct * 100.0R, "필요 host workset 열기 후 재시도", safeName)
                        saveNeeded = RunMultiForDocument(app, doc, requestedPath, safeName, basePct, False)
                    End If
                End If

                If saveNeeded Then
                    phase = "SAVE"
                    If createdLocal Then
                        ReportMultiProgress(basePct * 100.0R, "중앙파일 동기화 중", safeName)
                        Dim syncError As String = ""
                        If Not TrySynchronizeMultiLocalToCentral(doc, ResolveLinkWorksetSyncComment(_multiRequest), syncError) Then
                            Throw New InvalidOperationException("중앙파일 동기화 실패: " & syncError)
                        End If
                        ReportMultiProgress(basePct * 100.0R, "중앙파일 동기화 완료", safeName)
                    Else
                        ReportMultiProgress(basePct * 100.0R, "파일 저장 중", safeName)
                        doc.Save()
                        ReportMultiProgress(basePct * 100.0R, "파일 저장 완료", safeName)
                    End If
                End If
                AppendMultiRunItem(safeName, "success", "", "DONE", started)
            Catch ex As Exception
                AppendMultiConnectorError(safeName, $"파일 처리 실패: {ex.Message}")
                ReportMultiProgress(basePct * 100.0R, "파일 처리 실패 (건너뜀)", safeName)
                AppendMultiRunItem(safeName, "failed", ex.Message, phase, started)
                SendToWeb("host:warn", New With {.message = $"파일 처리 실패: {safeName} - {ex.Message}"})
            Finally
                If doc IsNot Nothing Then
                    Try
                        doc.Close(False)
                    Catch
                    End Try
                End If
                If createdLocal Then
                    TryDeleteMultiTempFile(openPath)
                End If
            End Try

NextItem:
            SyncLock _multiLock
                _multiBusy = False
                If _multiQueue IsNot Nothing AndAlso _multiIndex < _multiQueue.Count Then
                    _multiPending = True
                Else
                    _multiActive = False
                End If
            End SyncLock

            If Not _multiActive Then
                FinishMultiRun()
            End If
        End Sub

        Private Function RunMultiForDocument(app As UIApplication,
                                             doc As Document,
                                             path As String,
                                             safeName As String,
                                             basePct As Double,
                                             Optional linkWorksetUiPrimed As Boolean = False) As Boolean
            Dim steps As Integer = CountEnabledFeatures(_multiRequest)
            Dim stepIndex As Integer = 0
            Dim saveNeeded As Boolean = False

            If _multiRequest.Connector.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 실행 중", safeName)
                Dim extras = BuildConnectorExtraParams(_multiRequest.Common.ExtraParams,
                                                       _multiRequest.Connector.IncludePointXY,
                                                       _multiRequest.Connector.IncludeLinearMetrics)
                Dim rows = ConnectorDiagnosticsService.RunOnDocument(doc, _multiRequest.Connector.Tol, _multiRequest.Connector.Unit, _multiRequest.Connector.Param, extras, _multiRequest.Common.TargetFilter, _multiRequest.Common.ExcludeEndDummy, Sub(pct, msg)
                                                                                                                                                                                                                                                         Dim overallPct = ((basePct + (pct / 100.0R) / Math.Max(_multiTotal, 1)) * 100.0R)
                                                                                                                                                                                                                                                         ReportMultiProgress(overallPct, "커넥터 진단 실행 중", $"{safeName} · {msg}")
                                                                                                                                                                                                                                                     End Sub)
                If rows IsNot Nothing AndAlso rows.Count > 0 Then
                    For Each row In rows
                        If row IsNot Nothing Then row("File") = safeName
                    Next
                    If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
                    _multiConnectorRows.AddRange(rows)
                    _multiConnectorExtras = extras
                Else
                    If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
                    _multiConnectorRows.Add(New Dictionary(Of String, Object) From {
                        {"File", safeName},
                        {"ConnectionType", "OK"},
                        {"ParamCompare", "OK"},
                        {"Status", "오류 없음"},
                        {"ErrorMessage", ""}
                    })
                    _multiConnectorExtras = extras
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "커넥터 진단 완료", safeName)
            End If

            If _multiRequest.FloorInfo.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "층정보 검토 실행 중", safeName)
                RunFloorInfoMultiForDocument(doc, safeName, basePct)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "층정보 검토 완료", safeName)
            End If

            If _multiRequest.Pms.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 검토 실행 중", safeName)
                If _pmsRows Is Nothing OrElse _pmsRows.Count = 0 Then
                    SendToWeb("host:warn", New With {.message = "PMS Excel이 등록되지 않았습니다."})
                Else
                    Dim opts As New SegmentPmsCheckService.ExtractOptions With {
                        .NdRound = _multiRequest.Pms.NdRound,
                        .ToleranceMm = _multiRequest.Pms.TolMm
                    }
                    Dim compareOpts As New SegmentPmsCheckService.CompareOptions With {
                        .NdRound = _multiRequest.Pms.NdRound,
                        .TolMm = _multiRequest.Pms.TolMm,
                        .ClassMatch = _multiRequest.Pms.ClassMatch
                    }
                    Dim ds = SegmentPmsCheckService.ExtractFromDocument(app, doc, path, opts, Nothing)
                    Dim groups = SegmentPmsCheckService.BuildGroups(ds)
                    Dim suggestions = SegmentPmsCheckService.SuggestGroupMappings(groups, _pmsRows)
                    Dim mappings = BuildMappingsFromSuggestions(suggestions)
                    Dim run = SegmentPmsCheckService.RunCompare(ds, _pmsRows, mappings, compareOpts)
                    AppendSegmentPmsRows(run, ds)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "PMS 검토 완료", safeName)
            End If

            If _multiRequest.Guid.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 실행 중", safeName)
                Dim res = GuidAuditService.Run(app, If(_multiRequest.Guid.IncludeFamily, 2, 1), New List(Of String) From {path}, Nothing, Nothing, _multiRequest.Guid.IncludeFamily, _multiRequest.Guid.IncludeAnnotation)
                MergeGuidResult(res)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "GUID 검토 완료", safeName)
            End If

            If _multiRequest.FamilyLink.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "패밀리 연동 검토 실행 중", safeName)
                Dim rows = FamilyLinkAuditService.RunOnDocument(doc, path, _multiRequest.FamilyLink.Targets, Nothing)
                If rows IsNot Nothing Then
                    If _multiFamilyLinkRows Is Nothing Then _multiFamilyLinkRows = New List(Of FamilyLinkAuditRow)()
                    _multiFamilyLinkRows.AddRange(FilterFamilyLinkIssueRows(rows))
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "패밀리 연동 검토 완료", safeName)
            End If

            If _multiRequest.Points.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Point 추출 실행 중", safeName)
                Dim rows = ExportPointsService.RunOnDocument(doc, safeName, Nothing)
                If rows IsNot Nothing Then
                    If _multiPointRows Is Nothing Then _multiPointRows = New List(Of ExportPointsService.Row)()
                    _multiPointRows.AddRange(rows)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Point 추출 완료", safeName)
            End If

            If _multiRequest.LinkWorkset.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "링크 기본 웍셋 점검 중", safeName)
                Dim rows = LinkWorksetAuditService.RunOnDocument(doc, path, _multiRequest.LinkWorkset.ApplyDefaultWorksetOnly, Nothing, linkWorksetUiPrimed)
                If rows IsNot Nothing Then
                    If _multiLinkWorksetRows Is Nothing Then _multiLinkWorksetRows = New List(Of LinkWorksetAuditRow)()
                    _multiLinkWorksetRows.AddRange(rows)
                    PublishLinkWorksetDiagnostics(rows, safeName)
                    saveNeeded = rows.Any(Function(r) r IsNot Nothing AndAlso r.Applied)
                End If
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "링크 기본 웍셋 점검 완료", safeName)
            End If
            Return saveNeeded
        End Function

        Private Shared Function ShouldWarnActiveLinkWorksetRefresh(req As MultiRunRequest) As Boolean
            Return req IsNot Nothing AndAlso
                   req.UseActiveDocument AndAlso
                   req.LinkWorkset IsNot Nothing AndAlso
                   req.LinkWorkset.Enabled
        End Function

        Private Shared Function ShouldOfferLegacyManageLinksSwitch(app As UIApplication, req As MultiRunRequest) As Boolean
            If app Is Nothing OrElse req Is Nothing Then Return False
            If req.LinkWorkset Is Nothing OrElse Not req.LinkWorkset.Enabled Then Return False

            Dim versionNumber = GetRevitMajorVersion(app)
            If versionNumber < 2025 Then Return False

            Return Not IsLegacyManageLinksEnabled(versionNumber)
        End Function

        Private Shared Function ConfirmLegacyManageLinksSwitch(app As UIApplication) As TaskDialogResult
            Dim versionText As String = "현재 버전"
            Try
                versionText = $"Revit {GetRevitMajorVersion(app)}"
            Catch
            End Try

            Dim result = ShowHubStyledYesNoDialog(
                "링크 기본 웍셋 점검/적용",
                "Legacy Manage Links 설정이 꺼져 있습니다.",
                $"{versionText}에서는 새 Manage Links 대화상자가 불안정할 수 있습니다." & vbCrLf &
                "예를 누르면 Legacy Manage Links 설정을 적용하고 Revit을 자동으로 다시 실행합니다." & vbCrLf &
                "아니오를 누르면 기능 실행을 취소하고 허브로 돌아갑니다.",
                "설정 후 재시작",
                "취소",
                "이 설정은 Revit.ini를 수정한 뒤 Revit을 다시 시작해야 적용됩니다. 재시작 후 Legacy Manage Links가 활성화됩니다.")

            Return If(result, TaskDialogResult.Yes, TaskDialogResult.No)
        End Function

        Private Shared Function ShowHubStyledYesNoDialog(title As String,
                                                         mainInstruction As String,
                                                         mainContent As String,
                                                         yesText As String,
                                                         noText As String,
                                                         Optional noteText As String = Nothing) As Boolean
            Dim isLight = String.Equals(HubHostWindow.CurrentThemeKey, "light", StringComparison.OrdinalIgnoreCase)

            Dim backColor = If(isLight, Drawing.Color.FromArgb(245, 248, 252), Drawing.Color.FromArgb(10, 18, 36))
            Dim panelColor = If(isLight, Drawing.Color.FromArgb(255, 255, 255), Drawing.Color.FromArgb(18, 31, 58))
            Dim borderColor = If(isLight, Drawing.Color.FromArgb(201, 214, 234), Drawing.Color.FromArgb(54, 82, 126))
            Dim titleColor = If(isLight, Drawing.Color.FromArgb(23, 38, 68), Drawing.Color.FromArgb(238, 244, 255))
            Dim textColor = If(isLight, Drawing.Color.FromArgb(66, 84, 117), Drawing.Color.FromArgb(184, 204, 238))
            Dim accentColor = If(isLight, Drawing.Color.FromArgb(44, 153, 255), Drawing.Color.FromArgb(48, 170, 255))
            Dim buttonGhostBack = If(isLight, Drawing.Color.FromArgb(239, 244, 252), Drawing.Color.FromArgb(25, 40, 69))

            Using form As New WinForms.Form()
                form.Text = title
                form.StartPosition = WinForms.FormStartPosition.CenterScreen
                form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
                form.MinimizeBox = False
                form.MaximizeBox = False
                form.ShowInTaskbar = False
                form.TopMost = True
                form.BackColor = backColor
                form.ClientSize = New Drawing.Size(1120, 560)
                form.MinimumSize = New Drawing.Size(1120, 560)
                form.Font = New Drawing.Font("Malgun Gothic", 10.5F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)

                Dim outer As New WinForms.Panel() With {
                    .Dock = WinForms.DockStyle.Fill,
                    .Padding = New WinForms.Padding(28),
                    .BackColor = backColor
                }

                Dim card As New WinForms.Panel() With {
                    .Dock = WinForms.DockStyle.Fill,
                    .BackColor = panelColor,
                    .Padding = New WinForms.Padding(36, 32, 36, 28),
                    .BorderStyle = WinForms.BorderStyle.FixedSingle
                }

                Dim titleLabel As New WinForms.Label() With {
                    .Dock = WinForms.DockStyle.Top,
                    .Height = 70,
                    .Text = mainInstruction,
                    .ForeColor = titleColor,
                    .Font = New Drawing.Font("Malgun Gothic", 19.0F, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point)
                }
                titleLabel.TextAlign = Drawing.ContentAlignment.MiddleLeft

                Dim contentLabel As New WinForms.Label() With {
                    .Dock = WinForms.DockStyle.Top,
                    .Height = 170,
                    .Text = mainContent,
                    .ForeColor = textColor,
                    .Font = New Drawing.Font("Malgun Gothic", 13.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
                }
                contentLabel.TextAlign = Drawing.ContentAlignment.TopLeft

                Dim noteLabel As New WinForms.Label() With {
                    .Dock = WinForms.DockStyle.Top,
                    .Height = 72,
                    .Text = If(String.IsNullOrWhiteSpace(noteText),
                               "취소하면 허브로 돌아가고, 계속하면 안내된 단계로 진행됩니다.",
                               noteText),
                    .ForeColor = accentColor,
                    .Font = New Drawing.Font("Malgun Gothic", 11.0F, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Point)
                }
                noteLabel.TextAlign = Drawing.ContentAlignment.MiddleLeft

                Dim buttonPanel As New WinForms.FlowLayoutPanel() With {
                    .Dock = WinForms.DockStyle.Bottom,
                    .Height = 72,
                    .FlowDirection = WinForms.FlowDirection.RightToLeft,
                    .WrapContents = False,
                    .BackColor = panelColor,
                    .Padding = New WinForms.Padding(0, 16, 0, 0)
                }

                Dim yesButton As New WinForms.Button() With {
                    .Text = yesText,
                    .Width = 210,
                    .Height = 52,
                    .FlatStyle = WinForms.FlatStyle.Flat,
                    .BackColor = accentColor,
                    .ForeColor = Drawing.Color.White,
                    .TabIndex = 0
                }
                yesButton.FlatAppearance.BorderSize = 0

                Dim noButton As New WinForms.Button() With {
                    .Text = noText,
                    .Width = 148,
                    .Height = 52,
                    .FlatStyle = WinForms.FlatStyle.Flat,
                    .BackColor = buttonGhostBack,
                    .ForeColor = titleColor,
                    .TabIndex = 1
                }
                noButton.FlatAppearance.BorderSize = 1
                noButton.FlatAppearance.BorderColor = borderColor

                AddHandler yesButton.Click, Sub()
                                                form.DialogResult = WinForms.DialogResult.Yes
                                                form.Close()
                                            End Sub
                AddHandler noButton.Click, Sub()
                                               form.DialogResult = WinForms.DialogResult.No
                                               form.Close()
                                           End Sub

                buttonPanel.Controls.Add(yesButton)
                buttonPanel.Controls.Add(noButton)

                card.Controls.Add(buttonPanel)
                card.Controls.Add(noteLabel)
                card.Controls.Add(contentLabel)
                card.Controls.Add(titleLabel)
                outer.Controls.Add(card)
                form.Controls.Add(outer)
                form.AcceptButton = yesButton
                form.CancelButton = noButton
                form.ActiveControl = yesButton

                Dim result = (form.ShowDialog() = WinForms.DialogResult.Yes)
                Try
                    If _host IsNot Nothing Then
                        _host.Activate()
                        _host.Focus()
                    End If
                Catch
                End Try
                Return result
            End Using
        End Function

        Private Shared Function GetRevitMajorVersion(app As UIApplication) As Integer
            If app Is Nothing OrElse app.Application Is Nothing Then Return 0

            Dim raw As String = ""
            Try
                raw = If(app.Application.VersionNumber, "").Trim()
            Catch
                raw = ""
            End Try

            Dim parsed As Integer = 0
            If Integer.TryParse(raw, parsed) Then Return parsed
            Return 0
        End Function

        Private Shared Function GetRevitIniPath(versionNumber As Integer) As String
            If versionNumber <= 0 Then Return String.Empty
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk",
                "Revit",
                "Autodesk Revit " & versionNumber.ToString(),
                "Revit.ini")
        End Function

        Private Shared Function IsLegacyManageLinksEnabled(versionNumber As Integer) As Boolean
            Dim iniPath = GetRevitIniPath(versionNumber)
            If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then Return False

            Try
                For Each line In File.ReadAllLines(iniPath)
                    Dim trimmed = If(line, String.Empty).Trim()
                    If trimmed.StartsWith("EnableOldManageLinksDialog", StringComparison.OrdinalIgnoreCase) Then
                        Dim parts = trimmed.Split({"="c}, 2)
                        If parts.Length = 2 AndAlso String.Equals(parts(1).Trim(), "1", StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function TryEnableLegacyManageLinksAndRestart(app As UIApplication, ByRef message As String) As Boolean
            message = "Legacy Manage Links 설정 적용에 실패했습니다."
            If app Is Nothing Then
                message = "Revit 앱 정보를 찾을 수 없습니다."
                Return False
            End If

            Dim versionNumber = GetRevitMajorVersion(app)
            Dim iniPath = GetRevitIniPath(versionNumber)
            If String.IsNullOrWhiteSpace(iniPath) Then
                message = "Revit.ini 경로를 찾을 수 없습니다."
                Return False
            End If

            Try
                Dim iniDir = Path.GetDirectoryName(iniPath)
                If Not String.IsNullOrWhiteSpace(iniDir) Then Directory.CreateDirectory(iniDir)
                If File.Exists(iniPath) Then
                    Dim backupPath = iniPath & ".bak-" & DateTime.Now.ToString("yyyyMMdd-HHmmss")
                    File.Copy(iniPath, backupPath, True)
                End If
                WriteLegacyManageLinksIni(iniPath)
            Catch ex As Exception
                message = "Legacy Manage Links 설정 저장 실패: " & ex.Message
                Return False
            End Try

            Dim exePath As String = ""
            Try
                exePath = Process.GetCurrentProcess().MainModule.FileName
            Catch ex As Exception
                message = "Revit 재실행 경로를 찾을 수 없습니다: " & ex.Message
                Return False
            End Try

            If String.IsNullOrWhiteSpace(exePath) OrElse Not File.Exists(exePath) Then
                message = "Revit 실행 파일 경로를 확인할 수 없습니다."
                Return False
            End If

            Try
                QueueRevitRelaunchAfterExit(exePath)
            Catch ex As Exception
                message = "Revit 재실행 예약 실패: " & ex.Message
                Return False
            End Try

            message = "Legacy Manage Links 설정을 적용했습니다. Revit을 종료 후 자동으로 다시 실행합니다."
            RequestRevitExit(app)
            Return True
        End Function

        Private Shared Sub WriteLegacyManageLinksIni(iniPath As String)
            Dim lines As New List(Of String)()
            If File.Exists(iniPath) Then
                lines.AddRange(File.ReadAllLines(iniPath))
            End If

            Dim miscIndex As Integer = -1
            Dim keyIndex As Integer = -1

            For i As Integer = 0 To lines.Count - 1
                Dim trimmed = If(lines(i), String.Empty).Trim()
                If String.Equals(trimmed, "[Misc]", StringComparison.OrdinalIgnoreCase) Then
                    miscIndex = i
                    For j As Integer = i + 1 To lines.Count - 1
                        Dim inner = If(lines(j), String.Empty).Trim()
                        If inner.StartsWith("[") AndAlso inner.EndsWith("]") Then Exit For
                        If inner.StartsWith("EnableOldManageLinksDialog", StringComparison.OrdinalIgnoreCase) Then
                            keyIndex = j
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next

            If miscIndex < 0 Then
                If lines.Count > 0 AndAlso Not String.IsNullOrWhiteSpace(lines(lines.Count - 1)) Then lines.Add("")
                lines.Add("[Misc]")
                lines.Add("EnableOldManageLinksDialog=1")
            ElseIf keyIndex >= 0 Then
                lines(keyIndex) = "EnableOldManageLinksDialog=1"
            Else
                lines.Insert(miscIndex + 1, "EnableOldManageLinksDialog=1")
            End If

            File.WriteAllLines(iniPath, lines)
        End Sub

        Private Shared Sub QueueRevitRelaunchAfterExit(revitExePath As String)
            Dim stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim queueDir = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "Relaunch", stamp)
            Directory.CreateDirectory(queueDir)

            Dim scriptPath = Path.Combine(queueDir, "restart-revit.ps1")
            Dim pidText = Process.GetCurrentProcess().Id.ToString()
            Dim escapedExe = revitExePath.Replace("'", "''")
            Dim script = String.Join(Environment.NewLine, New String() {
                "$ErrorActionPreference = 'Stop'",
                "$pidToWait = " & pidText,
                "$exePath = '" & escapedExe & "'",
                "while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 2 }",
                "Start-Sleep -Milliseconds 800",
                "Start-Process -FilePath $exePath"
            })
            File.WriteAllText(scriptPath, script)

            Dim psi As New ProcessStartInfo()
            psi.FileName = "powershell.exe"
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & QuoteProcessArgument(scriptPath)
            psi.UseShellExecute = True
            psi.WorkingDirectory = queueDir
            Process.Start(psi)
        End Sub

        Private Shared Function QuoteProcessArgument(value As String) As String
            If value Is Nothing Then Return """"""
            Return """" & value.Replace("""", "\""") & """"
        End Function

        Private Shared Sub RequestRevitExit(app As UIApplication)
            Try
                Dim cmdId = RevitCommandId.LookupCommandId("ID_APP_EXIT")
                If cmdId IsNot Nothing Then
                    app.PostCommand(cmdId)
                    Return
                End If
            Catch
            End Try

            Try
                Process.GetCurrentProcess().CloseMainWindow()
            Catch
            End Try
        End Sub

        Private Shared Function CountTopLevelLinkTypes(doc As Document) As Integer
            If doc Is Nothing Then Return 0
            Try
                Return New FilteredElementCollector(doc).
                    OfClass(GetType(RevitLinkType)).
                    Cast(Of RevitLinkType)().
                    Count(Function(x) x IsNot Nothing AndAlso Not x.IsNestedLink)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ConfirmActiveLinkWorksetRefresh(doc As Document, safeName As String) As Boolean
            Dim isWorkshared As Boolean = False
            Try
                isWorkshared = (doc IsNot Nothing AndAlso doc.IsWorkshared)
            Catch
                isWorkshared = False
            End Try

            Dim instruction = "이 기능을 실행하면 파일이 자동 동기화 후 재오픈됩니다."
            If Not isWorkshared Then
                instruction = "이 기능을 실행하면 파일이 자동 저장 후 재오픈됩니다."
            End If

            Return ShowHubStyledYesNoDialog(
                "링크 기본 웍셋 점검/적용",
                instruction,
                "재오픈 시 웍셋은 모두 닫힌 상태로 열립니다." & vbCrLf & "계속하시겠습니까?" & vbCrLf & safeName,
                "계속",
                "취소",
                "링크 적용 상태와 웍셋 열림 구성을 안정적으로 반영하려면 저장 또는 동기화 후 다시 열어야 합니다.")
        End Function

        Private Shared Function PersistActiveDocumentForLinkWorkset(doc As Document, safeName As String, syncComment As String) As Boolean
            If doc Is Nothing Then Return False

            Try
                If doc.IsWorkshared Then
                    Dim twc As New TransactWithCentralOptions()
                    Dim swc As New SynchronizeWithCentralOptions()
                    swc.Comment = If(syncComment, String.Empty)
                    Try
                        Dim rel As New RelinquishOptions(True)
                        swc.SetRelinquishOptions(rel)
                    Catch
                    End Try

                    doc.SynchronizeWithCentral(twc, swc)
                    SendToWeb("host:info", New With {.message = $"[linkworkset] 활성 문서 동기화 완료 | {safeName}"})
                    Return True
                End If
            Catch exSync As Exception
                SendToWeb("host:warn", New With {.message = $"활성 문서 동기화 실패, 저장으로 전환합니다: {exSync.Message}"})
            End Try

            Try
                doc.Save()
                SendToWeb("host:info", New With {.message = $"[linkworkset] 활성 문서 저장 완료 | {safeName}"})
                Return True
            Catch exSave As Exception
                SendToWeb("host:warn", New With {.message = $"활성 문서 저장 실패: {exSave.Message}"})
                Return False
            End Try
        End Function

        Private Shared Sub ScheduleActiveLinkWorksetReopen(docPath As String, safeName As String)
            SyncLock _multiLock
                _activeLinkWorksetReopenPending = True
                _activeLinkWorksetReopenQueued = False
                _activeLinkWorksetReopenPath = If(docPath, String.Empty)
                _activeLinkWorksetReopenName = If(safeName, String.Empty)
            End SyncLock
        End Sub

        Private Shared Function ShouldRetryLinkWorksetWithHostWorksets(req As MultiRunRequest) As Boolean
            If req Is Nothing Then Return False
            If req.LinkWorkset Is Nothing OrElse Not req.LinkWorkset.Enabled Then Return False
            Return True
        End Function

        Private Shared Function ResolveLinkWorksetSyncComment(req As MultiRunRequest) As String
            If req Is Nothing OrElse req.LinkWorkset Is Nothing Then Return String.Empty
            If Not req.LinkWorkset.UseSyncComment Then Return String.Empty

            Dim comment = SafeStr(req.LinkWorkset.SyncComment).Trim()
            If String.IsNullOrWhiteSpace(comment) Then
                Return "KKY Tools - 링크 기본 웍셋 적용"
            End If
            Return comment
        End Function

        Private Shared Function GetMultiLinkRowsSince(startIndex As Integer) As List(Of LinkWorksetAuditRow)
            Dim rows = If(_multiLinkWorksetRows, New List(Of LinkWorksetAuditRow)())
            If startIndex <= 0 Then Return New List(Of LinkWorksetAuditRow)(rows)
            If startIndex >= rows.Count Then Return New List(Of LinkWorksetAuditRow)()
            Return rows.Skip(startIndex).ToList()
        End Function

        Private Shared Sub TrimMultiLinkWorksetRows(startIndex As Integer)
            If _multiLinkWorksetRows Is Nothing Then Return
            If startIndex <= 0 Then
                _multiLinkWorksetRows.Clear()
                Return
            End If
            If startIndex >= _multiLinkWorksetRows.Count Then Return
            _multiLinkWorksetRows.RemoveRange(startIndex, _multiLinkWorksetRows.Count - startIndex)
        End Sub

        Private Shared Function CollectRetryHostWorksetNames(rows As IEnumerable(Of LinkWorksetAuditRow)) As List(Of String)
            Dim names As New List(Of String)()
            If rows Is Nothing Then Return names

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim message = If(row.Message, String.Empty)
                Dim diag = If(row.DiagnosticLog, String.Empty)
                If message.IndexOf("closed workset", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   diag.IndexOf("closed workset", StringComparison.OrdinalIgnoreCase) < 0 Then
                    Continue For
                End If

                Dim marker = "hostWorksets="
                Dim idx = diag.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
                If idx < 0 Then Continue For
                Dim rest = diag.Substring(idx + marker.Length)
                Dim endIdx = rest.IndexOf(" || ", StringComparison.Ordinal)
                If endIdx >= 0 Then rest = rest.Substring(0, endIdx)
                For Each token In rest.Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim name = token.Trim()
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                        names.Add(name)
                    End If
                Next
            Next

            Return names
        End Function

        Private Shared Sub ResetActiveLinkWorksetReopenState()
            SyncLock _multiLock
                _activeLinkWorksetReopenPending = False
                _activeLinkWorksetReopenQueued = False
                _activeLinkWorksetReopenPath = String.Empty
                _activeLinkWorksetReopenName = String.Empty
            End SyncLock
        End Sub

        Private Shared Sub PostActiveLinkWorksetCloseCommand(app As UIApplication, safeName As String)
            If app Is Nothing Then
                ResetActiveLinkWorksetReopenState()
                Return
            End If

            Try
                Dim cmdId As RevitCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Close)
                app.PostCommand(cmdId)
                SendToWeb("host:info", New With {.message = $"[linkworkset] 활성 문서 닫기 요청 | {safeName}"})
            Catch ex As Exception
                ResetActiveLinkWorksetReopenState()
                SendToWeb("host:warn", New With {.message = $"활성 문서 자동 닫기 실패: {ex.Message}"})
            End Try
        End Sub

        Private Sub HandlePendingActiveLinkWorksetReopen(app As UIApplication)
            Dim reopenPath As String = String.Empty
            Dim safeName As String = String.Empty
            SyncLock _multiLock
                reopenPath = _activeLinkWorksetReopenPath
                safeName = _activeLinkWorksetReopenName
            End SyncLock

            If String.IsNullOrWhiteSpace(reopenPath) Then
                ResetActiveLinkWorksetReopenState()
                Return
            End If

            Try
                Dim mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(reopenPath)
                app.OpenAndActivateDocument(mp, BuildReopenOpenOptions(), False)
                SendToWeb("host:info", New With {.message = $"[linkworkset] 활성 문서 재오픈 완료 | {If(String.IsNullOrWhiteSpace(safeName), Path.GetFileName(reopenPath), safeName)}"})
            Catch ex As Exception
                SendToWeb("host:warn", New With {.message = $"활성 문서 재오픈 실패: {ex.Message}"})
            Finally
                ResetActiveLinkWorksetReopenState()
            End Try
        End Sub

        Private Sub FinishMultiRun()
            Dim summary As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Connector IsNot Nothing AndAlso _multiRequest.Connector.Enabled Then
                summary("connector") = New With {.rows = If(_multiConnectorRows, New List(Of Dictionary(Of String, Object))()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FloorInfo IsNot Nothing AndAlso _multiRequest.FloorInfo.Enabled Then
                summary("floorinfo") = New With {.rows = GetMultiFloorInfoRowCount()}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Pms IsNot Nothing AndAlso _multiRequest.Pms.Enabled Then
                summary("pms") = New With {.rows = If(_multiPmsSizeRows, New List(Of Dictionary(Of String, Object))()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Guid IsNot Nothing AndAlso _multiRequest.Guid.Enabled Then
                summary("guid") = New With {.rows = If(_multiGuidProject, New DataTable()).Rows.Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FamilyLink IsNot Nothing AndAlso _multiRequest.FamilyLink.Enabled Then
                summary("familylink") = New With {.rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Points IsNot Nothing AndAlso _multiRequest.Points.Enabled Then
                summary("points") = New With {.rows = If(_multiPointRows, New List(Of ExportPointsService.Row)()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.LinkWorkset IsNot Nothing AndAlso _multiRequest.LinkWorkset.Enabled Then
                summary("linkworkset") = New With {.rows = If(_multiLinkWorksetRows, New List(Of LinkWorksetAuditRow)()).Count}
            End If
            SendToWeb("hub:multi-done", New With {.summary = summary})
            SendToWeb("multi:review-summary", BuildMultiSummaryPayload())
        End Sub

        ' === hub:multi-export ===
        ' payload: { key, excelMode }
        Private Sub HandleMultiExport(payload As Object)
            Dim keyObj As Object = GetProp(payload, "key")
            Dim key As String = NormalizeEventName(Convert.ToString(keyObj))
            Dim excelModeObj As Object = GetProp(payload, "excelMode")
            Dim excelMode As String = NormalizeEventName(Convert.ToString(excelModeObj))
            Dim doAutoFit As Boolean = ParseExcelMode(payload)

            Try
                Select Case If(key, "").ToLowerInvariant()
                    Case "connector"
                        ExportConnector(doAutoFit, excelMode)
                    Case "floorinfo"
                        ExportFloorInfo(doAutoFit, excelMode)
                    Case "pms"
                        ExportSegmentPms(doAutoFit, excelMode)
                    Case "guid"
                        ExportGuid(excelMode)
                    Case "familylink"
                        ExportFamilyLink(doAutoFit, excelMode)
                    Case "points"
                        ExportPoints(doAutoFit, excelMode)
                    Case "linkworkset"
                        ExportLinkWorkset(doAutoFit, excelMode)
                    Case Else
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "알 수 없는 기능 키입니다."})
                        Return
                End Select
            Catch ex As Exception
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = ex.Message})
                Return
            End Try
        End Sub

        ' [추가] 저장된 파일에 기능별(키별) 스타일 적용 (등록된 키만 적용됨)
        Private Sub TryApplyExportStyles(exportKey As String, savedPath As String, Optional doAutoFit As Boolean = True, Optional excelMode As String = "normal")
            If String.IsNullOrWhiteSpace(exportKey) Then Exit Sub
            If String.IsNullOrWhiteSpace(savedPath) Then Exit Sub
            Try
                Global.KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey(exportKey, savedPath, autoFit:=doAutoFit, excelMode:=excelMode)
            Catch
                ' 스타일 실패해도 저장 성공은 유지
            End Try
        End Sub

        Private Sub ExportConnector(doAutoFit As Boolean, excelMode As String)
            Dim allRows = If(_multiConnectorRows, New List(Of Dictionary(Of String, Object))())

            ' 파일 목록(선택 순서 유지)
            Dim fileList As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each p In _multiRequest.RvtPaths
                    Dim path As String = TryCast(p, String)
                    Dim name As String = ""
                    Try
                        name = System.IO.Path.GetFileName(path)
                    Catch
                        name = ""
                    End Try
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If seen.Add(name) Then fileList.Add(name)
                Next
            End If

            ' 선택 파일 목록이 없다면, rows의 File 컬럼에서 추정(순서 보존)
            If fileList.Count = 0 Then
                For Each r In allRows
                    Dim f As String = ""
                    Try
                        If r IsNot Nothing AndAlso r.ContainsKey("File") AndAlso r("File") IsNot Nothing Then
                            f = r("File").ToString()
                        End If
                    Catch
                        f = ""
                    End Try
                    If String.IsNullOrWhiteSpace(f) Then Continue For
                    If seen.Add(f) Then fileList.Add(f)
                Next
            End If

            If allRows.Count = 0 AndAlso fileList.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "커넥터 결과가 없습니다."})
                Return
            End If

            Dim extras = If(_multiConnectorExtras, New List(Of String)())

            Dim rawUnit As String = Nothing
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Connector IsNot Nothing Then
                rawUnit = _multiRequest.Connector.Unit
            End If
            If String.IsNullOrWhiteSpace(rawUnit) Then rawUnit = _lastConnectorUnit
            Dim uiUnit As String = NormalizeUiUnit(rawUnit)

            Dim excludeEndDummy As Boolean = _lastConnectorExcludeEndDummy
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                excludeEndDummy = _multiRequest.Common.ExcludeEndDummy
            End If

            ' ✅ 멀티 파라미터 목록 파싱(검토했으나 이슈 0건인 파라미터 안내행 출력용)
            Dim reviewParams As List(Of String) = Nothing
            Try
                Dim rawParamCsv As String = Nothing
                If _multiRequest IsNot Nothing AndAlso _multiRequest.Connector IsNot Nothing Then
                    rawParamCsv = _multiRequest.Connector.Param
                End If
                If String.IsNullOrWhiteSpace(rawParamCsv) Then rawParamCsv = _lastConnectorParam
                reviewParams = ParseReviewParamsCsv(rawParamCsv)
            Catch
                reviewParams = Nothing
            End Try
            If reviewParams Is Nothing Then reviewParams = New List(Of String)()

            ' ✅ 커넥터는 "이슈 항목만" 내보내는 정책 유지
            Dim issueRows As List(Of Dictionary(Of String, Object)) = allRows.Where(Function(r) ShouldExportIssueRow(r)).ToList()
            If excludeEndDummy Then
                issueRows = issueRows.Where(Function(r) Not ShouldExcludeEndDummyRow(r)).ToList()
            End If

            Dim headers As List(Of String) = BuildConnectorHeaders(extras, uiUnit)
            HostLog("debug", "[multi][connector] export headers => " & String.Join(" | ", headers))
            SendToWeb("host:info", New With {.message = "[multi][connector] export headers => " & String.Join(" | ", headers)})

            Dim exportCount As Integer = If(issueRows Is Nothing, 0, issueRows.Count)

            ' 기본 파일명(기존 규칙 유지 + 멀티파일이면 Selected n Files 규칙 반영)
            Dim baseRvtName As String = ""
            Try
                If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing AndAlso _multiRequest.RvtPaths.Count > 0 Then
                    Dim firstPath As String = TryCast(_multiRequest.RvtPaths(0), String)
                    If Not String.IsNullOrWhiteSpace(firstPath) Then
                        baseRvtName = System.IO.Path.GetFileNameWithoutExtension(firstPath)
                    End If
                End If
            Catch
                baseRvtName = ""
            End Try

            Dim defaultFileName As String = BuildTradeReviewDefaultExcelName(baseRvtName, exportCount)

            ' ✅ 2개 이상 선택 시: [첫번째 파일 규칙 prefix]+nFile_공종검토 / 규칙 불일치: Parameter 연속성검토_Selected n Files
            Try
                If fileList IsNot Nothing AndAlso fileList.Count >= 2 Then
                    Dim firstBase As String = System.IO.Path.GetFileNameWithoutExtension(fileList(0))
                    Dim prefix As String = ExtractTradePrefix(firstBase)
                    If Not String.IsNullOrWhiteSpace(prefix) Then
                        Dim addN As Integer = Math.Max(0, fileList.Count - 1)
                        defaultFileName = $"{prefix}+{addN}File_공종검토.xlsx"
                    Else
                        defaultFileName = $"Parameter 연속성검토_Selected {fileList.Count} Files.xlsx"
                    End If
                    defaultFileName = SanitizeFileName(defaultFileName)
                    If Not defaultFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then defaultFileName &= ".xlsx"
                End If
            Catch
                ' ignore
            End Try

            If String.IsNullOrWhiteSpace(defaultFileName) Then
                defaultFileName = $"Connector_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Dim saved As String = ""

            If fileList IsNot Nothing AndAlso fileList.Count >= 2 Then
                ' ✅ 파일별 시트 분리 저장
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()

                For Each fileName In fileList
                    Dim baseName As String = ""
                    Try
                        baseName = System.IO.Path.GetFileNameWithoutExtension(fileName)
                    Catch
                        baseName = fileName
                    End Try
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = fileName

                    Dim rowsForFile As List(Of Dictionary(Of String, Object)) =
                        issueRows.Where(Function(r)
                                            If r Is Nothing Then Return False
                                            Dim rf As String = ""
                                            Try
                                                If r.ContainsKey("File") AndAlso r("File") IsNot Nothing Then rf = r("File").ToString()
                                            Catch
                                                rf = ""
                                            End Try
                                            If String.IsNullOrWhiteSpace(rf) Then Return False

                                            Dim rfBase As String = rf
                                            Try
                                                rfBase = System.IO.Path.GetFileNameWithoutExtension(rf)
                                            Catch
                                                rfBase = rf
                                            End Try

                                            Return String.Equals(rf, fileName, StringComparison.OrdinalIgnoreCase) _
                                                OrElse String.Equals(rf, baseName, StringComparison.OrdinalIgnoreCase) _
                                                OrElse String.Equals(rfBase, baseName, StringComparison.OrdinalIgnoreCase)
                                        End Function).ToList()

                    ' ✅ 선택한 파라미터 중 이슈 0건인 항목도 검토 여부를 알 수 있도록 안내행 추가
                    If reviewParams IsNot Nothing AndAlso reviewParams.Count > 0 Then
                        Dim msgRows = BuildNoIssueMessageRows(rowsForFile, reviewParams)
                        If msgRows IsNot Nothing AndAlso msgRows.Count > 0 Then
                            For Each mr In msgRows
                                If mr IsNot Nothing Then mr("File") = fileName
                            Next
                            rowsForFile = msgRows.Concat(rowsForFile).ToList()
                        End If
                    End If

                    Dim table = BuildConnectorTableFromRows(headers, rowsForFile)
                    ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
                    If table.Rows.Count > 0 AndAlso Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: 커넥터")
                    sheetList.Add(New KeyValuePair(Of String, DataTable)(baseName, table))
                Next

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="connector", exportKind:="connector")
            Else
                ' 단일 파일
                Dim rowsForSingle As List(Of Dictionary(Of String, Object)) = issueRows

                ' ✅ 선택한 파라미터 중 이슈 0건인 항목도 안내행 추가(멀티와 동일)
                If reviewParams IsNot Nothing AndAlso reviewParams.Count > 0 Then
                    Dim msgRows = BuildNoIssueMessageRows(rowsForSingle, reviewParams)
                    If msgRows IsNot Nothing AndAlso msgRows.Count > 0 Then
                        Dim singleName As String = ""
                        Try
                            If fileList IsNot Nothing AndAlso fileList.Count = 1 Then singleName = fileList(0)
                        Catch
                            singleName = ""
                        End Try
                        For Each mr In msgRows
                            If mr Is Nothing Then Continue For
                            Try
                                If (Not mr.ContainsKey("File")) OrElse mr("File") Is Nothing OrElse String.IsNullOrWhiteSpace(mr("File").ToString()) Then
                                    mr("File") = singleName
                                End If
                            Catch
                            End Try
                        Next
                        rowsForSingle = msgRows.Concat(rowsForSingle).ToList()
                    End If
                End If

                Dim table = BuildConnectorTableFromRows(headers, rowsForSingle)
                ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
                If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: 커넥터")
                saved = ExcelCore.PickAndSaveXlsx("Connector Diagnostics", table, defaultFileName, doAutoFit, "hub:multi-progress", "connector")
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("connector", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub


        Private Sub ExportSegmentPms(doAutoFit As Boolean, excelMode As String)
            Dim classRows = If(_multiPmsClassRows, New List(Of Dictionary(Of String, Object))())
            Dim sizeRows = If(_multiPmsSizeRows, New List(Of Dictionary(Of String, Object))())
            Dim routingRows = If(_multiPmsRoutingRows, New List(Of Dictionary(Of String, Object))())
            Dim totalRowsCount As Integer = classRows.Count + sizeRows.Count + routingRows.Count
            Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
            Dim classHeaders = New List(Of String) From {"File", "PipeType", "Segment", "Class검토결과"}
            Dim sizeHeaders = New List(Of String) From {"FileName", "PipeType", "RevitSegment", "PMSCompared", "ND", "ID", "OD", "PMS ND", "PMS ID", "PMS OD", "Result"}
            Dim routingHeaders = New List(Of String) From {"File", "PipeType", "Part", "Type", "Class검토"}

            Dim classTable = BuildTableFromRows(classHeaders, classRows)
            Dim sizeTable = BuildTableFromRows(sizeHeaders, sizeRows)
            Dim routingTable = BuildTableFromRows(routingHeaders, routingRows)
            ExcelCore.EnsureNoDataRow(classTable, "오류가 없습니다.")
            ExcelCore.EnsureNoDataRow(sizeTable, "오류가 없습니다.")
            ExcelCore.EnsureNoDataRow(routingTable, "오류가 없습니다.")

            If totalRowsCount = 0 Then
                AddEmptyMessageRow(classTable)
                AddEmptyMessageRow(sizeTable)
                AddEmptyMessageRow(routingTable)
            End If

            If classTable.Rows.Count > 0 AndAlso Not ValidateSchema(classTable, classHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Class")
            If sizeTable.Rows.Count > 0 AndAlso Not ValidateSchema(sizeTable, sizeHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Size")
            If routingTable.Rows.Count > 0 AndAlso Not ValidateSchema(routingTable, routingHeaders) Then Throw New InvalidOperationException("스키마 검증 실패: PMS Routing")

            If totalRowsCount = 0 Then
                sheetList.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", classTable))
                sheetList.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", sizeTable))
                sheetList.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", routingTable))
            Else
                If classTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", classTable))
                If sizeTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", sizeTable))
                If routingTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", routingTable))
            End If

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, $"SegmentPms_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                ' 등록된 스타일 키가 있으면 적용(현재는 보통 no-op)
                TryApplyExportStyles("pms", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportGuid(excelMode As String)
            Dim doAutoFit As Boolean = String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase)
            If _multiGuidProject Is Nothing Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "GUID 결과가 없습니다."})
                Return
            End If
            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            sheets.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", GuidAuditService.PrepareExportTable(_multiGuidProject, 1)))
            If _multiGuidFamilyDetail IsNot Nothing Then
                sheets.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", GuidAuditService.PrepareExportTable(_multiGuidFamilyDetail, 2)))
            End If
            Dim saved = GuidAuditService.ExportMulti(sheets, excelMode, "hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("guid", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportFamilyLink(doAutoFit As Boolean, excelMode As String)
            Dim rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)())
            ExcelProgressReporter.Reset("hub:multi-progress")
            Dim saved = FamilyLinkAuditExport.Export(rows,
                                                     fastExport:=String.Equals(excelMode, "fast", StringComparison.OrdinalIgnoreCase),
                                                     autoFit:=doAutoFit,
                                                     progressChannel:="hub:multi-progress")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("familylink", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportPoints(doAutoFit As Boolean, excelMode As String)
            Dim pointRows = If(_multiPointRows, New List(Of ExportPointsService.Row)())
            Dim unit As String = "ft"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Points IsNot Nothing Then
                unit = _multiRequest.Points.Unit
            End If
            Dim headers = BuildPointHeaders(unit)
            Dim rows As New List(Of Dictionary(Of String, Object))()
            For Each r In pointRows
                rows.Add(New Dictionary(Of String, Object) From {
                    {"File", r.File},
                    {"ProjectPoint_E", ConvertPoint(r.ProjectE, unit)},
                    {"ProjectPoint_N", ConvertPoint(r.ProjectN, unit)},
                    {"ProjectPoint_Z", ConvertPoint(r.ProjectZ, unit)},
                    {"SurveyPoint_E", ConvertPoint(r.SurveyE, unit)},
                    {"SurveyPoint_N", ConvertPoint(r.SurveyN, unit)},
                    {"SurveyPoint_Z", ConvertPoint(r.SurveyZ, unit)},
                    {"TrueNorthAngle", Math.Round(r.TrueNorth, 3)}
                })
            Next
            Dim table = BuildPointTable(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: Points")
            Dim saved = ExcelCore.PickAndSaveXlsx("Points", table, $"Points_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", "points")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("points", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportLinkWorkset(doAutoFit As Boolean, excelMode As String)
            Dim rows = If(_multiLinkWorksetRows, New List(Of LinkWorksetAuditRow)())
            If rows.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "링크 기본 웍셋 결과가 없습니다."})
                Return
            End If

            Dim headers As New List(Of String) From {
                "HostFile",
                "LinkName",
                "AttachmentType",
                "WasLoadedBefore",
                "IsLoadedAfter",
                "IsWorkshared",
                "DefaultWorkset",
                "TotalUserWorksets",
                "OpenUserWorksetsBefore",
                "DefaultOnlyBefore",
                "OpenUserWorksetsAfter",
                "DefaultOnlyAfter",
                "ApplyRequested",
                "Applied",
                "Status",
                "Message"
            }

            Dim table = BuildLinkWorksetTable(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: LinkWorkset")
            Dim saved = ExcelCore.PickAndSaveXlsx("LinkWorkset", table, $"LinkWorkset_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", "linkworkset")
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                TryApplyExportStyles("linkworkset", saved, doAutoFit, If(excelMode, "normal"))
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Shared Sub ResetMultiCaches()
            _multiConnectorRows = Nothing
            _multiConnectorExtras = Nothing
            _multiFloorInfoRows = Nothing
            _multiFloorInfoFileSummaries = Nothing
            _multiFloorInfoWarnings = Nothing
            _multiPmsClassRows = Nothing
            _multiPmsSizeRows = Nothing
            _multiPmsRoutingRows = Nothing
            _multiGuidProject = Nothing
            _multiGuidFamilyDetail = Nothing
            _multiGuidFamilyIndex = Nothing
            _multiFamilyLinkRows = Nothing
            _multiPointRows = Nothing
            _multiLinkWorksetRows = Nothing
        End Sub

        Private Function ParseMultiRequest(payload As Object) As MultiRunRequest
            Dim req As New MultiRunRequest()
            Dim pd = ToDict(payload)
            req.RvtPaths = ExtractStringList(pd, "rvtPaths")
            req.UseActiveDocument = ToBool(GetDictValue(pd, "useActiveDocument"))
            If Not req.UseActiveDocument Then
                req.UseActiveDocument = ToBool(GetDictValue(pd, "useActiveDoc"))
            End If

            Dim commonObj As Object = Nothing
            If pd.TryGetValue("commonOptions", commonObj) Then
                Dim commonDict = ToDict(commonObj)
                req.Common.ExtraParams = SafeStr(GetDictValue(commonDict, "extraParams"))
                req.Common.TargetFilter = SafeStr(GetDictValue(commonDict, "targetFilter"))
                req.Common.ExcludeEndDummy = ToBool(GetDictValue(commonDict, "excludeEndDummy"))
            End If

            Dim featuresObj As Object = Nothing
            If pd.TryGetValue("features", featuresObj) Then
                Dim fd = ToDict(featuresObj)
                req.Connector = ParseConnector(fd)
                req.FloorInfo = ParseFloorInfo(fd)
                req.Pms = ParsePms(fd)
                req.Guid = ParseGuid(fd)
                req.FamilyLink = ParseFamilyLink(fd)
                req.Points = ParsePoints(fd)
                req.LinkWorkset = ParseLinkWorkset(fd)
            End If
            Return req
        End Function

        Private Function ParseConnector(fd As Dictionary(Of String, Object)) As MultiConnectorOptions
            Dim opt As New MultiConnectorOptions()
            Dim obj = GetDictValue(fd, "connector")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Tol = ToDouble(GetDictValue(d, "tol"), 1.0R)
            opt.Unit = SafeStr(GetDictValue(d, "unit"))
            opt.Param = SafeStr(GetDictValue(d, "param"))
            opt.IncludePointXY = ToBool(GetDictValue(d, "includePointXY"))
            opt.IncludeLinearMetrics = ToBool(GetDictValue(d, "includeLinearMetrics"))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "inch"
            If String.IsNullOrWhiteSpace(opt.Param) Then opt.Param = "Comments"
            Return opt
        End Function

        Private Function ParsePms(fd As Dictionary(Of String, Object)) As MultiPmsOptions
            Dim opt As New MultiPmsOptions()
            Dim obj = GetDictValue(fd, "pms")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.NdRound = ToInt(GetDictValue(d, "ndRound"), 3)
            opt.TolMm = ToDouble(GetDictValue(d, "tolMm"), 0.01R)
            opt.ClassMatch = ToBool(GetDictValue(d, "classMatch"))
            Return opt
        End Function

        Private Function ParseGuid(fd As Dictionary(Of String, Object)) As MultiGuidOptions
            Dim opt As New MultiGuidOptions()
            Dim obj = GetDictValue(fd, "guid")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.IncludeFamily = ToBool(GetDictValue(d, "includeFamily"))
            opt.IncludeAnnotation = ToBool(GetDictValue(d, "includeAnnotation"))
            Return opt
        End Function

        Private Function ParseFamilyLink(fd As Dictionary(Of String, Object)) As MultiFamilyLinkOptions
            Dim opt As New MultiFamilyLinkOptions()
            Dim obj = GetDictValue(fd, "familylink")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            Dim rawTargets = GetDictValue(d, "targets")
            Dim targets As New List(Of FamilyLinkTargetParam)()
            Dim arr = TryCast(rawTargets, System.Collections.IEnumerable)
            If arr IsNot Nothing AndAlso Not TypeOf rawTargets Is String Then
                For Each o In arr
                    Dim td = ToDict(o)
                    Dim name = SafeStr(GetDictValue(td, "name"))
                    Dim guidStr = SafeStr(GetDictValue(td, "guid"))
                    Dim g As Guid
                    If Guid.TryParse(guidStr, g) Then
                        targets.Add(New FamilyLinkTargetParam With {.Name = name, .Guid = g})
                    End If
                Next
            End If
            opt.Targets = targets
            Return opt
        End Function

        Private Function ParsePoints(fd As Dictionary(Of String, Object)) As MultiPointsOptions
            Dim opt As New MultiPointsOptions()
            Dim obj = GetDictValue(fd, "points")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Unit = SafeStr(GetDictValue(d, "unit"))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "ft"
            Return opt
        End Function

        Private Function ParseLinkWorkset(fd As Dictionary(Of String, Object)) As MultiLinkWorksetOptions
            Dim opt As New MultiLinkWorksetOptions()
            Dim obj = GetDictValue(fd, "linkworkset")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.ApplyDefaultWorksetOnly = ToBool(GetDictValue(d, "applyDefaultWorksetOnly"), True)
            opt.UseSyncComment = ToBool(GetDictValue(d, "useSyncComment"), False)
            opt.SyncComment = SafeStr(GetDictValue(d, "syncComment")).Trim()
            Return opt
        End Function

        Private Shared Function AnyFeatureEnabled(req As MultiRunRequest) As Boolean
            If req Is Nothing Then Return False
            Return req.Connector.Enabled OrElse req.FloorInfo.Enabled OrElse req.Pms.Enabled OrElse req.Guid.Enabled OrElse req.FamilyLink.Enabled OrElse req.Points.Enabled OrElse req.LinkWorkset.Enabled
        End Function

        Private Shared Function CountEnabledFeatures(req As MultiRunRequest) As Integer
            If req Is Nothing Then Return 0
            Dim count As Integer = 0
            If req.Connector.Enabled Then count += 1
            If req.FloorInfo.Enabled Then count += 1
            If req.Pms.Enabled Then count += 1
            If req.Guid.Enabled Then count += 1
            If req.FamilyLink.Enabled Then count += 1
            If req.Points.Enabled Then count += 1
            If req.LinkWorkset.Enabled Then count += 1
            Return Math.Max(count, 1)
        End Function

        Private Sub AppendMultiRunItem(fileName As String, status As String, reason As String, phase As String, started As DateTime)
            If _multiRunItems Is Nothing Then _multiRunItems = New List(Of MultiRunItem)()
            Dim elapsed = CLng((Date.Now - started).TotalMilliseconds)
            Dim displayName As String = GetSafeMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(displayName) Then displayName = SafeStr(fileName)
            _multiRunItems.Add(New MultiRunItem With {
                .File = displayName,
                .Status = status,
                .Reason = reason,
                .Phase = phase,
                .ElapsedMs = elapsed
            })
        End Sub

        Private Function BuildMultiSummaryPayload() As Object
            Dim items As List(Of MultiRunItem) = If(_multiRunItems, New List(Of MultiRunItem)())
            Dim featureSummaries = BuildMultiFeatureSummaries()
            Dim itemPayloads = items.
                Where(Function(x) x IsNot Nothing).
                Select(Function(x) New With {
                    .file = GetSafeMultiFileName(If(x.File, "")),
                    .status = SafeStr(If(x.Status, "")),
                    .reason = SafeStr(If(x.Reason, "")),
                    .phase = SafeStr(If(x.Phase, "")),
                    .elapsedMs = x.ElapsedMs
                }).
                ToList()

            Dim total As Integer = If(_multiTotal > 0, _multiTotal, items.Count)

            Dim success As Integer = items.Where(Function(x) String.Equals(x.Status, "success", StringComparison.OrdinalIgnoreCase)).Count()
            Dim skipped As Integer = items.Where(Function(x) String.Equals(x.Status, "skipped", StringComparison.OrdinalIgnoreCase)).Count()
            Dim failed As Integer = items.Where(Function(x) String.Equals(x.Status, "failed", StringComparison.OrdinalIgnoreCase)).Count()

            Return New With {
        .ok = True,
        .mode = "multiRvt",
        .featureId = "multi_rvt_batch",
        .title = "다중 RVT 검토",
        .finishedAt = Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
        .total = total,
        .success = success,
        .skipped = skipped,
        .failed = failed,
        .canceled = False,
        .featureSummaries = featureSummaries,
        .items = itemPayloads
    }
        End Function

        Private Function BuildMultiFeatureSummaries() As Dictionary(Of String, Object)
            Dim summaries As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If _multiRequest Is Nothing Then Return summaries

            If _multiRequest.Connector IsNot Nothing AndAlso _multiRequest.Connector.Enabled Then
                summaries("connector") = BuildConnectorMultiSummary()
            End If

            If _multiRequest.FloorInfo IsNot Nothing AndAlso _multiRequest.FloorInfo.Enabled Then
                summaries("floorinfo") = BuildFloorInfoMultiSummary()
            End If

            If _multiRequest.Guid IsNot Nothing AndAlso _multiRequest.Guid.Enabled Then
                summaries("guid") = BuildGuidMultiSummary()
            End If

            If _multiRequest.FamilyLink IsNot Nothing AndAlso _multiRequest.FamilyLink.Enabled Then
                summaries("familylink") = BuildFamilyLinkMultiSummary()
            End If

            If _multiRequest.Points IsNot Nothing AndAlso _multiRequest.Points.Enabled Then
                summaries("points") = BuildPointsMultiSummary()
            End If

            If _multiRequest.LinkWorkset IsNot Nothing AndAlso _multiRequest.LinkWorkset.Enabled Then
                summaries("linkworkset") = BuildLinkWorksetMultiSummary()
            End If

            Return summaries
        End Function

        Private Function BuildConnectorMultiSummary() As Object
            Dim rows = If(_multiConnectorRows, New List(Of Dictionary(Of String, Object))())
            Dim issueCount As Integer = rows.Where(Function(r) ShouldExportIssueRow(r)).Count()
            Dim mismatchCount As Integer = rows.Where(Function(r) IsMismatchRow(r)).Count()
            Dim nearCount As Integer = rows.Where(Function(r) IsZeroDistanceNotConnected(r)).Count()
            Dim errorCount As Integer = rows.Where(Function(r) String.Equals(ReadField(r, "Status"), "ERROR", StringComparison.OrdinalIgnoreCase)).Count()
            Dim normalCount As Integer = Math.Max(rows.Count - issueCount, 0)
            Dim fileCount As Integer = GetRequestedMultiFileCount()
            Dim fileSummaries = BuildConnectorFileSummaries(rows)

            Return New With {
                .key = "connector",
                .label = "파라미터 연속성 검토",
                .lines = New String() {
                    $"선택 파일 수: {fileCount}개",
                    $"전체 결과 건수: {rows.Count}건",
                    $"오류/불일치 건수: {errorCount + mismatchCount}건",
                    $"연결 필요 건수: {nearCount}건",
                    $"정상 건수: {normalCount}건",
                    $"엑셀 내보내기 대상: {issueCount}건"
                },
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildConnectorFileSummaries(rows As IList(Of Dictionary(Of String, Object))) As List(Of Object)
            Dim sourceRows = If(rows, New List(Of Dictionary(Of String, Object))())
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

            For Each row In sourceRows
                Dim safeName As String = GetSafeMultiFileName(ReadField(row, "File"))
                If String.IsNullOrWhiteSpace(safeName) Then Continue For
                If seen.Add(safeName) Then orderedNames.Add(safeName)
            Next

            Dim result As New List(Of Object)()
            For Each fileName In orderedNames
                Dim total As Integer = 0
                Dim issueCount As Integer = 0
                Dim nearCount As Integer = 0
                Dim statusText As String = "pending"

                For Each row In sourceRows
                    Dim rowFile As String = GetSafeMultiFileName(ReadField(row, "File"))
                    If Not String.Equals(rowFile, fileName, StringComparison.OrdinalIgnoreCase) Then Continue For
                    total += 1
                    If ShouldExportIssueRow(row) Then issueCount += 1
                    If IsZeroDistanceNotConnected(row) Then nearCount += 1
                Next

                If _multiRunItems IsNot Nothing Then
                    For Each item In _multiRunItems
                        If item Is Nothing Then Continue For
                        Dim itemFile As String = GetSafeMultiFileName(item.File)
                        If String.Equals(itemFile, fileName, StringComparison.OrdinalIgnoreCase) Then
                            statusText = If(String.IsNullOrWhiteSpace(item.Status), "pending", item.Status)
                            Exit For
                        End If
                    Next
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = total,
                    .issues = issueCount,
                    .near = nearCount,
                    .status = statusText
                })
            Next

            Return result
        End Function

        Private Function BuildGuidMultiSummary() As Object
            Dim projectRows As Integer = If(_multiGuidProject, New DataTable()).Rows.Count
            Dim familyRows As Integer = If(_multiGuidFamilyDetail, New DataTable()).Rows.Count
            Dim includeFamily As Boolean = (_multiRequest IsNot Nothing AndAlso _multiRequest.Guid IsNot Nothing AndAlso _multiRequest.Guid.IncludeFamily)
            Dim fileSummaries = BuildGuidFileSummaries()
            Dim lines As New List(Of String) From {
                $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                $"프로젝트 결과 행 수: {projectRows}행"
            }
            If includeFamily Then
                lines.Add($"패밀리 결과 행 수: {familyRows}행")
            End If
            lines.Add($"엑셀 시트 수: {If(includeFamily, 2, 1)}개")

            Return New With {
                .key = "guid",
                .label = "공유파라미터 GUID 검토",
                .lines = lines.ToArray(),
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildFamilyLinkMultiSummary() As Object
            Dim rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)())
            Dim errorCount As Integer = rows.Where(Function(r) r IsNot Nothing AndAlso String.Equals(If(r.Issue, ""), FamilyLinkAuditIssue.[Error].ToString(), StringComparison.OrdinalIgnoreCase)).Count()
            Dim targetCount As Integer = 0
            Dim fileSummaries = BuildFamilyLinkFileSummaries(rows)
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FamilyLink IsNot Nothing AndAlso _multiRequest.FamilyLink.Targets IsNot Nothing Then
                targetCount = _multiRequest.FamilyLink.Targets.Count
            End If

            Return New With {
                .key = "familylink",
                .label = "패밀리 공유파라미터 연동 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"검토 대상 파라미터 수: {targetCount}개",
                    $"이슈 결과 행 수: {rows.Count}행",
                    $"오류 행 수: {errorCount}행"
                },
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildPointsMultiSummary() As Object
            Dim rows = If(_multiPointRows, New List(Of ExportPointsService.Row)())
            Dim fileSummaries = BuildPointsFileSummaries(rows)
            Dim successFileSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row In rows
                If row Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(row.File) Then Continue For
                successFileSet.Add(row.File)
            Next
            Dim successFiles As Integer = successFileSet.Count
            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim failedCount As Integer = Math.Max(requestedCount - successFiles, 0)

            Return New With {
                .key = "points",
                .label = "Point 추출",
                .lines = New String() {
                    $"선택 파일 수: {requestedCount}개",
                    $"결과 행 수: {rows.Count}행",
                    $"성공 파일 수: {successFiles}개",
                    $"실패 파일 수: {failedCount}개"
                },
                .fileSummaries = fileSummaries
            }
        End Function

        Private Function BuildLinkWorksetMultiSummary() As Object
            Dim rows = If(_multiLinkWorksetRows, New List(Of LinkWorksetAuditRow)())
            Dim totalLinks As Integer = rows.Count
            Dim appliedCount As Integer = rows.Where(Function(r) r IsNot Nothing AndAlso r.Applied).Count()
            Dim okCount As Integer = rows.Where(Function(r) r IsNot Nothing AndAlso r.DefaultOnlyOpenAfter.HasValue AndAlso r.DefaultOnlyOpenAfter.Value).Count()
            Dim issueCount As Integer = rows.Where(Function(r) r IsNot Nothing AndAlso IsLinkWorksetIssue(r)).Count()
            Dim naCount As Integer = rows.Where(Function(r) r IsNot Nothing AndAlso String.Equals(If(r.Status, ""), "n/a", StringComparison.OrdinalIgnoreCase)).Count()

            Return New With {
                .key = "linkworkset",
                .label = "링크 기본 웍셋 점검/적용",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"링크 결과 행 수: {totalLinks}행",
                    $"기본 workset만 열린 링크 수: {okCount}개",
                    $"재적용된 링크 수: {appliedCount}개",
                    $"확인 필요 링크 수: {issueCount}개",
                    $"비Workshared 링크 수: {naCount}개"
                },
                .fileSummaries = BuildLinkWorksetFileSummaries(rows)
            }
        End Function

        Private Function BuildLinkWorksetFileSummaries(rows As IList(Of LinkWorksetAuditRow)) As List(Of Object)
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

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim safeName As String = GetSafeMultiFileName(row.HostFileName)
                If String.IsNullOrWhiteSpace(safeName) Then Continue For
                If seen.Add(safeName) Then orderedNames.Add(safeName)
            Next

            Dim result As New List(Of Object)()
            For Each fileName In orderedNames
                Dim perFileRows = rows.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(r.HostFileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()

                Dim total As Integer = perFileRows.Count
                Dim issueCount As Integer = perFileRows.Where(Function(r) IsLinkWorksetIssue(r)).Count()
                Dim appliedCount As Integer = perFileRows.Where(Function(r) r IsNot Nothing AndAlso r.Applied).Count()
                Dim reason As String = ""
                If total = 0 Then
                    reason = "링크 없음"
                Else
                    reason = $"적용 {appliedCount}건 / 확인필요 {issueCount}건"
                End If

                Dim statusText As String = "pending"
                If _multiRunItems IsNot Nothing Then
                    For Each item In _multiRunItems
                        If item Is Nothing Then Continue For
                        If String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase) Then
                            statusText = If(String.IsNullOrWhiteSpace(item.Status), "pending", item.Status)
                            Exit For
                        End If
                    Next
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = total,
                    .issues = issueCount,
                    .near = appliedCount,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function BuildGuidFileSummaries() As List(Of Object)
            Dim projectTable = If(_multiGuidProject, New DataTable())
            Dim familyTable = If(_multiGuidFamilyDetail, New DataTable())
            Dim orderedNames = BuildOrderedMultiFileNames(
                projectTable.Rows.Cast(Of DataRow)().Select(Function(r) ReadDataRowField(r, "RvtName")),
                familyTable.Rows.Cast(Of DataRow)().Select(Function(r) ReadDataRowField(r, "RvtName")))

            Dim result As New List(Of Object)()
            For Each fileName In orderedNames
                Dim projectRows = projectTable.Rows.Cast(Of DataRow)().
                    Where(Function(r) String.Equals(GetSafeMultiFileName(ReadDataRowField(r, "RvtName")), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim familyRows = familyTable.Rows.Cast(Of DataRow)().
                    Where(Function(r) String.Equals(GetSafeMultiFileName(ReadDataRowField(r, "RvtName")), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()

                Dim total As Integer = projectRows.Count + familyRows.Count
                Dim issueCount As Integer = projectRows.Where(Function(r) IsGuidIssueResult(ReadDataRowField(r, "Result"))).Count() +
                                          familyRows.Where(Function(r) IsGuidIssueResult(ReadDataRowField(r, "Result"))).Count()
                Dim runReason As String = ""
                Dim statusText As String = GetMultiRunItemStatus(fileName, runReason)
                Dim reasonParts As New List(Of String)()
                If projectRows.Count > 0 Then reasonParts.Add($"Project {projectRows.Count}행")
                If familyRows.Count > 0 Then reasonParts.Add($"Family {familyRows.Count}행")
                Dim reason As String = If(reasonParts.Count > 0, String.Join(" / ", reasonParts), runReason)

                result.Add(New With {
                    .file = fileName,
                    .total = total,
                    .issues = issueCount,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function BuildFamilyLinkFileSummaries(rows As IList(Of FamilyLinkAuditRow)) As List(Of Object)
            Dim sourceRows = If(rows, New List(Of FamilyLinkAuditRow)())
            Dim orderedNames = BuildOrderedMultiFileNames(sourceRows.Select(Function(r) If(r Is Nothing, "", r.FileName)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim perFileRows = sourceRows.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(r.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim issueCount As Integer = perFileRows.Where(Function(r) Not String.Equals(If(r.Issue, ""), FamilyLinkAuditIssue.OK.ToString(), StringComparison.OrdinalIgnoreCase)).Count()
                Dim runReason As String = ""
                Dim statusText As String = GetMultiRunItemStatus(fileName, runReason)
                Dim reason As String = If(perFileRows.Count > 0, $"이슈 {issueCount}건", runReason)

                result.Add(New With {
                    .file = fileName,
                    .total = perFileRows.Count,
                    .issues = issueCount,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function BuildPointsFileSummaries(rows As IList(Of ExportPointsService.Row)) As List(Of Object)
            Dim sourceRows = If(rows, New List(Of ExportPointsService.Row)())
            Dim orderedNames = BuildOrderedMultiFileNames(sourceRows.Select(Function(r) If(r Is Nothing, "", r.File)))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim perFileRows = sourceRows.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(r.File), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim runReason As String = ""
                Dim statusText As String = GetMultiRunItemStatus(fileName, runReason)
                Dim issueCount As Integer = If(perFileRows.Count = 0 AndAlso Not String.Equals(statusText, "success", StringComparison.OrdinalIgnoreCase), 1, 0)
                Dim reason As String = runReason
                If String.IsNullOrWhiteSpace(reason) Then
                    reason = If(perFileRows.Count > 0, $"포인트 {perFileRows.Count}건", "추출 결과 없음")
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = perFileRows.Count,
                    .issues = issueCount,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function BuildOrderedMultiFileNames(ParamArray nameGroups() As IEnumerable(Of String)) As List(Of String)
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

            For Each group In nameGroups
                If group Is Nothing Then Continue For
                For Each rawName In group
                    Dim safeName As String = GetSafeMultiFileName(rawName)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            Next

            Return orderedNames
        End Function

        Private Function GetMultiRunItemStatus(fileName As String, ByRef reason As String) As String
            reason = ""
            If _multiRunItems IsNot Nothing Then
                For Each item In _multiRunItems
                    If item Is Nothing Then Continue For
                    If String.Equals(GetSafeMultiFileName(item.File), fileName, StringComparison.OrdinalIgnoreCase) Then
                        reason = If(item.Reason, "")
                        Return If(String.IsNullOrWhiteSpace(item.Status), "pending", item.Status)
                    End If
                Next
            End If
            Return "pending"
        End Function

        Private Shared Function ReadDataRowField(row As DataRow, columnName As String) As String
            If row Is Nothing OrElse row.Table Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return ""
            If Not row.Table.Columns.Contains(columnName) Then Return ""
            Try
                Return SafeStr(row(columnName))
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function IsGuidIssueResult(resultText As String) As Boolean
            Dim value As String = SafeStr(resultText).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return False
            If value.StartsWith("OK", StringComparison.OrdinalIgnoreCase) Then Return False
            If String.Equals(value, "PROJECT_PARAM", StringComparison.OrdinalIgnoreCase) Then Return False
            Return True
        End Function

        Private Shared Function IsLinkWorksetIssue(row As LinkWorksetAuditRow) As Boolean
            If row Is Nothing Then Return False
            If String.Equals(If(row.Status, ""), "error", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(If(row.Status, ""), "warning", StringComparison.OrdinalIgnoreCase) Then Return True
            If row.IsWorkshared AndAlso row.DefaultOnlyOpenAfter.HasValue AndAlso row.DefaultOnlyOpenAfter.Value = False Then Return True
            If row.IsWorkshared AndAlso Not row.DefaultOnlyOpenAfter.HasValue AndAlso row.ApplyRequested Then Return True
            Return False
        End Function

        Private Sub PublishLinkWorksetDiagnostics(rows As IEnumerable(Of LinkWorksetAuditRow), safeName As String)
            If rows Is Nothing Then Return
            For Each row In rows
                If row Is Nothing Then Continue For
                Dim summary As String =
                    $"[linkworkset] {safeName} | {SafeStr(row.LinkName)} | status={SafeStr(row.Status)} | applied={BoolText(row.Applied)} | before={NullableBoolText(row.DefaultOnlyOpenBefore)} | after={NullableBoolText(row.DefaultOnlyOpenAfter)}"
                SendToWeb("host:info", New With {.message = summary})
                If Not String.IsNullOrWhiteSpace(row.DiagnosticLog) Then
                    SendToWeb("host:info", New With {.message = "[linkworkset][diag] " & row.DiagnosticLog})
                End If
            Next
        End Sub

        Private Function GetRequestedMultiFileCount() As Integer
            If _multiRequest Is Nothing OrElse _multiRequest.RvtPaths Is Nothing Then Return 0
            Dim fileSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each path In _multiRequest.RvtPaths
                If String.IsNullOrWhiteSpace(path) Then Continue For
                fileSet.Add(GetSafeMultiFileName(path))
            Next
            Return fileSet.Count
        End Function

        Private Shared Function GetSafeMultiFileName(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then Return String.Empty
            path = NormalizeMultiPath(path)
            Try
                Dim name As String = System.IO.Path.GetFileName(path)
                If Not String.IsNullOrWhiteSpace(name) Then
                    Dim withoutExtension As String = ""
                    Try
                        withoutExtension = System.IO.Path.GetFileNameWithoutExtension(name)
                    Catch
                        withoutExtension = ""
                    End Try
                    If Not String.IsNullOrWhiteSpace(withoutExtension) Then Return withoutExtension
                    Return name
                End If
            Catch
            End Try
            Return path
        End Function

        Private Shared Function NormalizeMultiPath(path As String) As String
            Dim value As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            value = New String(value.Where(Function(ch) Not Char.IsControl(ch)).ToArray())
            value = value.Trim(""""c)

            Try
                If value.IndexOf("\u", StringComparison.OrdinalIgnoreCase) >= 0 OrElse value.Contains("\\") Then
                    value = System.Text.RegularExpressions.Regex.Unescape(value)
                End If
            Catch
            End Try

            Do While value.EndsWith("""", StringComparison.Ordinal)
                value = value.Substring(0, value.Length - 1).TrimEnd()
            Loop

            Do While value.StartsWith("""", StringComparison.Ordinal)
                value = value.Substring(1).TrimStart()
            Loop

            Return value
        End Function

        Private Sub ReportMultiProgress(percent As Double, message As String, detail As String)
            SendToWeb("hub:multi-progress", New With {
                .percent = Math.Max(0.0R, Math.Min(100.0R, percent)),
                .message = message,
                .detail = detail,
                .title = "다중 RVT 검토"
            })
        End Sub

        Private Function CalcStepPercent(basePct As Double, stepIndex As Integer, totalSteps As Integer) As Double
            Dim perFile As Double = If(_multiTotal > 0, 1.0R / CDbl(_multiTotal), 1.0R)
            Dim stepShare As Double = perFile / CDbl(Math.Max(totalSteps, 1))
            Dim stepPct As Double = (basePct + (stepShare * CDbl(stepIndex))) * 100.0R
            Return Math.Min(stepPct, 99.9R)
        End Function

        Private Shared Function BuildOpenOptions(projectPath As ModelPath, preferConnectorWorksets As Boolean) As OpenOptions
            Return BuildOpenOptions(projectPath, preferConnectorWorksets, Nothing)
        End Function

        Private Shared Function BuildOpenOptions(projectPath As ModelPath,
                                                 preferConnectorWorksets As Boolean,
                                                 additionalOpenWorksetNames As IEnumerable(Of String)) As OpenOptions
            Dim opt As New OpenOptions()
            Try
                opt.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
            Catch
            End Try

            Try
                Dim ws = BuildSavedWorksetOpenConfiguration(projectPath, additionalOpenWorksetNames)
                opt.SetOpenWorksetsConfiguration(ws)
            Catch
            End Try

            Return opt
        End Function

        Private Shared Function BuildSavedWorksetOpenConfiguration(projectPath As ModelPath,
                                                                   Optional additionalOpenWorksetNames As IEnumerable(Of String) = Nothing) As WorksetConfiguration
            Dim config As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            If projectPath Is Nothing Then Return config

            Try
                Dim previews = WorksharingUtils.GetUserWorksetInfo(projectPath)
                If previews Is Nothing OrElse previews.Count = 0 Then Return config

                Dim extraNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                If additionalOpenWorksetNames IsNot Nothing Then
                    For Each name In additionalOpenWorksetNames
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        extraNames.Add(name.Trim())
                    Next
                End If

                Dim openIds As New List(Of WorksetId)()
                For Each preview In previews
                    If preview Is Nothing Then Continue For
                    Dim shouldOpen As Boolean = IsPreviewMarkedOpen(preview)
                    If Not shouldOpen AndAlso extraNames.Count > 0 Then
                        Try
                            Dim previewName = Convert.ToString(preview.Name)
                            shouldOpen = extraNames.Contains(If(previewName, String.Empty))
                        Catch
                            shouldOpen = False
                        End Try
                    End If
                    If shouldOpen Then
                        openIds.Add(preview.Id)
                    End If
                Next

                If openIds.Count > 0 Then
                    config.Open(openIds)
                End If
            Catch
            End Try

            Return config
        End Function

        Private Shared Function IsPreviewMarkedOpen(preview As WorksetPreview) As Boolean
            If preview Is Nothing Then Return False
            Try
                Dim prop = preview.GetType().GetProperty("IsOpen")
                If prop Is Nothing Then Return False
                Dim raw = prop.GetValue(preview, Nothing)
                If TypeOf raw Is Boolean Then Return CBool(raw)
            Catch
            End Try
            Return False
        End Function

        Private Shared Function TryExtractMultiBasicFileInfo(pathText As String) As BasicFileInfo
            Try
                If String.IsNullOrWhiteSpace(pathText) Then Return Nothing
                Return BasicFileInfo.Extract(pathText)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function CreateMultiNewLocalPath(centralPath As String) As String
            Dim localRoot = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "MultiRvt", DateTime.Now.ToString("yyyyMMdd"))
            Directory.CreateDirectory(localRoot)

            Dim fileName = Path.GetFileNameWithoutExtension(centralPath) & "_" & Environment.UserName & "_" & DateTime.Now.ToString("HHmmssfff") & ".rvt"
            Dim localPath = Path.Combine(localRoot, fileName)

            Dim sourcePath = ModelPathUtils.ConvertUserVisiblePathToModelPath(centralPath)
            Dim targetPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath)
            WorksharingUtils.CreateNewLocal(sourcePath, targetPath)
            Return localPath
        End Function

        Private Shared Function TrySynchronizeMultiLocalToCentral(doc As Document, comment As String, ByRef err As String) As Boolean
            err = String.Empty
            If doc Is Nothing Then
                err = "문서가 없습니다."
                Return False
            End If

            If Not doc.IsWorkshared Then
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

        Private Shared Sub TryDeleteMultiTempFile(pathText As String)
            If String.IsNullOrWhiteSpace(pathText) OrElse Not File.Exists(pathText) Then Return
            Try
                File.Delete(pathText)
            Catch
            End Try
        End Sub

        Private Shared Function BuildReopenOpenOptions() As OpenOptions
            Dim opt As New OpenOptions()
            Try
                Dim ws = New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opt.SetOpenWorksetsConfiguration(ws)
            Catch
            End Try
            Return opt
        End Function

        Private Shared Function BuildConnectorHeaders(extras As IList(Of String), uiUnit As String) As List(Of String)
            Dim distanceHeader As String = "Distance (inch)"
            If String.Equals(uiUnit, "mm", StringComparison.OrdinalIgnoreCase) Then
                distanceHeader = "Distance (mm)"
            End If

            ' ✅ 요청 스키마 반영
            ' - Category2 ↔ Family1 사이에 "검토내용", "비고(답변)" 2열 추가(값은 빈칸)
            ' - Status, ErrorMessage 컬럼은 엑셀 헤더에서 제외
            Dim headers As New List(Of String) From {
                "File", "Id1", "Id2",
                "Category1", "Category2",
                "검토내용", "비고(답변)",
                "Family1", "Family2",
                distanceHeader,
                "ConnectionType",
                "ParamName",
                "Value1", "Value2",
                "ParamCompare"
            }

            If extras IsNot Nothing Then
                For Each name In extras
                    headers.Add($"{name}(ID1)")
                    headers.Add($"{name}(ID2)")
                Next
            End If
            Return headers
        End Function

        Private Shared Function BuildConnectorTableFromRows(headers As IList(Of String), rows As IList(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Export")
            For Each h In headers
                If String.Equals(h, "Distance (mm)", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(h, "Distance (inch)", StringComparison.OrdinalIgnoreCase) Then
                    dt.Columns.Add(h, GetType(Double))
                Else
                    dt.Columns.Add(h)
                End If
            Next
            If rows Is Nothing OrElse rows.Count = 0 Then
                Dim dr = dt.NewRow()
                dr(0) = "오류가 없습니다."
                dt.Rows.Add(dr)
                Return dt
            End If

            For Each r In rows
                Dim dr = dt.NewRow()
                For i As Integer = 0 To headers.Count - 1
                    Dim key = headers(i)
                    If String.Equals(key, "Distance (mm)", StringComparison.OrdinalIgnoreCase) Then
                        Dim distInch As Double = GetDistanceInch(r)
                        If Not Double.IsNaN(distInch) Then
                            dr(i) = distInch * 25.4R
                        Else
                            dr(i) = DBNull.Value
                        End If
                    ElseIf String.Equals(key, "Distance (inch)", StringComparison.OrdinalIgnoreCase) Then
                        Dim distInch As Double = GetDistanceInch(r)
                        If Not Double.IsNaN(distInch) Then
                            dr(i) = distInch
                        Else
                            dr(i) = DBNull.Value
                        End If
                    Else
                        If String.Equals(key, "검토내용", StringComparison.Ordinal) Then
                            dr(i) = BuildConnectorReviewTextForExport(r)
                        ElseIf String.Equals(key, "ParamCompare", StringComparison.Ordinal) Then
                            dr(i) = NormalizeConnectorParamCompareForExport(r)
                        ElseIf String.Equals(key, "비고(답변)", StringComparison.Ordinal) Then
                            dr(i) = ""
                        Else
                            dr(i) = If(r IsNot Nothing AndAlso r.ContainsKey(key) AndAlso r(key) IsNot Nothing, r(key).ToString(), String.Empty)
                        End If
                    End If
                Next
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        Private Sub AppendMultiConnectorError(fileName As String, message As String)
            If _multiRequest Is Nothing OrElse _multiRequest.Connector Is Nothing OrElse Not _multiRequest.Connector.Enabled Then Return
            If _multiConnectorRows Is Nothing Then _multiConnectorRows = New List(Of Dictionary(Of String, Object))()
            _multiConnectorRows.Add(New Dictionary(Of String, Object) From {
                {"File", fileName},
                {"ConnectionType", "ERROR"},
                {"ParamName", _multiRequest.Connector.Param},
                {"ParamCompare", "N/A"},
                {"Status", "ERROR"},
                {"ErrorMessage", message}
            })
        End Sub

        Private Shared Function BuildTableFromRows(headers As IList(Of String), rows As IList(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Export")
            For Each h In headers
                dt.Columns.Add(h)
            Next
            If rows Is Nothing Then Return dt
            For Each r In rows
                Dim dr = dt.NewRow()
                For i As Integer = 0 To headers.Count - 1
                    Dim key = headers(i)
                    dr(i) = If(r IsNot Nothing AndAlso r.ContainsKey(key) AndAlso r(key) IsNot Nothing, r(key).ToString(), String.Empty)
                Next
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        Private Shared Sub AddEmptyMessageRow(table As DataTable)
            ExcelCore.EnsureNoDataRow(table, "오류가 없습니다.")
        End Sub

        Private Shared Function ValidateSchema(table As DataTable, headers As IList(Of String)) As Boolean
            If table Is Nothing OrElse headers Is Nothing Then Return False
            If table.Columns.Count <> headers.Count Then Return False
            For i As Integer = 0 To headers.Count - 1
                If Not String.Equals(table.Columns(i).ColumnName, headers(i), StringComparison.Ordinal) Then Return False
            Next
            Return True
        End Function

        Private Sub AppendSegmentPmsRows(run As SegmentPmsCheckService.RunResult, ds As DataSet)
            If run Is Nothing Then Return
            Dim classRows = SegmentPmsCheckService.BuildClassCheckRows(run.MapTable)
            Dim sizeRows = SegmentPmsCheckService.BuildSizeCheckRows(run.CompareTable)
            Dim routingRows = SegmentPmsCheckService.BuildRoutingClassRows(ds)

            If _multiPmsClassRows Is Nothing Then _multiPmsClassRows = New List(Of Dictionary(Of String, Object))()
            If _multiPmsSizeRows Is Nothing Then _multiPmsSizeRows = New List(Of Dictionary(Of String, Object))()
            If _multiPmsRoutingRows Is Nothing Then _multiPmsRoutingRows = New List(Of Dictionary(Of String, Object))()
            _multiPmsClassRows.AddRange(If(classRows, New List(Of Dictionary(Of String, Object))()))
            _multiPmsSizeRows.AddRange(If(sizeRows, New List(Of Dictionary(Of String, Object))()))
            _multiPmsRoutingRows.AddRange(If(routingRows, New List(Of Dictionary(Of String, Object))()))
        End Sub

        Private Sub MergeGuidResult(res As GuidAuditService.RunResult)
            If res Is Nothing Then Return
            _multiGuidProject = MergeTable(_multiGuidProject, res.Project)
            If res.IncludeFamily Then
                _multiGuidFamilyDetail = MergeTable(_multiGuidFamilyDetail, res.FamilyDetail)
                _multiGuidFamilyIndex = MergeTable(_multiGuidFamilyIndex, res.FamilyIndex)
            End If
        End Sub

        Private Shared Function FilterIssueRowsFromDict(styleKey As String, rows As List(Of Dictionary(Of String, Object))) As List(Of Dictionary(Of String, Object))
            Dim source As List(Of Dictionary(Of String, Object)) = If(rows, New List(Of Dictionary(Of String, Object))())
            If source.Count = 0 Then Return source

            Dim table As DataTable = DictListToDataTable(source, "ReviewRows")
            Dim filtered As DataTable = FilterIssueRowsCopy(styleKey, table)
            Return DataTableToObjects(filtered)
        End Function

        Private Shared Function MergeTable(master As DataTable, part As DataTable) As DataTable
            If part Is Nothing Then Return master
            If master Is Nothing Then
                master = part.Clone()
            End If
            For Each r As DataRow In part.Rows
                master.ImportRow(r)
            Next
            Return master
        End Function

        Private Shared Function BuildMappingsFromSuggestions(suggestions As IList(Of SegmentPmsCheckService.SuggestedMapping)) As List(Of SegmentPmsCheckService.MappingSelection)
            Dim list As New List(Of SegmentPmsCheckService.MappingSelection)()
            If suggestions Is Nothing Then Return list
            For Each s In suggestions
                If s Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(s.PmsSegmentKey) Then Continue For
                Dim item As New SegmentPmsCheckService.MappingSelection With {
                    .File = s.File,
                    .PipeTypeName = s.PipeTypeName,
                    .RuleIndex = s.RuleIndex,
                    .SegmentId = s.SegmentId,
                    .SegmentKey = s.SegmentKey,
                    .SelectedClass = s.PmsClass,
                    .SelectedPmsSegment = s.PmsSegmentKey,
                    .MappingSource = "AutoSuggest"
                }
                list.Add(item)
            Next
            Return list
        End Function

        Private Shared Function BuildPointHeaders(unit As String) As List(Of String)
            Dim suffix As String = "(ft)"
            If String.Equals(unit, "m", StringComparison.OrdinalIgnoreCase) Then
                suffix = "(m)"
            ElseIf String.Equals(unit, "mm", StringComparison.OrdinalIgnoreCase) Then
                suffix = "(mm)"
            End If
            Return New List(Of String) From {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            }
        End Function

        Private Shared Function BuildPointTable(headers As IList(Of String), rows As IList(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Points")
            For Each h In headers
                dt.Columns.Add(h)
            Next
            If rows Is Nothing OrElse rows.Count = 0 Then
                Dim dr = dt.NewRow()
                dr(0) = "오류가 없습니다."
                dt.Rows.Add(dr)
                Return dt
            End If
            For Each r In rows
                Dim dr = dt.NewRow()
                dr(0) = SafeStr(GetRowValue(r, "File"))
                dr(1) = SafeStr(GetRowValue(r, "ProjectPoint_E"))
                dr(2) = SafeStr(GetRowValue(r, "ProjectPoint_N"))
                dr(3) = SafeStr(GetRowValue(r, "ProjectPoint_Z"))
                dr(4) = SafeStr(GetRowValue(r, "SurveyPoint_E"))
                dr(5) = SafeStr(GetRowValue(r, "SurveyPoint_N"))
                dr(6) = SafeStr(GetRowValue(r, "SurveyPoint_Z"))
                dr(7) = SafeStr(GetRowValue(r, "TrueNorthAngle"))
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        Private Shared Function BuildLinkWorksetTable(headers As IList(Of String), rows As IList(Of LinkWorksetAuditRow)) As DataTable
            Dim dt As New DataTable("LinkWorkset")
            For Each h In headers
                dt.Columns.Add(h)
            Next

            If rows Is Nothing OrElse rows.Count = 0 Then
                Dim dr = dt.NewRow()
                dr(0) = "오류가 없습니다."
                dt.Rows.Add(dr)
                Return dt
            End If

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim dr = dt.NewRow()
                dr(0) = SafeStr(row.HostFileName)
                dr(1) = SafeStr(row.LinkName)
                dr(2) = SafeStr(row.AttachmentType)
                dr(3) = BoolText(row.WasLoadedBefore)
                dr(4) = BoolText(row.IsLoadedAfter)
                dr(5) = BoolText(row.IsWorkshared)
                dr(6) = SafeStr(row.DefaultWorksetName)
                dr(7) = row.TotalUserWorksets.ToString()
                dr(8) = SafeStr(row.OpenUserWorksetNamesBefore)
                dr(9) = NullableBoolText(row.DefaultOnlyOpenBefore)
                dr(10) = SafeStr(row.OpenUserWorksetNamesAfter)
                dr(11) = NullableBoolText(row.DefaultOnlyOpenAfter)
                dr(12) = BoolText(row.ApplyRequested)
                dr(13) = BoolText(row.Applied)
                dr(14) = SafeStr(row.Status)
                dr(15) = SafeStr(row.Message)
                dt.Rows.Add(dr)
            Next

            Return dt
        End Function

        Private Shared Function ConvertPoint(valueFt As Double, unit As String) As Double
            If String.Equals(unit, "m", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Round(valueFt * 0.3048R, 6)
            End If
            If String.Equals(unit, "mm", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Round(valueFt * 304.8R, 3)
            End If
            Return Math.Round(valueFt, 6)
        End Function

        Private Shared Function ToDict(obj As Object) As Dictionary(Of String, Object)
            Dim dict = TryCast(obj, Dictionary(Of String, Object))
            If dict IsNot Nothing Then Return dict
            Dim res As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If obj Is Nothing Then Return res
            Dim t = obj.GetType()
            For Each p In t.GetProperties()
                Try
                    res(p.Name) = p.GetValue(obj, Nothing)
                Catch
                End Try
            Next
            Return res
        End Function

        Private Shared Function ExtractStringList(dict As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim list As New List(Of String)()
            Dim raw = GetDictValue(dict, key)
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr IsNot Nothing AndAlso Not TypeOf raw Is String Then
                For Each o In arr
                    Dim s = NormalizeMultiPath(SafeStr(o))
                    If Not String.IsNullOrWhiteSpace(s) AndAlso Not list.Contains(s) Then
                        list.Add(s)
                    End If
                Next
            Else
                Dim s = NormalizeMultiPath(SafeStr(raw))
                If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
            End If
            Return list
        End Function

        Private Shared Function ToBool(obj As Object, Optional defaultValue As Boolean = False) As Boolean
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToBoolean(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function ToDouble(obj As Object, defaultValue As Double) As Double
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToDouble(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function ToInt(obj As Object, defaultValue As Integer) As Integer
            If obj Is Nothing Then Return defaultValue
            Try
                Return Convert.ToInt32(obj)
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetRowValue(row As Dictionary(Of String, Object), key As String) As Object
            If row Is Nothing Then Return Nothing
            Dim val As Object = Nothing
            If row.TryGetValue(key, val) Then Return val
            Return Nothing
        End Function

        Private Shared Function BoolText(value As Boolean) As String
            Return If(value, "Y", "N")
        End Function

        Private Shared Function NullableBoolText(value As Boolean?) As String
            If Not value.HasValue Then Return "N/A"
            Return BoolText(value.Value)
        End Function

    End Class

End Namespace
