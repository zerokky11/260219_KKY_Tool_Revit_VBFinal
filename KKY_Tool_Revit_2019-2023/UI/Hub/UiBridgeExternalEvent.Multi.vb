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
            Public Property ExcludeTargetFilter As String = String.Empty
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
                    .excludeTargetFilterText = stored.ExcludeTargetFilterText,
                    .excludeEndDummy = False,
                    .includePointXY = stored.IncludePointXY,
                    .includeLinearMetrics = stored.IncludeLinearMetrics
                })
            Catch ex As Exception
                SendToWeb("commonoptions:loaded", New With {
                    .extraParamsText = "",
                    .targetFilterText = "",
                    .excludeTargetFilterText = "",
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
            Dim excludeFilterText As String = Convert.ToString(GetProp(pd, "excludeTargetFilterText"))
            If String.IsNullOrWhiteSpace(excludeFilterText) Then
                excludeFilterText = Convert.ToString(GetProp(pd, "excludeTargetFilter"))
            End If
            Dim includePointXY As Boolean = SafeBoolObj(GetProp(pd, "includePointXY"), False)
            Dim includeLinearMetrics As Boolean = SafeBoolObj(GetProp(pd, "includeLinearMetrics"), False)
            Dim options As New HubCommonOptionsStorageService.HubCommonOptions() With {
                .ExtraParamsText = If(extraText, String.Empty),
                .TargetFilterText = If(filterText, String.Empty),
                .ExcludeTargetFilterText = If(excludeFilterText, String.Empty),
                .ExcludeEndDummy = False,
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
            Public Property ExcludeEndDummy As Boolean
            Public Property IncludePointXY As Boolean
            Public Property IncludeLinearMetrics As Boolean
        End Class

        Private Class MultiTapAlignOptions
            Public Property Enabled As Boolean
            Public Property Tol As Double = 0.5R
            Public Property Unit As String = "mm"
            Public Property Domain As String = "all"
            Public Property FeatureTargetFilter As String = String.Empty
            Public Property ExportLocale As String = "ko"
        End Class

        Private Class MultiDupClashOptions
            Public Property Enabled As Boolean
            Public Property Mode As String = "duplicate"
            Public Property TolFeet As Double = 1.0R / 64.0R
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
            Public Property FamilySuitability As MultiFamilySuitabilityOptions = New MultiFamilySuitabilityOptions()
            Public Property TapAlign As MultiTapAlignOptions = New MultiTapAlignOptions()
            Public Property DupClash As MultiDupClashOptions = New MultiDupClashOptions()
            Public Property WorksetAssignment As MultiWorksetAssignmentOptions = New MultiWorksetAssignmentOptions()
            Public Property ProjectParameterDuplication As MultiProjectParameterDuplicationOptions = New MultiProjectParameterDuplicationOptions()
            Public Property ParameterMissing As MultiParameterMissingOptions = New MultiParameterMissingOptions()
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
        Private Shared _multiTapAlignRows As List(Of Dictionary(Of String, Object))
        Private Shared _multiTapAlignExtras As List(Of String)
        Private Shared _multiTapAlignUnit As String = "mm"
        Private Shared _multiTapAlignLocale As String = "ko"
        Private Shared _multiLastExportFolder As String = String.Empty
        Private Shared _multiDupRows As List(Of Exports.DupRowDto)
        Private Shared _multiDupTargetCounts As Dictionary(Of String, Integer)
        Private Shared _multiClashRows As List(Of Exports.DupRowDto)
        Private Shared _multiClashPairs As List(Of Exports.PairRowDto)
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
                Case "tapalign"
                    _multiTapAlignRows = Nothing
                    _multiTapAlignExtras = Nothing
                    _multiTapAlignUnit = "mm"
                    _multiTapAlignLocale = "ko"
                Case "dupclash"
                    _multiDupRows = Nothing
                    _multiDupTargetCounts = Nothing
                    _multiClashRows = Nothing
                    _multiClashPairs = Nothing
                Case "worksetassignment"
                    ClearMultiWorksetAssignmentCache()
                Case "parameterduplication"
                    ClearMultiProjectParameterDuplicationCache()
                Case "floorinfo"
                    ClearMultiFloorInfoCache()
                Case "familysuitability"
                    ClearMultiFamilySuitabilityCache()
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
        '  { rvtPaths:[], commonOptions:{extraParams,targetFilter,excludeTargetFilter},
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

            Try
                PrepareFamilySuitabilityCriteria(req)
            Catch ex As Exception
                SendToWeb("hub:multi-error", New With {.message = ex.Message})
                Return
            End Try

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
                Dim rows = ConnectorDiagnosticsService.RunOnDocument(doc, _multiRequest.Connector.Tol, _multiRequest.Connector.Unit, _multiRequest.Connector.Param, extras, _multiRequest.Common.TargetFilter, _multiRequest.Common.ExcludeTargetFilter, _multiRequest.Connector.ExcludeEndDummy, Sub(pct, msg)
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

            If _multiRequest.FamilySuitability.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Family 적합성 검토 실행 중", safeName)
                RunFamilySuitabilityMultiForDocument(doc, safeName, basePct)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Family 적합성 검토 완료", safeName)
            End If

            If _multiRequest.TapAlign.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "탭/분기 축 틀어짐 검토 실행 중", safeName)
                Dim extras = BuildConnectorExtraParams(_multiRequest.Common.ExtraParams,
                                                       _multiRequest.Common.IncludePointXY,
                                                       _multiRequest.Common.IncludeLinearMetrics)
                Dim combinedTargetFilter = TapAlignmentReviewService.CombineTargetFilterText(_multiRequest.Common.TargetFilter,
                                                                                            _multiRequest.TapAlign.FeatureTargetFilter)
                Dim targetCount = TapAlignmentReviewService.CountTargetsOnDocument(doc,
                                                                                   _multiRequest.TapAlign.Domain,
                                                                                   combinedTargetFilter,
                                                                                   _multiRequest.Common.ExcludeTargetFilter)
                Dim rows = TapAlignmentReviewService.RunOnDocument(doc,
                                                                   _multiRequest.TapAlign.Tol,
                                                                   _multiRequest.TapAlign.Unit,
                                                                   _multiRequest.TapAlign.Domain,
                                                                   extras,
                                                                   combinedTargetFilter,
                                                                   _multiRequest.Common.ExcludeTargetFilter,
                                                                   Sub(pct, msg)
                                                                       Dim fraction As Double = Math.Max(0.0R, Math.Min(CDbl(pct), 1.0R))
                                                                       Dim overallPct = ((basePct + fraction / Math.Max(_multiTotal, 1)) * 100.0R)
                                                                       ReportMultiProgress(overallPct, "탭/분기 축 틀어짐 검토 실행 중", $"{safeName} · {msg}")
                                                                   End Sub)
                If rows IsNot Nothing AndAlso rows.Count > 0 Then
                    For Each row In rows
                        If row IsNot Nothing Then row("File") = safeName
                    Next
                    If _multiTapAlignRows Is Nothing Then _multiTapAlignRows = New List(Of Dictionary(Of String, Object))()
                    _multiTapAlignRows.AddRange(rows)
                Else
                    If _multiTapAlignRows Is Nothing Then _multiTapAlignRows = New List(Of Dictionary(Of String, Object))()
                    _multiTapAlignRows.Add(New Dictionary(Of String, Object) From {
                        {"File", safeName},
                        {"Status", If(targetCount > 0, "OK", "NO_TARGET")},
                        {"TargetCount", targetCount.ToString(Globalization.CultureInfo.InvariantCulture)},
                        {"Message", ""},
                        {"DistanceFromCenter", ""},
                        {"ModeledAngle", ""},
                        {"Domain", ""}
                    })
                End If
                _multiTapAlignExtras = extras
                _multiTapAlignUnit = NormalizeTapAlignUnit(_multiRequest.TapAlign.Unit)
                _multiTapAlignLocale = NormalizeTapAlignExportLocale(_multiRequest.TapAlign.ExportLocale)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "탭/분기 축 틀어짐 검토 완료", safeName)
            End If

            If _multiRequest.DupClash.Enabled Then
                stepIndex += 1
                Dim dupClashMode As String = NormalizeMultiDupClashMode(_multiRequest.DupClash.Mode)
                Dim dupClashLabel As String = ResolveMultiDupClashModeLabel(dupClashMode)
                ReportMultiProgress(CalcStepProgressPercent(basePct, stepIndex, steps, 0.0R), $"{dupClashLabel} 실행 중", safeName)
                RunDupClashMultiForDocument(doc, safeName, basePct, stepIndex, steps)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), $"{dupClashLabel} 완료", safeName)
            End If

            If _multiRequest.WorksetAssignment.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "웍셋 배정 검토 실행 중", safeName)
                RunWorksetAssignmentMultiForDocument(doc, safeName, basePct)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "웍셋 배정 검토 완료", safeName)
            End If

            If _multiRequest.ProjectParameterDuplication.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Project Parameter 중복 검토 실행 중", safeName)
                RunProjectParameterDuplicationMultiForDocument(doc, safeName, basePct)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "Project Parameter 중복 검토 완료", safeName)
            End If

            If _multiRequest.ParameterMissing.Enabled Then
                stepIndex += 1
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "파라미터 누락 검토 실행 중", safeName)
                RunParameterMissingMultiForDocument(doc, safeName, basePct)
                ReportMultiProgress(CalcStepPercent(basePct, stepIndex, steps), "파라미터 누락 검토 완료", safeName)
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
                ReportMultiProgress(CalcStepProgressPercent(basePct, stepIndex, steps, 0.0R), "패밀리 연동 검토 실행 중", safeName)
                Dim rows = FamilyLinkAuditService.RunOnDocument(doc, path, _multiRequest.FamilyLink.Targets,
                                                                Sub(pct, msg)
                                                                    Dim overallPct As Double = CalcStepProgressPercent(basePct, stepIndex, steps, CDbl(pct) / 100.0R)
                                                                    Dim detail As String = BuildMultiFamilyLinkProgressDetail(safeName, msg)
                                                                    ReportMultiProgress(overallPct, "패밀리 연동 검토 실행 중", detail)
                                                                End Sub)
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

        Private Sub RunDupClashMultiForDocument(doc As Document,
                                                safeName As String,
                                                basePct As Double,
                                                stepIndex As Integer,
                                                steps As Integer)
            If doc Is Nothing Then Return

            Dim tolFeet As Double = 1.0R / 64.0R
            Dim mode As String = "duplicate"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.DupClash IsNot Nothing Then
                tolFeet = Math.Max(0.000001R, _multiRequest.DupClash.TolFeet)
                mode = NormalizeMultiDupClashMode(_multiRequest.DupClash.Mode)
            End If

            Dim targetFilter As String = ""
            Dim excludeTargetFilter As String = ""
            Dim extraParamNames As List(Of String) = New List(Of String)()
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                targetFilter = SafeStr(_multiRequest.Common.TargetFilter)
                excludeTargetFilter = SafeStr(_multiRequest.Common.ExcludeTargetFilter)
                extraParamNames = ParseExtraParams(_multiRequest.Common.ExtraParams)
            End If

            Dim previousRows = _lastRows
            Dim previousPairs = _lastPairs
            Dim previousMode = _lastMode
            Dim previousTargetCount = _lastDupClashTargetCount
            Dim dupStart As Integer = If(_multiDupRows, New List(Of Exports.DupRowDto)()).Count
            Dim clashRowStart As Integer = If(_multiClashRows, New List(Of Exports.DupRowDto)()).Count
            Dim clashPairStart As Integer = If(_multiClashPairs, New List(Of Exports.PairRowDto)()).Count

            Try
                PrepareNestedSharedIds(doc)
                Dim scopeIds = BuildDupClashScopeIds(doc, targetFilter, excludeTargetFilter)

                If String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase) Then
                    RunSelfClash(doc, tolFeet, scopeIds, Nothing, Nothing)
                    If _multiClashRows Is Nothing Then _multiClashRows = New List(Of Exports.DupRowDto)()
                    _multiClashRows.AddRange(BuildMultiDupExportRows(doc, safeName, _lastRows, extraParamNames))

                    If _multiClashPairs Is Nothing Then _multiClashPairs = New List(Of Exports.PairRowDto)()
                    _multiClashPairs.AddRange(BuildMultiClashExportPairs(doc, safeName, _lastPairs, extraParamNames))
                Else
                    RunDuplicate(doc, tolFeet, scopeIds, Nothing, Nothing)
                    If _multiDupRows Is Nothing Then _multiDupRows = New List(Of Exports.DupRowDto)()
                    _multiDupRows.AddRange(BuildMultiDupExportRows(doc, safeName, _lastRows, extraParamNames))
                    If _multiDupTargetCounts Is Nothing Then _multiDupTargetCounts = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                    _multiDupTargetCounts(ResolveRequestedMultiFileName(safeName)) = _lastDupClashTargetCount
                End If
            Catch
                TrimMultiDupRows(dupStart)
                TrimMultiClashRows(clashRowStart)
                TrimMultiClashPairs(clashPairStart)
                Throw
            Finally
                _lastRows = previousRows
                _lastPairs = previousPairs
                _lastMode = previousMode
                _lastDupClashTargetCount = previousTargetCount
            End Try
        End Sub

        Private Shared Function NormalizeMultiDupClashMode(value As String) As String
            Dim normalized As String = SafeStr(value).Trim().ToLowerInvariant()
            If normalized = "clash" OrElse normalized = "selfclash" OrElse normalized = "self-clash" OrElse normalized = "interference" Then
                Return "clash"
            End If
            Return "duplicate"
        End Function

        Private Shared Function ResolveMultiDupClashModeLabel(value As String) As String
            If String.Equals(NormalizeMultiDupClashMode(value), "clash", StringComparison.OrdinalIgnoreCase) Then
                Return "자체간섭 검토"
            End If
            Return "중복 검토"
        End Function

        Private Function GetCurrentMultiDupClashMode() As String
            If _multiRequest Is Nothing OrElse _multiRequest.DupClash Is Nothing Then
                Return "duplicate"
            End If
            Return NormalizeMultiDupClashMode(_multiRequest.DupClash.Mode)
        End Function

        Private Function BuildDupClashScopeIds(doc As Document,
                                               targetFilter As String,
                                               excludeTargetFilter As String) As HashSet(Of Integer)
            If doc Is Nothing Then Return Nothing
            If String.IsNullOrWhiteSpace(targetFilter) AndAlso String.IsNullOrWhiteSpace(excludeTargetFilter) Then Return Nothing

            Dim evaluator = ConnectorDiagnosticsService.CreateCommonOptionsElementEvaluator(targetFilter, excludeTargetFilter)
            If evaluator Is Nothing Then Return Nothing

            Dim allowed As New HashSet(Of Integer)()
            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()

            For Each e As Element In collector
                If e Is Nothing Then Continue For

                Dim ok As Boolean = False
                Try
                    ok = evaluator(e)
                Catch
                    ok = False
                End Try

                If ok Then
                    allowed.Add(e.Id.IntegerValue)
                End If
            Next

            If allowed.Count = 0 Then
                allowed.Add(Integer.MinValue)
            End If

            Return allowed
        End Function

        Private Function TryBuildCommonScopeIds(doc As Document,
                                                targetFilter As String,
                                                excludeTargetFilter As String,
                                                ByRef allowedElementIds As List(Of Integer)) As Boolean
            allowedElementIds = New List(Of Integer)()
            If doc Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(targetFilter) AndAlso String.IsNullOrWhiteSpace(excludeTargetFilter) Then Return False

            Dim scopeIds = BuildDupClashScopeIds(doc, targetFilter, excludeTargetFilter)
            If scopeIds Is Nothing Then Return False

            allowedElementIds = scopeIds.
                Where(Function(id) id > 0).
                Distinct().
                ToList()
            Return True
        End Function

        Private Function BuildMultiDupExportRows(doc As Document,
                                                 safeName As String,
                                                 rows As IEnumerable(Of DupRowDto),
                                                 extraParamNames As IList(Of String)) As List(Of Exports.DupRowDto)
            Dim result As New List(Of Exports.DupRowDto)()
            If rows Is Nothing Then Return result

            For Each row In rows
                If row Is Nothing Then Continue For

                result.Add(New Exports.DupRowDto With {
                    .FileName = ResolveRequestedMultiFileName(ResolveMultiDupFileName(safeName, row.FileName)),
                    .Id = row.ElementId.ToString(),
                    .Category = SafeStr(row.Category),
                    .Family = SafeStr(row.Family),
                    .Type = SafeStr(row.Type),
                    .Comment = SafeStr(row.Comment),
                    .ConnectedIds = ParseMultiDupConnectedIds(row.ConnectedIds),
                    .GroupKey = SafeStr(row.GroupKey),
                    .ExtraParams = ReadElementParameterMap(doc, row.ElementId, extraParamNames)
                })
            Next

            Return result
        End Function

        Private Function BuildMultiClashExportPairs(doc As Document,
                                                    safeName As String,
                                                    pairs As IEnumerable(Of PairRowDto),
                                                    extraParamNames As IList(Of String)) As List(Of Exports.PairRowDto)
            Dim result As New List(Of Exports.PairRowDto)()
            If pairs Is Nothing Then Return result

            For Each pair In pairs
                If pair Is Nothing Then Continue For

                Dim aId As Integer = SafeToInt(pair.AId)
                Dim bId As Integer = SafeToInt(pair.BId)

                result.Add(New Exports.PairRowDto With {
                    .FileName = ResolveRequestedMultiFileName(ResolveMultiDupFileName(safeName, pair.FileName)),
                    .GroupKey = SafeStr(pair.GroupKey),
                    .AId = SafeStr(pair.AId),
                    .ACategory = SafeStr(pair.ACategory),
                    .AFamily = SafeStr(pair.AFamily),
                    .AType = SafeStr(pair.AType),
                    .BId = SafeStr(pair.BId),
                    .BCategory = SafeStr(pair.BCategory),
                    .BFamily = SafeStr(pair.BFamily),
                    .BType = SafeStr(pair.BType),
                    .Comment = SafeStr(pair.Comment),
                    .AExtraParams = ReadElementParameterMap(doc, aId, extraParamNames),
                    .BExtraParams = ReadElementParameterMap(doc, bId, extraParamNames)
                })
            Next

            Return result
        End Function

        Private Shared Function ParseMultiDupConnectedIds(raw As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(raw) Then Return result

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim parts = raw.Split(New Char() {","c, ";"c, "|"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
            For Each part In parts
                Dim value As String = SafeStr(part).Trim()
                If String.IsNullOrWhiteSpace(value) Then Continue For
                If seen.Add(value) Then result.Add(value)
            Next

            Return result
        End Function

        Private Shared Function ResolveMultiDupFileName(defaultName As String, candidate As String) As String
            Dim resolved As String = GetSafeMultiFileName(candidate)
            If Not String.IsNullOrWhiteSpace(resolved) Then Return resolved
            Return GetSafeMultiFileName(defaultName)
        End Function

        Private Function BuildMultiDuplicateExportRowsWithPlaceholders(sourceRows As IEnumerable(Of Exports.DupRowDto),
                                                                       exportLocale As String) As List(Of Exports.DupRowDto)
            Dim rows As List(Of Exports.DupRowDto) =
                If(sourceRows, Enumerable.Empty(Of Exports.DupRowDto)()).
                Where(Function(item) item IsNot Nothing).
                ToList()
            Dim orderedNames = BuildOrderedMultiFileNames(rows.Select(Function(item) If(item Is Nothing, "", item.FileName)))
            Dim result As New List(Of Exports.DupRowDto)()

            For Each fileName In orderedNames
                Dim perFileRows = rows.
                    Where(Function(item) String.Equals(ResolveRequestedMultiFileName(item.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                If perFileRows.Count > 0 Then
                    result.AddRange(perFileRows)
                Else
                    result.Add(BuildMultiDuplicatePlaceholderRow(fileName, exportLocale))
                End If
            Next

            If result.Count = 0 Then
                result.Add(BuildMultiDuplicatePlaceholderRow("", exportLocale))
            End If

            Return result
        End Function

        Private Function BuildMultiDuplicatePlaceholderRowsForFile(fileName As String,
                                                                   exportLocale As String) As List(Of Exports.DupRowDto)
            Return New List(Of Exports.DupRowDto) From {
                BuildMultiDuplicatePlaceholderRow(fileName, exportLocale)
            }
        End Function

        Private Function BuildMultiDuplicatePlaceholderRow(fileName As String,
                                                           exportLocale As String) As Exports.DupRowDto
            Dim safeFileName As String = ResolveRequestedMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(safeFileName) Then safeFileName = SafeStr(fileName)

            Dim runReason As String = ""
            Dim statusText As String = GetMultiRunItemStatus(safeFileName, runReason)
            Dim comment As String

            If String.Equals(statusText, "failed", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(statusText, "skipped", StringComparison.OrdinalIgnoreCase) Then
                comment = If(String.IsNullOrWhiteSpace(runReason),
                             Exports.DuplicateExport.GetDuplicateNoResultComment(exportLocale),
                             runReason)
            Else
                Dim targetCount As Integer = GetMultiDuplicateTargetCount(safeFileName)
                comment = If(targetCount <= 0,
                             Exports.DuplicateExport.GetDuplicateNoTargetComment(exportLocale),
                             Exports.DuplicateExport.GetDuplicateNoIssueComment(exportLocale))
            End If

            Return New Exports.DupRowDto With {
                .FileName = safeFileName,
                .Comment = comment
            }
        End Function

        Private Function GetMultiDuplicateTargetCount(fileName As String) As Integer
            Dim safeFileName As String = ResolveRequestedMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(safeFileName) Then safeFileName = SafeStr(fileName)

            If _multiDupTargetCounts Is Nothing Then Return 0

            Dim targetCount As Integer = 0
            If _multiDupTargetCounts.TryGetValue(safeFileName, targetCount) Then
                Return Math.Max(targetCount, 0)
            End If

            Return 0
        End Function

        Private Shared Sub TrimMultiDupRows(startIndex As Integer)
            If _multiDupRows Is Nothing Then Return
            If startIndex <= 0 Then
                _multiDupRows.Clear()
                Return
            End If
            If startIndex >= _multiDupRows.Count Then Return
            _multiDupRows.RemoveRange(startIndex, _multiDupRows.Count - startIndex)
        End Sub

        Private Shared Sub TrimMultiClashRows(startIndex As Integer)
            If _multiClashRows Is Nothing Then Return
            If startIndex <= 0 Then
                _multiClashRows.Clear()
                Return
            End If
            If startIndex >= _multiClashRows.Count Then Return
            _multiClashRows.RemoveRange(startIndex, _multiClashRows.Count - startIndex)
        End Sub

        Private Shared Sub TrimMultiClashPairs(startIndex As Integer)
            If _multiClashPairs Is Nothing Then Return
            If startIndex <= 0 Then
                _multiClashPairs.Clear()
                Return
            End If
            If startIndex >= _multiClashPairs.Count Then Return
            _multiClashPairs.RemoveRange(startIndex, _multiClashPairs.Count - startIndex)
        End Sub

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
            If _multiRequest IsNot Nothing AndAlso _multiRequest.FamilySuitability IsNot Nothing AndAlso _multiRequest.FamilySuitability.Enabled Then
                summary("familysuitability") = New With {.rows = GetMultiFamilySuitabilityRowCount()}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.TapAlign IsNot Nothing AndAlso _multiRequest.TapAlign.Enabled Then
                summary("tapalign") = New With {.rows = If(_multiTapAlignRows, New List(Of Dictionary(Of String, Object))()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.DupClash IsNot Nothing AndAlso _multiRequest.DupClash.Enabled Then
                Dim mode As String = NormalizeMultiDupClashMode(_multiRequest.DupClash.Mode)
                Dim dupCount As Integer = If(_multiDupRows, New List(Of Exports.DupRowDto)()).Count
                Dim clashCount As Integer = If(_multiClashPairs, New List(Of Exports.PairRowDto)()).Count
                If clashCount <= 0 Then clashCount = If(_multiClashRows, New List(Of Exports.DupRowDto)()).Count
                summary("dupclash") = New With {.rows = If(String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase), clashCount, dupCount)}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.WorksetAssignment IsNot Nothing AndAlso _multiRequest.WorksetAssignment.Enabled Then
                summary("worksetassignment") = New With {.rows = If(_multiWorksetAssignmentRows, New List(Of WorksetAssignmentReviewService.ReviewRow)()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.ProjectParameterDuplication IsNot Nothing AndAlso _multiRequest.ProjectParameterDuplication.Enabled Then
                summary("parameterduplication") = New With {.rows = If(_multiParameterDuplicationRows, New List(Of ProjectParameterDuplicationReviewService.ReviewRow)()).Count}
            End If
            If _multiRequest IsNot Nothing AndAlso _multiRequest.ParameterMissing IsNot Nothing AndAlso _multiRequest.ParameterMissing.Enabled Then
                summary("parametermissing") = New With {.rows = If(_multiParameterMissingRows, New List(Of ParameterMissingReviewService.ReviewRow)()).Count}
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
            Dim exportLocale As String = ParseExcelLocale(payload)
            Dim outputFolder As String = String.Empty

            If ParseSplitByFile(payload) Then
                outputFolder = PickMultiExportFolder()
                If String.IsNullOrWhiteSpace(outputFolder) Then
                    SendToWeb("hub:multi-exported", New With {
                        .ok = False,
                        .cancelled = True,
                        .message = "폴더 선택이 취소되었습니다."
                    })
                    Return
                End If
            End If

            Try
                Select Case If(key, "").ToLowerInvariant()
                    Case "connector"
                        ExportConnector(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "floorinfo"
                        ExportFloorInfo(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "familysuitability"
                        ExportFamilySuitability(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "tapalign"
                        ExportTapAlign(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "dupclash"
                        ExportDupClash(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "worksetassignment"
                        ExportWorksetAssignment(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "parameterduplication"
                        ExportProjectParameterDuplication(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "parametermissing"
                        ExportParameterMissing(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "pms"
                        ExportSegmentPms(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "guid"
                        ExportGuid(excelMode, exportLocale, outputFolder)
                    Case "familylink"
                        ExportFamilyLink(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "points"
                        ExportPoints(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case "linkworkset"
                        ExportLinkWorkset(doAutoFit, excelMode, exportLocale, outputFolder)
                    Case Else
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "알 수 없는 기능 키입니다."})
                        Return
                End Select
            Catch ex As Exception
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = ex.Message})
                Return
            End Try
        End Sub

        Private Shared Function ParseSplitByFile(payload As Object) As Boolean
            Try
                Return SafeBoolObj(GetProp(payload, "splitByFile"), False)
            Catch
                Return False
            End Try
        End Function

        Private Function PickMultiExportFolder() As String
            Dim initialDirectory As String = SafeStr(_multiLastExportFolder).Trim()
            If String.IsNullOrWhiteSpace(initialDirectory) OrElse Not Directory.Exists(initialDirectory) Then
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            Using dlg As New WinForms.OpenFileDialog()
                dlg.Title = "파일별 엑셀 저장 폴더 선택"
                dlg.Filter = "폴더 선택|*.folder"
                dlg.CheckFileExists = False
                dlg.CheckPathExists = False
                dlg.ValidateNames = False
                dlg.Multiselect = False
                dlg.RestoreDirectory = True
                dlg.DereferenceLinks = True
                dlg.InitialDirectory = initialDirectory
                dlg.FileName = "이 폴더 선택.folder"

                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return String.Empty

                Dim selectedPath As String = NormalizePickedFolderPath(dlg.FileName, dlg.InitialDirectory)
                If String.IsNullOrWhiteSpace(selectedPath) Then Return String.Empty

                If Not Directory.Exists(selectedPath) Then
                    Directory.CreateDirectory(selectedPath)
                End If

                _multiLastExportFolder = selectedPath
                Return selectedPath
            End Using
        End Function

        Private Shared Function NormalizePickedFolderPath(rawPath As String, fallbackDirectory As String) As String
            Dim value As String = SafeStr(rawPath).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            Try
                value = value.Replace("/"c, Path.DirectorySeparatorChar)
            Catch
            End Try

            Try
                If Directory.Exists(value) Then
                    Return Path.GetFullPath(value)
                End If
            Catch
            End Try

            Dim fileName As String = ""
            Try
                fileName = Path.GetFileName(value)
            Catch
                fileName = ""
            End Try

            If String.Equals(fileName, "이 폴더 선택.folder", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim parent As String = Path.GetDirectoryName(value)
                    If Not String.IsNullOrWhiteSpace(parent) Then
                        Return Path.GetFullPath(parent)
                    End If
                Catch
                End Try
            End If

            If Not String.IsNullOrWhiteSpace(fallbackDirectory) Then
                Try
                    If Not Path.IsPathRooted(value) Then
                        Dim combined As String = Path.Combine(fallbackDirectory, value)
                        If Directory.Exists(combined) Then
                            Return Path.GetFullPath(combined)
                        End If
                    End If
                Catch
                End Try
            End If

            Try
                Return Path.GetFullPath(value)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function NormalizeSplitExportLocale(exportLocale As String) As String
            If String.Equals(SafeStr(exportLocale).Trim(), "en", StringComparison.OrdinalIgnoreCase) Then Return "en"
            Return "ko"
        End Function

        Private Shared Function ResolveSplitExportFeatureFileLabel(featureKey As String,
                                                                  exportLocale As String,
                                                                  Optional fallbackLabel As String = Nothing,
                                                                  Optional featureMode As String = Nothing) As String
            Dim locale As String = NormalizeSplitExportLocale(exportLocale)
            Dim key As String = SafeStr(featureKey).Trim().ToLowerInvariant()
            Dim mode As String = SafeStr(featureMode).Trim().ToLowerInvariant()

            If String.Equals(locale, "en", StringComparison.OrdinalIgnoreCase) Then
                Select Case key
                    Case "connector"
                        Return "S5_UTILITY Continuity Error (Location-based)"
                    Case "tapalign"
                        Return "Tap Branch Axis Misalignment Review"
                    Case "dupclash"
                        If String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase) Then Return "Self Clash Review"
                        Return "Modeling Duplication"
                    Case "worksetassignment"
                        Return "Workset Assignment Error"
                    Case "parameterduplication"
                        Return "Parameter Duplication"
                    Case "parametermissing"
                        Return "Parameter Value Omission"
                    Case "familysuitability"
                        Return "Not Approved Family Review"
                End Select
            Else
                Select Case key
                    Case "connector"
                        Return "파라미터 연속성 검토"
                    Case "tapalign"
                        Return "탭분기 축 틀어짐 검토"
                    Case "dupclash"
                        If String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase) Then Return "자체간섭검토"
                        Return "중복검토"
                    Case "worksetassignment"
                        Return "웍셋 배정 검토"
                    Case "parameterduplication"
                        Return "Parameter 중복검토"
                    Case "parametermissing"
                        Return "속성누락검토"
                    Case "familysuitability"
                        Return "패밀리 적합성검토"
                End Select
            End If

            Dim fallback As String = SafeStr(fallbackLabel).Trim()
            If String.IsNullOrWhiteSpace(fallback) Then
                fallback = If(String.Equals(locale, "en", StringComparison.OrdinalIgnoreCase), "Review Result", "검토결과")
            End If
            Return fallback
        End Function

        Private Shared Function FormatSplitExportIssueCount(issueCount As Integer) As String
            Dim safeCount As Integer = Math.Max(0, issueCount)
            Return safeCount.ToString("00", Globalization.CultureInfo.InvariantCulture) & "EA"
        End Function

        Private Shared Function IsSplitExportMessageText(value As String) As Boolean
            Dim text As String = SafeStr(value).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return False

            Select Case text
                Case "오류가 없습니다.",
                     "No issues.",
                     "집계 가능한 객체가 없습니다.",
                     "No rows to export.",
                     "추출 결과 없음",
                     "No data."
                    Return True
            End Select

            Return False
        End Function

        Private Shared Function InferSplitExportIssueCount(table As DataTable) As Integer
            If table Is Nothing OrElse table.Rows.Count = 0 Then Return 0

            If table.Rows.Count = 1 Then
                Dim firstRow As DataRow = table.Rows(0)
                For Each column As DataColumn In table.Columns
                    If column Is Nothing Then Continue For
                    If IsSplitExportMessageText(SafeStr(firstRow(column.ColumnName))) Then Return 0
                Next
            End If

            Return Math.Max(0, table.Rows.Count)
        End Function

        Private Shared Function InferSplitExportIssueCount(sheets As IList(Of KeyValuePair(Of String, DataTable))) As Integer
            If sheets Is Nothing OrElse sheets.Count = 0 Then Return 0

            Dim total As Integer = 0
            For Each sheet In sheets
                total += InferSplitExportIssueCount(sheet.Value)
            Next
            Return Math.Max(0, total)
        End Function

        Private Shared Sub SetSplitExportIssueCount(fileIssueCounts As IDictionary(Of String, Integer), fileName As String, issueCount As Integer)
            If fileIssueCounts Is Nothing Then Return

            Dim safeName As String = GetSafeMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(safeName) Then Return

            fileIssueCounts(safeName) = Math.Max(0, issueCount)
        End Sub

        Private Shared Function ResolveSplitExportIssueCount(fileName As String,
                                                             fileIssueCounts As IDictionary(Of String, Integer),
                                                             fallbackCount As Integer) As Integer
            Dim safeName As String = GetSafeMultiFileName(fileName)
            If fileIssueCounts IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(safeName) Then
                Dim resolvedCount As Integer = 0
                If fileIssueCounts.TryGetValue(safeName, resolvedCount) Then
                    Return Math.Max(0, resolvedCount)
                End If
            End If

            Return Math.Max(0, fallbackCount)
        End Function

        Private Shared Function BuildSplitExportFilePath(outputFolder As String,
                                                         fileName As String,
                                                         featureKey As String,
                                                         exportLocale As String,
                                                         issueCount As Integer,
                                                         Optional fallbackLabel As String = Nothing,
                                                         Optional featureMode As String = Nothing) As String
            Dim safeFolder As String = SafeStr(outputFolder).Trim()
            If String.IsNullOrWhiteSpace(safeFolder) Then Throw New InvalidOperationException("저장 폴더가 선택되지 않았습니다.")

            Dim baseName As String = GetSafeMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(baseName) Then baseName = "Export"

            Dim featureLabel As String = ResolveSplitExportFeatureFileLabel(featureKey, exportLocale, fallbackLabel, featureMode)
            Dim safeSuffix As String = SanitizeFileName(SafeStr(featureLabel).Trim())
            If String.IsNullOrWhiteSpace(safeSuffix) Then safeSuffix = "검토결과"

            Dim issueSuffix As String = FormatSplitExportIssueCount(issueCount)
            Dim fileNameOnly As String = SanitizeFileName(baseName & "_" & safeSuffix & "_" & issueSuffix)
            If String.IsNullOrWhiteSpace(fileNameOnly) Then fileNameOnly = "Export"

            Dim fullPath As String = Path.Combine(safeFolder, fileNameOnly & ".xlsx")
            Return EnsureUniqueExportFilePath(fullPath)
        End Function

        Private Shared Function EnsureUniqueExportFilePath(filePath As String) As String
            Dim candidate As String = filePath
            If String.IsNullOrWhiteSpace(candidate) Then Return String.Empty
            If Not File.Exists(candidate) Then Return candidate

            Dim directoryPath As String = Path.GetDirectoryName(candidate)
            Dim baseName As String = Path.GetFileNameWithoutExtension(candidate)
            Dim extensionName As String = Path.GetExtension(candidate)
            Dim index As Integer = 2

            Do
                candidate = Path.Combine(directoryPath, $"{baseName} ({index}){extensionName}")
                index += 1
            Loop While File.Exists(candidate)

            Return candidate
        End Function

        Private Function SaveSplitSingleSheetTables(outputFolder As String,
                                                    featureKey As String,
                                                    featureSuffix As String,
                                                    sheetName As String,
                                                    fileTables As IList(Of KeyValuePair(Of String, DataTable)),
                                                    doAutoFit As Boolean,
                                                    excelMode As String,
                                                    exportLocale As String,
                                                    Optional progressKey As String = "hub:multi-progress",
                                                    Optional fileIssueCounts As IDictionary(Of String, Integer) = Nothing,
                                                    Optional featureMode As String = Nothing) As Integer
            If fileTables Is Nothing OrElse fileTables.Count = 0 Then Return 0

            Dim savedCount As Integer = 0
            Dim total As Integer = Math.Max(1, fileTables.Count)
            ExcelProgressReporter.Reset(progressKey)

            For i As Integer = 0 To fileTables.Count - 1
                Dim fileTable = fileTables(i)
                If fileTable.Value Is Nothing Then Continue For

                ExcelProgressReporter.Report(progressKey,
                                             "EXCEL_INIT",
                                             $"파일별 엑셀 저장 중... ({i + 1}/{total})",
                                             i,
                                             total,
                                             percentOverride:=CDbl(i) / total,
                                             force:=True,
                                             batchStartPercent:=CDbl(i) / total,
                                             batchEndPercent:=CDbl(i + 1) / total)

                Dim issueCount As Integer = ResolveSplitExportIssueCount(fileTable.Key, fileIssueCounts, InferSplitExportIssueCount(fileTable.Value))
                Dim savedPath As String = BuildSplitExportFilePath(outputFolder, fileTable.Key, featureKey, exportLocale, issueCount, featureSuffix, featureMode)
                ExcelCore.SaveXlsx(savedPath, sheetName, fileTable.Value, doAutoFit, sheetKey:=featureKey, progressKey:=progressKey, exportKind:=featureKey, exportLocale:=exportLocale)
                savedCount += 1
            Next

            ExcelProgressReporter.Report(progressKey, "DONE", "파일별 엑셀 저장 완료", savedCount, total, 1.0R, True)
            Return savedCount
        End Function

        Private Function SaveSplitMultiSheetTables(outputFolder As String,
                                                   featureKey As String,
                                                   featureSuffix As String,
                                                   fileWorkbooks As IList(Of KeyValuePair(Of String, List(Of KeyValuePair(Of String, DataTable)))),
                                                   doAutoFit As Boolean,
                                                   excelMode As String,
                                                   exportLocale As String,
                                                   Optional progressKey As String = "hub:multi-progress",
                                                   Optional sheetKeyOverride As String = Nothing,
                                                   Optional fileIssueCounts As IDictionary(Of String, Integer) = Nothing,
                                                   Optional featureMode As String = Nothing) As Integer
            If fileWorkbooks Is Nothing OrElse fileWorkbooks.Count = 0 Then Return 0

            Dim savedCount As Integer = 0
            Dim total As Integer = Math.Max(1, fileWorkbooks.Count)
            ExcelProgressReporter.Reset(progressKey)

            For i As Integer = 0 To fileWorkbooks.Count - 1
                Dim workbookItem = fileWorkbooks(i)
                Dim sheets = workbookItem.Value
                If sheets Is Nothing OrElse sheets.Count = 0 Then Continue For

                ExcelProgressReporter.Report(progressKey,
                                             "EXCEL_INIT",
                                             $"파일별 엑셀 저장 중... ({i + 1}/{total})",
                                             i,
                                             total,
                                             percentOverride:=CDbl(i) / total,
                                             force:=True,
                                             batchStartPercent:=CDbl(i) / total,
                                             batchEndPercent:=CDbl(i + 1) / total)

                Dim issueCount As Integer = ResolveSplitExportIssueCount(workbookItem.Key, fileIssueCounts, InferSplitExportIssueCount(sheets))
                Dim savedPath As String = BuildSplitExportFilePath(outputFolder, workbookItem.Key, featureKey, exportLocale, issueCount, featureSuffix, featureMode)
                ExcelCore.SaveXlsxMulti(savedPath, sheets, doAutoFit, progressKey, sheetKeyOverride:=sheetKeyOverride, exportKind:=featureKey, exportLocale:=exportLocale)
                savedCount += 1
            Next

            ExcelProgressReporter.Report(progressKey, "DONE", "파일별 엑셀 저장 완료", savedCount, total, 1.0R, True)
            Return savedCount
        End Function

        Private Sub SendSplitExportCompleted(outputFolder As String, savedCount As Integer)
            SendToWeb("hub:multi-exported", New With {
                .ok = True,
                .path = outputFolder,
                .kind = "folder",
                .message = $"{savedCount}개 엑셀 파일을 저장했습니다."
            })
        End Sub

        Private Sub ExportConnector(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
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
            Dim splitByFile As Boolean = Not String.IsNullOrWhiteSpace(outputFolder)

            If splitByFile OrElse (fileList IsNot Nothing AndAlso fileList.Count >= 2) Then
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

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
                    Dim issueCount As Integer = rowsForFile.Count

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
                    SetSplitExportIssueCount(fileIssueCounts, baseName, issueCount)
                Next

                If splitByFile Then
                    Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "connector", "파라미터연속성검토", "Connector Diagnostics", sheetList, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                    If savedCount <= 0 Then
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                    Else
                        SendSplitExportCompleted(outputFolder, savedCount)
                    End If
                    Return
                End If

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="connector", exportKind:="connector", exportLocale:=exportLocale)
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
                saved = ExcelCore.PickAndSaveXlsx("Connector Diagnostics", table, defaultFileName, doAutoFit, "hub:multi-progress", "connector", exportLocale:=exportLocale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportTapAlign(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiTapAlignRows, New List(Of Dictionary(Of String, Object))())
            If rows.Count = 0 Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "탭/분기 축 결과가 없습니다."})
                Return
            End If

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

            If fileList.Count = 0 Then
                For Each r In rows
                    Dim fileName As String = ""
                    Try
                        If r IsNot Nothing AndAlso r.ContainsKey("File") AndAlso r("File") IsNot Nothing Then
                            fileName = r("File").ToString()
                        End If
                    Catch
                        fileName = ""
                    End Try
                    If String.IsNullOrWhiteSpace(fileName) Then Continue For
                    If seen.Add(fileName) Then fileList.Add(fileName)
                Next
            End If

            Dim extras = If(_multiTapAlignExtras, New List(Of String)())
            Dim unit = NormalizeTapAlignUnit(If(_multiTapAlignUnit, "mm"))
            Dim locale = NormalizeTapAlignExportLocale(exportLocale)

            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim defaultFileName As String
            If requestedCount >= 2 Then
                defaultFileName = $"TapAlign_Selected {requestedCount} Files.xlsx"
            Else
                defaultFileName = $"TapAlign_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Dim saved As String = ""
            Dim splitByFile As Boolean = Not String.IsNullOrWhiteSpace(outputFolder)
            If splitByFile OrElse fileList.Count >= 2 Then
                Dim sheetList As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In fileList
                    Dim baseName As String = ""
                    Try
                        baseName = System.IO.Path.GetFileNameWithoutExtension(fileName)
                    Catch
                        baseName = fileName
                    End Try
                    If String.IsNullOrWhiteSpace(baseName) Then baseName = fileName

                    Dim rowsForFile As List(Of Dictionary(Of String, Object)) =
                        rows.Where(Function(r)
                                       If r Is Nothing Then Return False

                                       Dim rowFile As String = ""
                                       Try
                                           If r.ContainsKey("File") AndAlso r("File") IsNot Nothing Then rowFile = r("File").ToString()
                                       Catch
                                           rowFile = ""
                                       End Try

                                       If String.IsNullOrWhiteSpace(rowFile) Then Return False

                                       Dim rowBase As String = rowFile
                                       Try
                                           rowBase = System.IO.Path.GetFileNameWithoutExtension(rowFile)
                                       Catch
                                           rowBase = rowFile
                                       End Try

                                        Return String.Equals(rowFile, fileName, StringComparison.OrdinalIgnoreCase) _
                                            OrElse String.Equals(rowFile, baseName, StringComparison.OrdinalIgnoreCase) _
                                            OrElse String.Equals(rowBase, baseName, StringComparison.OrdinalIgnoreCase)
                                    End Function).ToList()
                    Dim issueCount As Integer = rowsForFile.Where(Function(item) IsTapAlignIssueRow(item)).Count()

                    Dim table = BuildTapAlignDataTable(rowsForFile, unit, extras, locale)
                    Dim headers = table.Columns.Cast(Of DataColumn)().Select(Function(col) col.ColumnName).ToList()
                    If table.Rows.Count > 0 AndAlso Not ValidateSchema(table, headers) Then
                        Throw New InvalidOperationException("스키마 검증 실패: TapAlign")
                    End If

                    sheetList.Add(New KeyValuePair(Of String, DataTable)(baseName, table))
                    SetSplitExportIssueCount(fileIssueCounts, baseName, issueCount)
                Next

                If splitByFile Then
                    Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "tapalign", "탭분기축틀어짐검토", "Tap Alignment", sheetList, doAutoFit, excelMode, locale, fileIssueCounts:=fileIssueCounts)
                    If savedCount <= 0 Then
                        SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                    Else
                        SendSplitExportCompleted(outputFolder, savedCount)
                    End If
                    Return
                End If

                saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, defaultFileName, doAutoFit, "hub:multi-progress", sheetKeyOverride:="tapalign", exportKind:="tapalign", exportLocale:=locale)
            Else
                Dim table = BuildTapAlignDataTable(rows, unit, extras, locale)
                Dim headers = table.Columns.Cast(Of DataColumn)().Select(Function(col) col.ColumnName).ToList()
                If table.Rows.Count > 0 AndAlso Not ValidateSchema(table, headers) Then
                    Throw New InvalidOperationException("스키마 검증 실패: TapAlign")
                End If

                saved = ExcelCore.PickAndSaveXlsx("Tap Alignment", table, defaultFileName, doAutoFit, "hub:multi-progress", "tapalign", exportLocale:=locale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportDupClash(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim dupRows = If(_multiDupRows, New List(Of Exports.DupRowDto)())
            Dim clashRows = If(_multiClashRows, New List(Of Exports.DupRowDto)())
            Dim clashPairs = If(_multiClashPairs, New List(Of Exports.PairRowDto)())
            Dim mode As String = GetCurrentMultiDupClashMode()
            Dim isClashMode As Boolean = String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase)

            Dim requestedCount As Integer = GetRequestedMultiFileCount()
            Dim defaultFileName As String
            If requestedCount >= 2 Then
                defaultFileName = $"{If(isClashMode, "SelfClash", "Duplicate")}_Selected {requestedCount} Files.xlsx"
            Else
                defaultFileName = $"{If(isClashMode, "SelfClash", "Duplicate")}_{Date.Now:yyyyMMdd_HHmm}.xlsx"
            End If

            Dim extraParamNames As List(Of String) = New List(Of String)()
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                extraParamNames = ParseExtraParams(_multiRequest.Common.ExtraParams)
            End If

            Dim splitByFile As Boolean = Not String.IsNullOrWhiteSpace(outputFolder)
            If splitByFile Then
                Dim orderedNames = BuildOrderedMultiFileNames(
                    dupRows.Select(Function(item) If(item Is Nothing, "", item.FileName)),
                    clashRows.Select(Function(item) If(item Is Nothing, "", item.FileName)),
                    clashPairs.Select(Function(item) If(item Is Nothing, "", item.FileName)))
                Dim savedCount As Integer = 0

                ExcelProgressReporter.Reset("hub:multi-progress")
                For i As Integer = 0 To orderedNames.Count - 1
                    Dim fileName = orderedNames(i)
                    If String.IsNullOrWhiteSpace(fileName) Then Continue For

                    ExcelProgressReporter.Report("hub:multi-progress",
                                                 "EXCEL_INIT",
                                                 $"파일별 엑셀 저장 중... ({i + 1}/{Math.Max(1, orderedNames.Count)})",
                                                 i,
                                                 Math.Max(1, orderedNames.Count),
                                                 percentOverride:=CDbl(i) / Math.Max(1, orderedNames.Count),
                                                 force:=True,
                                                 batchStartPercent:=CDbl(i) / Math.Max(1, orderedNames.Count),
                                                 batchEndPercent:=CDbl(i + 1) / Math.Max(1, orderedNames.Count))
                    Dim savedPath As String = ""
                    Dim issueCount As Integer = 0

                    If isClashMode Then
                        Dim perFilePairs = clashPairs.
                            Where(Function(item) item IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(item.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                            ToList()
                        Dim perFileClashRows = clashRows.
                            Where(Function(item) item IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(item.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                            ToList()
                        issueCount = If(perFilePairs.Count > 0,
                                        perFilePairs.Count,
                                        perFileClashRows.Count)
                        savedPath = BuildSplitExportFilePath(outputFolder, fileName, "dupclash", exportLocale, issueCount, "자체간섭검토", "clash")
                        If perFilePairs.Count > 0 OrElse perFileClashRows.Count = 0 Then
                            Global.KKY_Tool_Revit.Exports.DuplicateExport.ExportPairs(savedPath, perFilePairs, doAutoFit, "hub:multi-progress", "Self Clash (Batch)", extraParamNames, exportLocale)
                        Else
                            Global.KKY_Tool_Revit.Exports.DuplicateExport.Export(savedPath, perFileClashRows, doAutoFit, "hub:multi-progress", "Self Clash (Batch)", extraParamNames, exportLocale)
                        End If
                    Else
                        Dim perFileDupRows = dupRows.
                            Where(Function(item) item IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(item.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                            ToList()
                        Dim exportDupRows As List(Of Exports.DupRowDto) =
                            If(perFileDupRows.Count > 0,
                               perFileDupRows,
                               BuildMultiDuplicatePlaceholderRowsForFile(fileName, exportLocale))
                        issueCount = perFileDupRows.Count
                        savedPath = BuildSplitExportFilePath(outputFolder, fileName, "dupclash", exportLocale, issueCount, "중복검토", "duplicate")
                        Global.KKY_Tool_Revit.Exports.DuplicateExport.Export(savedPath, exportDupRows, doAutoFit, "hub:multi-progress", "Duplicates (Batch)", extraParamNames, exportLocale)
                    End If

                    savedCount += 1
                Next

                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            ExcelProgressReporter.Reset("hub:multi-progress")
            Dim saved As String = ""
            If isClashMode Then
                If clashPairs.Count > 0 OrElse clashRows.Count = 0 Then
                    saved = Global.KKY_Tool_Revit.Exports.DuplicateExport.SavePairsWithDefaultName(clashPairs,
                                                                                                  defaultFileName,
                                                                                                  doAutoFit,
                                                                                                  "hub:multi-progress",
                                                                                                  "Self Clash (Batch)",
                                                                                                  extraParamNames,
                                                                                                  exportLocale)
                Else
                    saved = Global.KKY_Tool_Revit.Exports.DuplicateExport.SaveWithDefaultName(clashRows,
                                                                                              defaultFileName,
                                                                                              doAutoFit,
                                                                                              "hub:multi-progress",
                                                                                              "Self Clash (Batch)",
                                                                                              extraParamNames,
                                                                                              exportLocale)
                End If
            Else
                Dim exportDupRows = BuildMultiDuplicateExportRowsWithPlaceholders(dupRows, exportLocale)
                saved = Global.KKY_Tool_Revit.Exports.DuplicateExport.SaveWithDefaultName(exportDupRows,
                                                                                          defaultFileName,
                                                                                          doAutoFit,
                                                                                          "hub:multi-progress",
                                                                                          "Duplicates (Batch)",
                                                                                          extraParamNames,
                                                                                          exportLocale)
            End If

            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Return
            End If

            SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
        End Sub


        Private Sub ExportSegmentPms(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
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

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim orderedNames = BuildOrderedMultiFileNames(
                    classRows.Select(Function(row) ReadField(row, "File")),
                    sizeRows.Select(Function(row) ReadField(row, "FileName")),
                    routingRows.Select(Function(row) ReadField(row, "File")))
                Dim workbooks As New List(Of KeyValuePair(Of String, List(Of KeyValuePair(Of String, DataTable))))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim perFileClassRows = classRows.
                        Where(Function(row) String.Equals(GetSafeMultiFileName(ReadField(row, "File")), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    Dim perFileSizeRows = sizeRows.
                        Where(Function(row) String.Equals(GetSafeMultiFileName(ReadField(row, "FileName")), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    Dim perFileRoutingRows = routingRows.
                        Where(Function(row) String.Equals(GetSafeMultiFileName(ReadField(row, "File")), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()

                    Dim perFileClassTable = BuildTableFromRows(classHeaders, perFileClassRows)
                    Dim perFileSizeTable = BuildTableFromRows(sizeHeaders, perFileSizeRows)
                    Dim perFileRoutingTable = BuildTableFromRows(routingHeaders, perFileRoutingRows)
                    ExcelCore.EnsureNoDataRow(perFileClassTable, "오류가 없습니다.")
                    ExcelCore.EnsureNoDataRow(perFileSizeTable, "오류가 없습니다.")
                    ExcelCore.EnsureNoDataRow(perFileRoutingTable, "오류가 없습니다.")

                    Dim perFileSheets As New List(Of KeyValuePair(Of String, DataTable))()
                    perFileSheets.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", perFileClassTable))
                    perFileSheets.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", perFileSizeTable))
                    perFileSheets.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", perFileRoutingTable))
                    workbooks.Add(New KeyValuePair(Of String, List(Of KeyValuePair(Of String, DataTable)))(fileName, perFileSheets))
                    SetSplitExportIssueCount(fileIssueCounts, fileName, perFileClassRows.Count + perFileSizeRows.Count + perFileRoutingRows.Count)
                Next

                Dim savedCount = SaveSplitMultiSheetTables(outputFolder, "pms", "SegmentPms검토", workbooks, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            If totalRowsCount = 0 Then
                sheetList.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", classTable))
                sheetList.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", sizeTable))
                sheetList.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", routingTable))
            Else
                If classTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Pipe Segment Class검토", classTable))
                If sizeTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("PMS vs Segment Size검토", sizeTable))
                If routingTable.Rows.Count > 0 Then sheetList.Add(New KeyValuePair(Of String, DataTable)("Routing Class검토", routingTable))
            End If

            Dim saved = ExcelCore.PickAndSaveXlsxMulti(sheetList, $"SegmentPms_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", exportKind:="pms", exportLocale:=exportLocale)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportGuid(excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim doAutoFit As Boolean = String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase)
            If _multiGuidProject Is Nothing Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "GUID 결과가 없습니다."})
                Return
            End If

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim projectSource = _multiGuidProject
                Dim familySource = _multiGuidFamilyDetail
                Dim orderedNames = BuildOrderedMultiFileNames(
                    projectSource.Rows.Cast(Of DataRow)().Select(Function(row) ReadDataRowField(row, "RvtName")),
                    If(familySource Is Nothing,
                       Enumerable.Empty(Of String)(),
                       familySource.Rows.Cast(Of DataRow)().Select(Function(row) ReadDataRowField(row, "RvtName"))))
                Dim workbooks As New List(Of KeyValuePair(Of String, List(Of KeyValuePair(Of String, DataTable))))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim projectTable = projectSource.Clone()
                    Dim projectRowsForFile = projectSource.Rows.Cast(Of DataRow)().
                        Where(Function(item) String.Equals(GetSafeMultiFileName(ReadDataRowField(item, "RvtName")), fileName, StringComparison.OrdinalIgnoreCase))
                    For Each row In projectRowsForFile
                        projectTable.ImportRow(row)
                    Next

                    Dim perFileSheets As New List(Of KeyValuePair(Of String, DataTable))()
                    perFileSheets.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", GuidAuditService.PrepareExportTable(projectTable, 1)))

                    If familySource IsNot Nothing Then
                        Dim familyTable = familySource.Clone()
                        Dim familyRowsForFile = familySource.Rows.Cast(Of DataRow)().
                            Where(Function(item) String.Equals(GetSafeMultiFileName(ReadDataRowField(item, "RvtName")), fileName, StringComparison.OrdinalIgnoreCase))
                        For Each row In familyRowsForFile
                            familyTable.ImportRow(row)
                        Next

                        If familyTable.Rows.Count > 0 Then
                            perFileSheets.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", GuidAuditService.PrepareExportTable(familyTable, 2)))
                        End If
                        SetSplitExportIssueCount(fileIssueCounts, fileName, projectRowsForFile.Count() + familyRowsForFile.Count())
                    Else
                        SetSplitExportIssueCount(fileIssueCounts, fileName, projectRowsForFile.Count())
                    End If

                    workbooks.Add(New KeyValuePair(Of String, List(Of KeyValuePair(Of String, DataTable)))(fileName, perFileSheets))
                Next

                Dim savedCount = SaveSplitMultiSheetTables(outputFolder, "guid", "파라미터GUID검토", workbooks, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            sheets.Add(New KeyValuePair(Of String, DataTable)("RVT 검토결과", GuidAuditService.PrepareExportTable(_multiGuidProject, 1)))
            If _multiGuidFamilyDetail IsNot Nothing Then
                sheets.Add(New KeyValuePair(Of String, DataTable)("Family 검토결과", GuidAuditService.PrepareExportTable(_multiGuidFamilyDetail, 2)))
            End If
            Dim saved = GuidAuditService.ExportMulti(sheets, excelMode, "hub:multi-progress", exportLocale)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportFamilyLink(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
            Dim rows = If(_multiFamilyLinkRows, New List(Of FamilyLinkAuditRow)())
            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim orderedNames = BuildOrderedMultiFileNames(rows.Select(Function(item) If(item Is Nothing, "", item.FileName)))
                Dim fileTables As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim perFileRows = rows.
                        Where(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    If perFileRows.Count = 0 Then Continue For

                    Dim table = FamilyLinkAuditExport.ToDataTable(perFileRows)
                    ExcelCore.EnsureMessageRow(table, "오류가 없습니다.")
                    fileTables.Add(New KeyValuePair(Of String, DataTable)(fileName, table))
                    SetSplitExportIssueCount(fileIssueCounts, fileName, perFileRows.Where(Function(item) item IsNot Nothing AndAlso Not String.Equals(If(item.Issue, ""), FamilyLinkAuditIssue.OK.ToString(), StringComparison.OrdinalIgnoreCase)).Count())
                Next

                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "familylink", "패밀리공유파라미터연동검토", "FamilyLinkAudit", fileTables, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            ExcelProgressReporter.Reset("hub:multi-progress")
            Dim saved = FamilyLinkAuditExport.Export(rows,
                                                     fastExport:=String.Equals(excelMode, "fast", StringComparison.OrdinalIgnoreCase),
                                                     autoFit:=doAutoFit,
                                                     progressChannel:="hub:multi-progress",
                                                     exportLocale:=exportLocale)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportPoints(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
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

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim orderedNames = BuildOrderedMultiFileNames(pointRows.Select(Function(item) If(item Is Nothing, "", item.File)))
                Dim fileTables As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim perFileRows = rows.
                        Where(Function(item) String.Equals(GetSafeMultiFileName(ReadField(item, "File")), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    Dim tablePerFile = BuildPointTable(headers, perFileRows)
                    If Not ValidateSchema(tablePerFile, headers) Then Throw New InvalidOperationException("스키마 검증 실패: Points")
                    fileTables.Add(New KeyValuePair(Of String, DataTable)(fileName, tablePerFile))
                    SetSplitExportIssueCount(fileIssueCounts, fileName, 0)
                Next

                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "points", "Point좌표추출", "Points", fileTables, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim table = BuildPointTable(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: Points")
            Dim saved = ExcelCore.PickAndSaveXlsx("Points", table, $"Points_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", "points", exportLocale:=exportLocale)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Sub ExportLinkWorkset(doAutoFit As Boolean, excelMode As String, Optional exportLocale As String = "ko", Optional outputFolder As String = Nothing)
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

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim orderedNames = BuildOrderedMultiFileNames(rows.Select(Function(item) If(item Is Nothing, "", item.HostFileName)))
                Dim fileTables As New List(Of KeyValuePair(Of String, DataTable))()
                Dim fileIssueCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                For Each fileName In orderedNames
                    Dim perFileRows = rows.
                    Where(Function(item) item IsNot Nothing AndAlso String.Equals(GetSafeMultiFileName(item.HostFileName), fileName, StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    If perFileRows.Count = 0 Then Continue For

                    Dim tablePerFile = BuildLinkWorksetTable(headers, perFileRows)
                    If Not ValidateSchema(tablePerFile, headers) Then Throw New InvalidOperationException("스키마 검증 실패: LinkWorkset")
                    fileTables.Add(New KeyValuePair(Of String, DataTable)(fileName, tablePerFile))
                    SetSplitExportIssueCount(fileIssueCounts, fileName, perFileRows.Where(Function(item) IsLinkWorksetIssue(item)).Count())
                Next

                Dim savedCount = SaveSplitSingleSheetTables(outputFolder, "linkworkset", "링크기본웍셋점검적용", "LinkWorkset", fileTables, doAutoFit, excelMode, exportLocale, fileIssueCounts:=fileIssueCounts)
                If savedCount <= 0 Then
                    SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
                Else
                    SendSplitExportCompleted(outputFolder, savedCount)
                End If
                Return
            End If

            Dim table = BuildLinkWorksetTable(headers, rows)
            If Not ValidateSchema(table, headers) Then Throw New InvalidOperationException("스키마 검증 실패: LinkWorkset")
            Dim saved = ExcelCore.PickAndSaveXlsx("LinkWorkset", table, $"LinkWorkset_{Date.Now:yyyyMMdd_HHmm}.xlsx", doAutoFit, "hub:multi-progress", "linkworkset", exportLocale:=exportLocale)
            If String.IsNullOrWhiteSpace(saved) Then
                SendToWeb("hub:multi-exported", New With {.ok = False, .message = "엑셀 저장이 취소되었습니다."})
            Else
                SendToWeb("hub:multi-exported", New With {.ok = True, .path = saved})
            End If
        End Sub

        Private Shared Sub ResetMultiCaches()
            _multiConnectorRows = Nothing
            _multiConnectorExtras = Nothing
            _multiTapAlignRows = Nothing
            _multiTapAlignExtras = Nothing
            _multiTapAlignUnit = "mm"
            _multiTapAlignLocale = "ko"
            _multiDupRows = Nothing
            _multiDupTargetCounts = Nothing
            _multiClashRows = Nothing
            _multiClashPairs = Nothing
            _multiFloorInfoRows = Nothing
            _multiFloorInfoFileSummaries = Nothing
            _multiFloorInfoWarnings = Nothing
            _multiFamilySuitabilityRows = Nothing
            _multiFamilySuitabilityFileSummaries = Nothing
            _multiFamilySuitabilityWarnings = Nothing
            _multiWorksetAssignmentRows = Nothing
            _multiWorksetAssignmentFileSummaries = Nothing
            _multiParameterDuplicationRows = Nothing
            _multiParameterDuplicationFileSummaries = Nothing
            _multiParameterMissingRows = Nothing
            _multiParameterMissingFileSummaries = Nothing
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
                If String.IsNullOrWhiteSpace(req.Common.ExtraParams) Then
                    req.Common.ExtraParams = SafeStr(GetDictValue(commonDict, "extraParamsText"))
                End If
                req.Common.TargetFilter = SafeStr(GetDictValue(commonDict, "targetFilter"))
                If String.IsNullOrWhiteSpace(req.Common.TargetFilter) Then
                    req.Common.TargetFilter = SafeStr(GetDictValue(commonDict, "targetFilterText"))
                End If
                req.Common.ExcludeTargetFilter = SafeStr(GetDictValue(commonDict, "excludeTargetFilter"))
                If String.IsNullOrWhiteSpace(req.Common.ExcludeTargetFilter) Then
                    req.Common.ExcludeTargetFilter = SafeStr(GetDictValue(commonDict, "excludeTargetFilterText"))
                End If
                req.Common.ExcludeEndDummy = False
                req.Common.IncludePointXY = ToBool(GetDictValue(commonDict, "includePointXY"))
                req.Common.IncludeLinearMetrics = ToBool(GetDictValue(commonDict, "includeLinearMetrics"))
            End If

            Dim featuresObj As Object = Nothing
            If pd.TryGetValue("features", featuresObj) Then
                Dim fd = ToDict(featuresObj)
                req.Connector = ParseConnector(fd)
                req.FloorInfo = ParseFloorInfo(fd)
                req.FamilySuitability = ParseFamilySuitability(fd)
                req.TapAlign = ParseTapAlign(fd)
                req.DupClash = ParseDupClash(fd)
                req.WorksetAssignment = ParseWorksetAssignment(fd)
                req.ProjectParameterDuplication = ParseProjectParameterDuplication(fd)
                req.ParameterMissing = ParseParameterMissing(fd)
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
            opt.ExcludeEndDummy = ToBool(GetDictValue(d, "excludeEndDummy"))
            opt.IncludePointXY = ToBool(GetDictValue(d, "includePointXY"))
            opt.IncludeLinearMetrics = ToBool(GetDictValue(d, "includeLinearMetrics"))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "inch"
            If String.IsNullOrWhiteSpace(opt.Param) Then opt.Param = "Comments"
            Return opt
        End Function

        Private Function ParseTapAlign(fd As Dictionary(Of String, Object)) As MultiTapAlignOptions
            Dim opt As New MultiTapAlignOptions()
            Dim obj = GetDictValue(fd, "tapalign")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Tol = ToDouble(GetDictValue(d, "tol"), 0.5R)
            opt.Unit = NormalizeTapAlignUnit(SafeStr(GetDictValue(d, "unit")))
            opt.Domain = NormalizeTapAlignDomain(SafeStr(GetDictValue(d, "domain")))
            opt.FeatureTargetFilter = SafeStr(GetDictValue(d, "featureTargetFilter"))
            opt.ExportLocale = NormalizeTapAlignExportLocale(SafeStr(GetDictValue(d, "exportLocale")))
            If String.IsNullOrWhiteSpace(opt.Unit) Then opt.Unit = "mm"
            If String.IsNullOrWhiteSpace(opt.Domain) Then opt.Domain = "all"
            If opt.FeatureTargetFilter Is Nothing Then opt.FeatureTargetFilter = String.Empty
            If String.IsNullOrWhiteSpace(opt.ExportLocale) Then opt.ExportLocale = "ko"
            Return opt
        End Function

        Private Function ParseDupClash(fd As Dictionary(Of String, Object)) As MultiDupClashOptions
            Dim opt As New MultiDupClashOptions()
            Dim obj = GetDictValue(fd, "dupclash")
            Dim d = ToDict(obj)
            opt.Enabled = ToBool(GetDictValue(d, "enabled"))
            opt.Mode = NormalizeMultiDupClashMode(SafeStr(GetDictValue(d, "mode")))
            opt.TolFeet = ToDouble(GetDictValue(d, "tolFeet"), 1.0R / 64.0R)
            If opt.TolFeet <= 0 Then opt.TolFeet = 1.0R / 64.0R
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
            For Each o In EnumeratePayloadItems(rawTargets)
                Dim td = ToDict(o)
                Dim name = NormalizeWrappedQuotesText(SafeStr(GetDictValue(td, "name"))).Trim()
                Dim guidStr = NormalizeWrappedQuotesText(SafeStr(GetDictValue(td, "guid"))).Trim()
                Dim g As Guid
                If Not String.IsNullOrWhiteSpace(name) AndAlso Guid.TryParse(guidStr, g) Then
                    targets.Add(New FamilyLinkTargetParam With {.Name = name, .Guid = g})
                End If
            Next
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
            Return req.Connector.Enabled OrElse req.FloorInfo.Enabled OrElse req.FamilySuitability.Enabled OrElse req.TapAlign.Enabled OrElse req.DupClash.Enabled OrElse req.WorksetAssignment.Enabled OrElse req.ProjectParameterDuplication.Enabled OrElse req.ParameterMissing.Enabled OrElse req.Pms.Enabled OrElse req.Guid.Enabled OrElse req.FamilyLink.Enabled OrElse req.Points.Enabled OrElse req.LinkWorkset.Enabled
        End Function

        Private Shared Function CountEnabledFeatures(req As MultiRunRequest) As Integer
            If req Is Nothing Then Return 0
            Dim count As Integer = 0
            If req.Connector.Enabled Then count += 1
            If req.FloorInfo.Enabled Then count += 1
            If req.FamilySuitability.Enabled Then count += 1
            If req.TapAlign.Enabled Then count += 1
            If req.DupClash.Enabled Then count += 1
            If req.WorksetAssignment.Enabled Then count += 1
            If req.ProjectParameterDuplication.Enabled Then count += 1
            If req.ParameterMissing.Enabled Then count += 1
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

            If _multiRequest.FamilySuitability IsNot Nothing AndAlso _multiRequest.FamilySuitability.Enabled Then
                summaries("familysuitability") = BuildFamilySuitabilityMultiSummary()
            End If

            If _multiRequest.TapAlign IsNot Nothing AndAlso _multiRequest.TapAlign.Enabled Then
                summaries("tapalign") = BuildTapAlignMultiSummary()
            End If

            If _multiRequest.DupClash IsNot Nothing AndAlso _multiRequest.DupClash.Enabled Then
                summaries("dupclash") = BuildDupClashMultiSummary()
            End If

            If _multiRequest.WorksetAssignment IsNot Nothing AndAlso _multiRequest.WorksetAssignment.Enabled Then
                summaries("worksetassignment") = BuildWorksetAssignmentMultiSummary()
            End If

            If _multiRequest.ProjectParameterDuplication IsNot Nothing AndAlso _multiRequest.ProjectParameterDuplication.Enabled Then
                summaries("parameterduplication") = BuildProjectParameterDuplicationMultiSummary()
            End If

            If _multiRequest.ParameterMissing IsNot Nothing AndAlso _multiRequest.ParameterMissing.Enabled Then
                summaries("parametermissing") = BuildParameterMissingMultiSummary()
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

        Private Function BuildDupClashMultiSummary() As Object
            Dim mode As String = GetCurrentMultiDupClashMode()
            Dim modeLabel As String = ResolveMultiDupClashModeLabel(mode)
            Dim dupRows = If(_multiDupRows, New List(Of Exports.DupRowDto)())
            Dim clashRows = If(_multiClashRows, New List(Of Exports.DupRowDto)())
            Dim clashPairs = If(_multiClashPairs, New List(Of Exports.PairRowDto)())
            Dim isClashMode As Boolean = String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase)
            Dim clashCount As Integer = If(clashPairs.Count > 0, clashPairs.Count, clashRows.Count)
            Dim dupGroups As Integer = dupRows.
                Select(Function(r) SafeStr(If(r Is Nothing, "", r.GroupKey))).
                Where(Function(g) Not String.IsNullOrWhiteSpace(g)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                Count()
            Dim clashFiles As Integer = clashPairs.
                Select(Function(r) SafeStr(If(r Is Nothing, "", r.FileName))).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                Count()
            If clashFiles = 0 Then
                clashFiles = clashRows.
                    Select(Function(r) SafeStr(If(r Is Nothing, "", r.FileName))).
                    Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    Count()
            End If

            Dim extraParamCount As Integer = 0
            Dim filterLabel As String = "포함 필터 없음"
            Dim excludeFilterLabel As String = "제외 필터 없음"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.Common IsNot Nothing Then
                extraParamCount = ParseExtraParams(_multiRequest.Common.ExtraParams).Count
                If Not String.IsNullOrWhiteSpace(_multiRequest.Common.TargetFilter) Then
                    filterLabel = _multiRequest.Common.TargetFilter.Trim()
                End If
                If Not String.IsNullOrWhiteSpace(_multiRequest.Common.ExcludeTargetFilter) Then
                    excludeFilterLabel = _multiRequest.Common.ExcludeTargetFilter.Trim()
                End If
            End If

            Dim lines As New List(Of String) From {
                $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                $"검토 모드: {modeLabel}"
            }
            If isClashMode Then
                lines.Add($"자체간섭 결과 건수: {clashCount}건")
                lines.Add($"자체간섭 검출 파일 수: {clashFiles}개")
            Else
                lines.Add($"중복 결과 행 수: {dupRows.Count}행")
                lines.Add($"중복 그룹 수: {dupGroups}개")
            End If
            lines.Add($"공통 포함 필터: {filterLabel}")
            lines.Add($"공통 제외 필터: {excludeFilterLabel} / 추가 파라미터 {extraParamCount}개")

            Return New With {
                .key = "dupclash",
                .label = modeLabel,
                .lines = lines.ToArray(),
                .fileSummaries = BuildDupClashFileSummaries()
            }
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

        Private Function BuildTapAlignMultiSummary() As Object
            Dim rows = If(_multiTapAlignRows, New List(Of Dictionary(Of String, Object))())
            Dim issueCount As Integer = rows.Where(Function(r) IsTapAlignIssueRow(r)).Count()
            Dim issueFiles As Integer = rows.
                Where(Function(r) IsTapAlignIssueRow(r)).
                Select(Function(r) GetSafeMultiFileName(ReadField(r, "File"))).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                Count()
            Dim extraCount As Integer = If(_multiTapAlignExtras, New List(Of String)()).Count
            Dim tolText As String = "0.5"
            Dim unitText As String = NormalizeTapAlignUnit(_multiTapAlignUnit)
            Dim scopeText As String = "배관 + 덕트"
            If _multiRequest IsNot Nothing AndAlso _multiRequest.TapAlign IsNot Nothing Then
                tolText = _multiRequest.TapAlign.Tol.ToString()
                unitText = NormalizeTapAlignUnit(_multiRequest.TapAlign.Unit)
                scopeText = ResolveTapAlignDomainLabel(_multiRequest.TapAlign.Domain)
            End If

            Return New With {
                .key = "tapalign",
                .label = "탭/분기 축 틀어짐 검토",
                .lines = New String() {
                    $"선택 파일 수: {GetRequestedMultiFileCount()}개",
                    $"오류 행 수: {issueCount}행",
                    $"오류 파일 수: {issueFiles}개",
                    $"허용범위: {tolText} {unitText}",
                    $"검토 범위: {scopeText}",
                    $"추가 추출 컬럼 수: {extraCount}개"
                },
                .fileSummaries = BuildTapAlignFileSummaries(rows)
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

        Private Function BuildTapAlignFileSummaries(rows As IList(Of Dictionary(Of String, Object))) As List(Of Object)
            Dim sourceRows = If(rows, New List(Of Dictionary(Of String, Object))())
            Dim orderedNames = BuildOrderedMultiFileNames(sourceRows.Select(Function(r) ReadField(r, "File")))
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim perFileRows = sourceRows.
                    Where(Function(r) String.Equals(GetSafeMultiFileName(ReadField(r, "File")), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim issueRows = perFileRows.Where(Function(r) IsTapAlignIssueRow(r)).ToList()
                Dim runReason As String = ""
                Dim statusText As String = GetMultiRunItemStatus(fileName, runReason)
                Dim maxDistance As Double = 0.0R
                If issueRows.Count > 0 Then
                    maxDistance = issueRows.
                        Select(Function(r) ToDouble(ReadField(r, "DistanceFromCenter"), 0.0R)).
                        DefaultIfEmpty(0.0R).
                        Max()
                End If
                Dim reason As String = runReason
                If String.IsNullOrWhiteSpace(reason) Then
                    reason = If(issueRows.Count > 0,
                                $"최대 이탈거리 {Math.Round(maxDistance, 3)} {_multiTapAlignUnit}",
                                "오류 없음")
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = perFileRows.Count,
                    .issues = issueRows.Count,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
                })
            Next

            Return result
        End Function

        Private Function BuildDupClashFileSummaries() As List(Of Object)
            Dim mode As String = GetCurrentMultiDupClashMode()
            Dim isClashMode As Boolean = String.Equals(mode, "clash", StringComparison.OrdinalIgnoreCase)
            Dim dupRows = If(_multiDupRows, New List(Of Exports.DupRowDto)())
            Dim clashRows = If(_multiClashRows, New List(Of Exports.DupRowDto)())
            Dim clashPairs = If(_multiClashPairs, New List(Of Exports.PairRowDto)())
            Dim orderedNames As List(Of String)
            If isClashMode Then
                orderedNames = BuildOrderedMultiFileNames(
                    clashRows.Select(Function(r) If(r Is Nothing, "", r.FileName)),
                    clashPairs.Select(Function(r) If(r Is Nothing, "", r.FileName)))
            Else
                orderedNames = BuildOrderedMultiFileNames(
                    dupRows.Select(Function(r) If(r Is Nothing, "", r.FileName)))
            End If
            Dim result As New List(Of Object)()

            For Each fileName In orderedNames
                Dim perFileDupRows = dupRows.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(r.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim perFileClashRows = clashRows.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(r.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim perFileClashPairs = clashPairs.
                    Where(Function(r) r IsNot Nothing AndAlso String.Equals(ResolveRequestedMultiFileName(r.FileName), fileName, StringComparison.OrdinalIgnoreCase)).
                    ToList()

                Dim runReason As String = ""
                Dim statusText As String = GetMultiRunItemStatus(fileName, runReason)
                Dim clashCount As Integer = If(perFileClashPairs.Count > 0, perFileClashPairs.Count, perFileClashRows.Count)
                Dim issueCount As Integer = If(isClashMode, clashCount, perFileDupRows.Count)
                Dim reason As String = runReason
                If String.IsNullOrWhiteSpace(reason) Then
                    If issueCount > 0 Then
                        reason = If(isClashMode,
                                    $"간섭 {clashCount}건",
                                    $"중복 {perFileDupRows.Count}건")
                    Else
                        reason = "오류 없음"
                    End If
                End If

                result.Add(New With {
                    .file = fileName,
                    .total = issueCount,
                    .issues = issueCount,
                    .near = 0,
                    .status = statusText,
                    .reason = reason
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

        Private Shared Function IsTapAlignIssueRow(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            If Not String.IsNullOrWhiteSpace(ReadField(row, "ElementId")) Then Return True
            Dim message = ReadField(row, "Message")
            Return Not String.IsNullOrWhiteSpace(message)
        End Function

        Private Function BuildOrderedMultiFileNames(ParamArray nameGroups() As IEnumerable(Of String)) As List(Of String)
            Dim orderedNames As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each path In _multiRequest.RvtPaths
                    Dim safeName As String = ResolveRequestedMultiFileName(path)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            If _multiRunItems IsNot Nothing Then
                For Each item In _multiRunItems
                    If item Is Nothing Then Continue For
                    Dim safeName As String = ResolveRequestedMultiFileName(item.File)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            End If

            For Each group In nameGroups
                If group Is Nothing Then Continue For
                For Each rawName In group
                    Dim safeName As String = ResolveRequestedMultiFileName(rawName)
                    If String.IsNullOrWhiteSpace(safeName) Then Continue For
                    If seen.Add(safeName) Then orderedNames.Add(safeName)
                Next
            Next

            Return orderedNames
        End Function

        Private Function GetMultiRunItemStatus(fileName As String, ByRef reason As String) As String
            reason = ""
            Dim safeFileName As String = ResolveRequestedMultiFileName(fileName)
            If String.IsNullOrWhiteSpace(safeFileName) Then safeFileName = SafeStr(fileName)
            If _multiRunItems IsNot Nothing Then
                For Each item In _multiRunItems
                    If item Is Nothing Then Continue For
                    If String.Equals(ResolveRequestedMultiFileName(item.File), safeFileName, StringComparison.OrdinalIgnoreCase) Then
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
                fileSet.Add(ResolveRequestedMultiFileName(path))
            Next
            Return fileSet.Count
        End Function

        Private Function ResolveRequestedMultiFileName(pathOrName As String) As String
            Dim safeName As String = GetSafeMultiFileName(pathOrName)
            If String.IsNullOrWhiteSpace(safeName) Then Return SafeStr(pathOrName)

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each requestedPath In _multiRequest.RvtPaths
                    Dim requestedName As String = GetSafeMultiFileName(requestedPath)
                    If String.Equals(requestedName, safeName, StringComparison.OrdinalIgnoreCase) Then
                        Return requestedName
                    End If
                Next
            End If

            Dim normalizedName As String = NormalizeGeneratedMultiLocalSafeName(safeName)
            If String.IsNullOrWhiteSpace(normalizedName) Then Return safeName

            If _multiRequest IsNot Nothing AndAlso _multiRequest.RvtPaths IsNot Nothing Then
                For Each requestedPath In _multiRequest.RvtPaths
                    Dim requestedName As String = GetSafeMultiFileName(requestedPath)
                    If String.Equals(requestedName, normalizedName, StringComparison.OrdinalIgnoreCase) Then
                        Return requestedName
                    End If
                Next
            End If

            Return normalizedName
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

        Private Shared Function NormalizeGeneratedMultiLocalSafeName(pathOrName As String) As String
            Dim safeName As String = GetSafeMultiFileName(pathOrName)
            If String.IsNullOrWhiteSpace(safeName) Then Return String.Empty

            Dim userName As String = SafeStr(Environment.UserName).Trim()
            If String.IsNullOrWhiteSpace(userName) Then Return safeName

            Dim token As String = "_" & userName & "_"
            Dim markerIndex As Integer = safeName.LastIndexOf(token, StringComparison.OrdinalIgnoreCase)
            If markerIndex <= 0 Then Return safeName

            Dim suffix As String = safeName.Substring(markerIndex + token.Length)
            If String.IsNullOrWhiteSpace(suffix) OrElse suffix.Length < 6 Then Return safeName

            For Each ch As Char In suffix
                If Not Char.IsDigit(ch) Then Return safeName
            Next

            Return safeName.Substring(0, markerIndex)
        End Function

        Private Shared Function NormalizeMultiPath(path As String) As String
            Dim value As String = SafeStr(path).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            value = New String(value.Where(Function(ch) Not Char.IsControl(ch)).ToArray())
            value = value.Trim(""""c)
            value = MaybeUnescapeSerializedText(value)

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
            Return CalcStepProgressPercent(basePct, stepIndex, totalSteps, 1.0R)
        End Function

        Private Function CalcStepProgressPercent(basePct As Double, stepIndex As Integer, totalSteps As Integer, stepProgress As Double) As Double
            Dim perFile As Double = If(_multiTotal > 0, 1.0R / CDbl(_multiTotal), 1.0R)
            Dim stepShare As Double = perFile / CDbl(Math.Max(totalSteps, 1))
            Dim clampedStep As Double = Math.Max(0.0R, Math.Min(1.0R, stepProgress))
            Dim completedSteps As Double = Math.Max(0, stepIndex - 1)
            Dim stepPct As Double = (basePct + (stepShare * (completedSteps + clampedStep))) * 100.0R
            Return Math.Min(stepPct, 99.9R)
        End Function

        Private Function BuildMultiFamilyLinkProgressDetail(safeName As String, progressMessage As String) As String
            Dim baseName As String = If(safeName, "").Trim()
            Dim raw As String = If(progressMessage, "").Trim()
            If String.IsNullOrWhiteSpace(raw) Then Return baseName

            Dim prefix As String = baseName & " - "
            Dim messageBody As String = raw
            If Not String.IsNullOrWhiteSpace(baseName) AndAlso raw.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                messageBody = raw.Substring(prefix.Length).Trim()
            End If

            Dim marker As String = " ("
            Dim markerIndex As Integer = messageBody.LastIndexOf(marker, StringComparison.Ordinal)
            If markerIndex > 0 AndAlso messageBody.EndsWith("패밀리 검사 중", StringComparison.Ordinal) Then
                Dim familyName As String = messageBody.Substring(0, markerIndex).Trim()
                Dim countStart As Integer = markerIndex + marker.Length
                Dim countEnd As Integer = messageBody.IndexOf(")", countStart, StringComparison.Ordinal)
                If countEnd > countStart Then
                    Dim countText As String = messageBody.Substring(countStart, countEnd - countStart).Trim()
                    If Not String.IsNullOrWhiteSpace(familyName) AndAlso Not String.IsNullOrWhiteSpace(countText) Then
                        Return baseName & " - " & familyName & vbLf & countText & " 패밀리 검사 중"
                    End If
                End If
            End If

            If String.IsNullOrWhiteSpace(baseName) Then Return raw
            If raw.StartsWith(baseName, StringComparison.OrdinalIgnoreCase) Then Return raw
            Return baseName & vbLf & raw
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
                            dr(i) = BuildConnectorCommentTextForExport(r)
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
            If TryPopulateFromJsonObject(obj, res) Then Return res
            Dim t = obj.GetType()
            For Each p In t.GetProperties()
                Try
                    res(p.Name) = ConvertPayloadValue(p.GetValue(obj, Nothing))
                Catch
                End Try
            Next
            Return res
        End Function

        Private Shared Function ExtractStringList(dict As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim list As New List(Of String)()
            Dim raw = GetDictValue(dict, key)
            Dim items = EnumeratePayloadItems(raw)
            If items.Count > 0 Then
                For Each o In items
                    Dim s = NormalizeMultiPath(SafeStr(ConvertPayloadValue(o)))
                    If Not String.IsNullOrWhiteSpace(s) AndAlso Not list.Contains(s) Then
                        list.Add(s)
                    End If
                Next
            ElseIf Not IsJsonArrayValue(raw) Then
                Dim s = NormalizeMultiPath(SafeStr(ConvertPayloadValue(raw)))
                If Not String.IsNullOrWhiteSpace(s) Then list.Add(s)
            End If
            Return list
        End Function

        Private Shared Function EnumeratePayloadItems(raw As Object) As List(Of Object)
            Dim items As New List(Of Object)()
            If raw Is Nothing Then Return items

            If IsJsonArrayValue(raw) Then
                Dim arr = TryCast(InvokeRuntimeMethod(raw, "EnumerateArray"), System.Collections.IEnumerable)
                If arr Is Nothing Then Return items
                For Each item In arr
                    items.Add(ConvertPayloadValue(item))
                Next
                Return items
            End If

            Dim enumerable = TryCast(raw, System.Collections.IEnumerable)
            If enumerable Is Nothing OrElse TypeOf raw Is String Then Return items

            For Each item In enumerable
                items.Add(ConvertPayloadValue(item))
            Next
            Return items
        End Function

        Private Shared Function ConvertPayloadValue(raw As Object) As Object
            If raw Is Nothing Then Return Nothing

            If IsJsonElementValue(raw) Then
                Dim kind = GetJsonValueKindName(raw)
                Select Case kind
                    Case "Object"
                        Return ToDict(raw)
                    Case "Array"
                        Return EnumeratePayloadItems(raw)
                    Case "String"
                        Return SafeStr(InvokeRuntimeMethod(raw, "GetString"))
                    Case "Null", "Undefined"
                        Return Nothing
                    Case Else
                        Return SafeStr(raw)
                End Select
            End If

            Return raw
        End Function

        Private Shared Function TryPopulateFromJsonObject(obj As Object, target As Dictionary(Of String, Object)) As Boolean
            If target Is Nothing OrElse Not IsJsonElementValue(obj) Then Return False
            If Not String.Equals(GetJsonValueKindName(obj), "Object", StringComparison.OrdinalIgnoreCase) Then Return False

            Dim props = TryCast(InvokeRuntimeMethod(obj, "EnumerateObject"), System.Collections.IEnumerable)
            If props Is Nothing Then Return False

            For Each item In props
                Dim name = SafeStr(GetRuntimePropertyValue(item, "Name"))
                If String.IsNullOrWhiteSpace(name) Then Continue For
                Dim value = GetRuntimePropertyValue(item, "Value")
                target(name) = ConvertPayloadValue(value)
            Next
            Return True
        End Function

        Private Shared Function IsJsonArrayValue(obj As Object) As Boolean
            Return IsJsonElementValue(obj) AndAlso String.Equals(GetJsonValueKindName(obj), "Array", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsJsonElementValue(obj As Object) As Boolean
            If obj Is Nothing Then Return False
            Dim t = obj.GetType()
            Return t IsNot Nothing AndAlso String.Equals(t.FullName, "System.Text.Json.JsonElement", StringComparison.Ordinal)
        End Function

        Private Shared Function GetJsonValueKindName(obj As Object) As String
            Return SafeStr(GetRuntimePropertyValue(obj, "ValueKind"))
        End Function

        Private Shared Function GetRuntimePropertyValue(obj As Object, propertyName As String) As Object
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return Nothing
            Try
                Dim prop = obj.GetType().GetProperty(propertyName)
                If prop Is Nothing Then Return Nothing
                Return prop.GetValue(obj, Nothing)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function InvokeRuntimeMethod(obj As Object, methodName As String) As Object
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(methodName) Then Return Nothing
            Try
                Dim methodInfo = obj.GetType().GetMethod(methodName, Type.EmptyTypes)
                If methodInfo Is Nothing Then Return Nothing
                Return methodInfo.Invoke(obj, Nothing)
            Catch
                Return Nothing
            End Try
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
