Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows
Imports System.Windows.Threading
Imports Autodesk.Revit.UI

' 네임스페이스는 프로젝트와 일치시켜 주세요.
Namespace UI.Hub

    ' Web(JS) ↔ Revit(VB) 브릿지의 중심. 외부이벤트 큐, 메시지 라우팅, 공통 로그/브로드캐스트 담당
    Partial Public Class UiBridgeExternalEvent

        ' -----------------------------
        ' 상태/싱글톤
        ' -----------------------------
        Friend Shared _host As HubHostWindow ' 다른 partial(Export 등)에서 사용
        Private Shared ReadOnly _self As New UiBridgeExternalEvent() ' 인스턴스 메서드 호출용
        Private Shared ReadOnly _gate As New Object()
        Private Shared ReadOnly _queue As New Queue(Of Action(Of UIApplication))()

        Private Shared _extEv As ExternalEvent
        Private Shared _handler As IExternalEventHandler

        ' -----------------------------
        ' 공개 진입점
        ' -----------------------------

        ''' <summary>
        ''' CmdOpenHub 에서 최초 1회 호출. Host 저장 + ExternalEvent 준비 + 초기 브로드캐스트.
        ''' </summary>
        Public Shared Sub Initialize(host As HubHostWindow)
            _host = host

            If _extEv Is Nothing Then
                _handler = New BridgeHandler(AddressOf ProcessQueue)
                _extEv = ExternalEvent.Create(_handler)
            End If

            ' 초기 상태 브로드캐스트(항상 위, 연결)
            BroadcastTopmost()
            SendToWeb("host:connected", New With {.ok = True})
        End Sub

        ''' <summary>
        ''' Host(또는 Web 메시지 핸들러)에서 호출: 이름/페이로드를 외부이벤트 큐에 넣고 Revit UI 스레드에서 처리.
        ''' </summary>
        Public Shared Sub Raise(name As String, payload As Object)
            Enqueue(Sub(app) Dispatch(app, name, payload))
        End Sub

        ' -----------------------------
        ' 큐/외부이벤트
        ' -----------------------------
        Private Shared Sub Enqueue(work As Action(Of UIApplication))
            SyncLock _gate
                _queue.Enqueue(work)
            End SyncLock
            If _extEv IsNot Nothing Then
                _extEv.Raise()
            End If
        End Sub

        Private Shared Sub ProcessQueue(app As UIApplication)
            Dim todo As Action(Of UIApplication) = Nothing
            Do
                SyncLock _gate
                    If _queue.Count > 0 Then
                        todo = _queue.Dequeue()
                    Else
                        todo = Nothing
                    End If
                End SyncLock
                If todo Is Nothing Then Exit Do
                Try
                    todo(app)
                Catch ex As Exception
                    SendToWeb("host:error", New With {.message = ex.Message})
                End Try
            Loop
        End Sub

        ' -----------------------------
        ' 라우팅
        ' -----------------------------
        Private Shared Sub Dispatch(app As UIApplication, name As String, payload As Object)
            name = NormalizeEventName(name)
            ' 공통(웹 UI 쪽 요청)
            Select Case name
                Case "ui:query-topmost"
                    BroadcastTopmost()
                    Return

                Case "ui:set-topmost"
                    Dim turnOn As Boolean = False
                    Try
                        Dim raw = GetProp(payload, "on")
                        If raw IsNot Nothing Then turnOn = Convert.ToBoolean(raw)
                    Catch
                    End Try
                    Try
                        If _host IsNot Nothing Then _host.Topmost = turnOn
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return

                Case "ui:toggle-topmost"
                    Try
                        If _host IsNot Nothing Then _host.Topmost = Not _host.Topmost
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return
            End Select

            ' 기능 디스패치: 이벤트명 → 내부 핸들러명
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' Duplicate Inspector
            map.Add("dup:run", "HandleDupRun")
            map.Add("duplicate:export", "HandleDuplicateExport")
            map.Add("duplicate:delete", "HandleDuplicateDelete")
            map.Add("duplicate:restore", "HandleDuplicateRestore")
            map.Add("duplicate:select", "HandleDuplicateSelect")
            ' Connector Diagnostics
            map.Add("connector:run", "HandleConnectorRun")
            map.Add("connector:param-list", "HandleConnectorParamList")
            map.Add("connector:save-excel", "HandleConnectorSaveExcel")
            map.Add("floorinfo:config-load", "HandleFloorInfoConfigLoad")
            ' Export Points with Angle
            map.Add("export:browse-folder", "HandleExportBrowse")
            map.Add("export:add-rvt-files", "HandleExportAddRvtFiles")
            map.Add("export:preview", "HandleExportPreview")
            map.Add("export:save-excel", "HandleExportSaveExcel")
            ' Shared Parameter Propagator
            map.Add("paramprop:run", "HandleSharedParamRun")
            map.Add("sharedparam:run", "HandleSharedParamRun")
            map.Add("sharedparam:list", "HandleSharedParamList")
            map.Add("sharedparam:status", "HandleSharedParamStatus")
            map.Add("sharedparam:export-excel", "HandleSharedParamExport")
            ' (호환) 일부 번들/구버전 UI에서 paramprop:* 로 호출하는 경우를 허용
            map.Add("paramprop:list", "HandleSharedParamList")
            map.Add("paramprop:status", "HandleSharedParamStatus")
            map.Add("paramprop:export-excel", "HandleSharedParamExport")
            map.Add("sharedparam:export", "HandleSharedParamExport")
            ' Shared Parameter Batch
            map.Add("sharedparambatch:init", "HandleSharedParamBatchInit")
            map.Add("sharedparambatch:browse-rvts", "HandleSharedParamBatchBrowseRvts")
            map.Add("sharedparambatch:browse-folder", "HandleSharedParamBatchBrowseFolder")
            map.Add("sharedparambatch:run", "HandleSharedParamBatchRun")
            map.Add("sharedparambatch:export-excel", "HandleSharedParamBatchExportExcel")
            ' 공통 Excel 동작
            map.Add("excel:open", "HandleExcelOpen")
            ' Segment ↔ PMS Check
            map.Add("segmentpms:rvt-pick-files", "HandleSegmentPmsRvtPickFiles")
            map.Add("segmentpms:rvt-pick-folder", "HandleSegmentPmsRvtPickFolder")
            map.Add("segmentpms:extract", "HandleSegmentPmsExtractStart")
            map.Add("segmentpms:load-extract", "HandleSegmentPmsLoadExtract")
            map.Add("segmentpms:save-extract", "HandleSegmentPmsSaveExtract")
            map.Add("segmentpms:register-pms", "HandleSegmentPmsRegisterPms")
            map.Add("segmentpms:pms-template", "HandleSegmentPmsExportTemplate")
            map.Add("segmentpms:prepare-mapping", "HandleSegmentPmsPrepareMapping")
            map.Add("segmentpms:run", "HandleSegmentPmsRun")
            map.Add("segmentpms:save-result", "HandleSegmentPmsSaveResult")
            ' GUID Audit
            map.Add("guid:add-files", "HandleGuidAddFiles")
            map.Add("guid:run", "HandleGuidRun")
            map.Add("guid:export", "HandleGuidExport")
            map.Add("guid:request-family-detail", "HandleGuidRequestFamilyDetail")
            ' Family Link Audit
            map.Add("familylink:init", "HandleFamilyLinkInit")
            map.Add("familylink:pick-rvts", "HandleFamilyLinkPickRvts")
            map.Add("familylink:run", "HandleFamilyLinkRun")
            map.Add("familylink:export", "HandleFamilyLinkExport")
            ' Multi RVT Hub
            map.Add("hub:pick-rvt", "HandleMultiPickRvt")
            map.Add("hub:multi-run", "HandleMultiRun")
            map.Add("hub:multi-export", "HandleMultiExport")
            map.Add("hub:multi-clear", "HandleMultiClear")
            map.Add("commonoptions:get", "HandleCommonOptionsGet")
            map.Add("commonoptions:save", "HandleCommonOptionsSave")
            map.Add("deliverycleaner:init", "HandleDeliveryCleanerInit")
            map.Add("deliverycleaner:pick-rvts", "HandleDeliveryCleanerPickRvts")
            map.Add("deliverycleaner:browse-output-folder", "HandleDeliveryCleanerBrowseOutputFolder")
            map.Add("deliverycleaner:filter-import", "HandleDeliveryCleanerFilterImport")
            map.Add("deliverycleaner:filter-save", "HandleDeliveryCleanerFilterSave")
            map.Add("deliverycleaner:filter-doc-list", "HandleDeliveryCleanerFilterDocList")
            map.Add("deliverycleaner:filter-doc-extract", "HandleDeliveryCleanerFilterDocExtract")
            map.Add("deliverycleaner:run", "HandleDeliveryCleanerRun")
            map.Add("deliverycleaner:verify", "HandleDeliveryCleanerVerify")
            map.Add("deliverycleaner:extract", "HandleDeliveryCleanerExtract")
            map.Add("deliverycleaner:purge", "HandleDeliveryCleanerPurge")
            map.Add("deliverycleaner:purge-status", "HandleDeliveryCleanerPurgeStatus")
            map.Add("deliverycleaner:export-verify", "HandleDeliveryCleanerExportVerify")
            map.Add("deliverycleaner:export-extract", "HandleDeliveryCleanerExportExtract")
            map.Add("deliverycleaner:export-designoption", "HandleDeliveryCleanerExportDesignOption")
            map.Add("deliverycleaner:export-purge", "HandleDeliveryCleanerExportPurge")
            map.Add("deliverycleaner:export-log", "HandleDeliveryCleanerExportLog")
            map.Add("deliverycleaner:open-folder", "HandleDeliveryCleanerOpenFolder")
            ' App Update
            map.Add("update:query", "HandleUpdateQuery")
            map.Add("update:check", "HandleUpdateCheck")
            map.Add("update:install", "HandleUpdateInstall")

            Dim methodName As String = Nothing
            If Not map.TryGetValue(name, methodName) Then
                ' warn만 남기면(특히 Busy 상태) UI가 복귀하지 못하는 경우가 있어 error도 함께 전송
                SendToWeb("host:warn", New With {.message = String.Format("알 수 없는 이벤트 '{0}'", name)})
                SendToWeb("host:error", New With {.message = String.Format("알 수 없는 이벤트 '{0}'", name)})
                Return
            End If

            ' 동일 Partial 클래스 안의 Private 메서드를 리플렉션으로 찾아 호출
            Dim t As Type = GetType(UiBridgeExternalEvent)
            Dim flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public
            Dim m As MethodInfo = t.GetMethod(methodName, flags)

            If m Is Nothing Then
                ' 핸들러 누락은 실제로 기능이 동작하지 않는 치명 오류이므로, warn + error 둘 다 전송
                SendToWeb("host:warn", New With {.message = String.Format("핸들러 '{0}' 가 구현되어 있지 않습니다.", methodName)})
                SendToWeb("host:error", New With {.message = String.Format("핸들러 '{0}' 가 구현되어 있지 않습니다.", methodName)})
                Return
            End If

            ' 시그니처: (UIApplication, payload) or (UIApplication) or (payload) or ()
            Dim ps() As ParameterInfo = m.GetParameters()
            Dim args() As Object
            Select Case ps.Length
                Case 2
                    args = New Object() {app, payload}
                Case 1
                    If ps(0).ParameterType Is GetType(UIApplication) Then
                        args = New Object() {app}
                    Else
                        args = New Object() {payload}
                    End If
                Case Else
                    args = New Object() {}
            End Select

            Try
                m.Invoke(_self, args)
            Catch ex As TargetInvocationException
                Dim msg As String
                If ex.InnerException IsNot Nothing Then
                    msg = ex.InnerException.Message
                Else
                    msg = ex.Message
                End If
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, msg)})
            Catch ex As Exception
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, ex.Message)})
            End Try
        End Sub

        ' -----------------------------
        ' 공통 유틸/브로드캐스트
        ' -----------------------------

        Friend Shared Sub SendToWeb(channel As String, payload As Object)
            Dim host = _host
            If host Is Nothing Then Return

            Dim dispatch As Action =
                Sub()
                    Try
                        host.SendToWeb(channel, payload)
                    Catch
                    End Try
                End Sub

            Try
                Dim dispatcher = host.Dispatcher
                If dispatcher Is Nothing Then
                    dispatch()
                    Return
                End If

                If dispatcher.CheckAccess() Then
                    dispatch()
                Else
                    dispatcher.BeginInvoke(dispatch, DispatcherPriority.Background)
                End If
            Catch
                dispatch()
            End Try
        End Sub

        Friend Shared Sub SendToWebAfterDialog(channel As String, payload As Object)
            Dim host = _host
            If host Is Nothing Then
                SendToWeb(channel, payload)
                Return
            End If

            Dim dispatch As Action =
                Sub()
                    Try
                        If host.WindowState = WindowState.Minimized Then
                            host.WindowState = WindowState.Normal
                        End If
                    Catch
                    End Try

                    Try
                        host.Activate()
                        host.Focus()
                    Catch
                    End Try

                    Try
                        host.SendToWeb(channel, payload)
                    Catch
                    End Try
                End Sub

            Try
                Dim dispatcher = host.Dispatcher
                If dispatcher Is Nothing Then
                    dispatch()
                    Return
                End If

                If dispatcher.CheckAccess() Then
                    dispatcher.BeginInvoke(dispatch, DispatcherPriority.ApplicationIdle)
                Else
                    dispatcher.BeginInvoke(dispatch, DispatcherPriority.ApplicationIdle)
                End If
            Catch
                dispatch()
            End Try
        End Sub

        Private Shared Sub BroadcastTopmost()
            Try
                Dim onTop As Boolean = False
                If _host IsNot Nothing Then onTop = _host.Topmost
                SendToWeb("host:topmost", New With {.on = onTop})
            Catch
            End Try
        End Sub

        Friend Shared Function ParseExcelMode(payload As Object) As Boolean
            Dim mode As String = "normal"

            Try
                If payload IsNot Nothing Then
                    Dim modeProp = GetProp(payload, "excelMode")
                    If modeProp IsNot Nothing Then
                        Dim raw = Convert.ToString(modeProp)
                        If Not String.IsNullOrWhiteSpace(raw) Then mode = raw
                    End If
                End If
            Catch
            End Try

            Return Not String.Equals(mode, "fast", StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Shared Function FilterIssueRowsCopy(styleKey As String, table As DataTable) As DataTable
            If table Is Nothing Then Return Nothing
            Dim copy As DataTable = table.Copy()
            Global.KKY_Tool_Revit.Infrastructure.ResultTableFilter.KeepOnlyIssues(styleKey, copy)
            Return copy
        End Function

        ' payload 속성 안전 추출(익명/Dictionary 수용)

        
        ' 문자열이 ""..."" 또는 '...'(따옴표)로 감싸져 들어오는 경우(특히 JSON/인자 구성 과정) 제거
        Private Shared Function NormalizeWrappedQuotesText(value As String) As String
            Dim s As String = If(value, "")
            s = s.Trim()

            For i As Integer = 0 To 1
                If s.Length >= 2 AndAlso s(0) = """"c AndAlso s(s.Length - 1) = """"c Then
                    s = s.Substring(1, s.Length - 2).Trim()
                    Continue For
                End If
                If s.Length >= 2 AndAlso s(0) = "'"c AndAlso s(s.Length - 1) = "'"c Then
                    s = s.Substring(1, s.Length - 2).Trim()
                    Continue For
                End If
                Exit For
            Next

            Return s
        End Function

Private Shared Function NormalizeEventName(name As String) As String
            Dim s As String = If(name, "")
            s = s.Trim()

            ' Some host parsers pass JSON raw text for strings: ""ui:xxx"" (including the quotes).
            ' Normalize by stripping wrapping quotes (single or double).
            For i As Integer = 0 To 1
                If s.Length >= 2 AndAlso s(0) = """"c AndAlso s(s.Length - 1) = """"c Then
                    s = s.Substring(1, s.Length - 2).Trim()
                    Continue For
                End If
                If s.Length >= 2 AndAlso s(0) = "'"c AndAlso s(s.Length - 1) = "'"c Then
                    s = s.Substring(1, s.Length - 2).Trim()
                    Continue For
                End If
                Exit For
            Next

            Return s
        End Function

        Private Shared Function GetProp(obj As Object, prop As String) As Object
            If obj Is Nothing Then Return Nothing

            Dim d = TryCast(obj, IDictionary(Of String, Object))
            If d IsNot Nothing Then
                Dim v As Object = Nothing
                If d.TryGetValue(prop, v) Then Return v
                Return Nothing
            End If

            Dim t = obj.GetType()
            Dim p = t.GetProperty(prop, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)
            If p Is Nothing Then Return Nothing
            Return p.GetValue(obj, Nothing)
        End Function

        ' (필요시) 외부에서 직접 로그를 남길 때 사용
        Public Shared Sub HostLog(kind As String, text As String)
            SendToWeb("host:log", New With {.kind = kind, .text = text})
        End Sub

        Friend Shared Sub LogAutoFitDecision(doAutoFit As Boolean, context As String)
            Dim tag As String = If(doAutoFit, "NORMAL_EXPORT: AutoFit applied", "FAST_EXPORT: AutoFit skipped")
            HostLog("debug", $"{tag} [{context}]")
        End Sub

        Private Sub HandleExcelOpen(payload As Object)
            Dim inputPath As String = TryCast(GetProp(payload, "path"), String)
            If String.IsNullOrWhiteSpace(inputPath) Then
                SendToWeb("host:warn", New With {.message = "엑셀 경로가 비어 있습니다."})
                Return
            End If

            ' 일부 호출 경로에서 "C:\...ile.xlsx" 처럼 따옴표가 포함되어 전달되는 경우가 있어 정규화
            inputPath = NormalizeWrappedQuotesText(inputPath)

            ' file:///C:/... 형태 대응
            Try
                If inputPath.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
                    Dim u As New Uri(inputPath)
                    If u IsNot Nothing AndAlso u.IsFile Then
                        inputPath = u.LocalPath
                    End If
                End If
            Catch
                ' ignore
            End Try

            Dim fullPath As String = inputPath
            Try
                If System.IO.Path.IsPathRooted(inputPath) Then
                    fullPath = inputPath
                Else
                    fullPath = System.IO.Path.GetFullPath(inputPath)
                End If
            Catch
                ' ignore
            End Try

            Dim fileExists As Boolean = False
            Try
                fileExists = System.IO.File.Exists(fullPath)
            Catch
                ' ignore
            End Try

            If Not fileExists Then
                SendToWeb("host:warn", New With {.message = "엑셀 파일을 찾을 수 없습니다: " & fullPath, .path = fullPath})
                Return
            End If

            Try
                Dim info As New System.IO.FileInfo(fullPath)
                Dim lengthOk As Boolean = False
                For i As Integer = 0 To 9
                    info.Refresh()
                    If info.Exists AndAlso info.Length > 0 Then
                        lengthOk = True
                        Exit For
                    End If
                    System.Threading.Thread.Sleep(200)
                Next
                If Not lengthOk Then
                    SendToWeb("host:warn", New With {.message = "파일 크기가 0입니다. 열기를 시도합니다: " & fullPath, .path = fullPath})
                End If
            Catch ex As Exception
                SendToWeb("host:warn", New With {.message = "파일 상태 확인에 실패했습니다: " & ex.Message, .path = fullPath})
            End Try

            Dim opened As Boolean = False
            Dim firstError As Exception = Nothing

            Try
                Dim psi As New System.Diagnostics.ProcessStartInfo(fullPath)
                psi.UseShellExecute = True

                Dim directoryPath As String = String.Empty
                Try
                    directoryPath = System.IO.Path.GetDirectoryName(fullPath)
                Catch
                    directoryPath = String.Empty
                End Try
                If Not String.IsNullOrWhiteSpace(directoryPath) Then
                    psi.WorkingDirectory = directoryPath
                End If

                System.Diagnostics.Process.Start(psi)
                opened = True
                SendToWeb("host:info", New With {.message = "엑셀 열기를 시도했습니다: " & fullPath, .path = fullPath})
            Catch ex As Exception
                firstError = ex
            End Try

            If Not opened Then
                Try
                    Dim psi As New System.Diagnostics.ProcessStartInfo("explorer.exe", "/select,""" & fullPath & """")
                    psi.UseShellExecute = True
                    System.Diagnostics.Process.Start(psi)
                    opened = True

                    Dim warnMsg As String = "엑셀 열기에 실패하여 탐색기로 열었습니다: " & fullPath
                    If firstError IsNot Nothing Then
                        warnMsg &= " (" & firstError.Message & ")"
                    End If
                    SendToWeb("host:warn", New With {.message = warnMsg, .path = fullPath})
                Catch ex As Exception
                    Dim msg As String = "엑셀 열기 실패: " & ex.Message
                    If firstError IsNot Nothing Then
                        msg &= " / 최초 오류: " & firstError.Message
                    End If
                    SendToWeb("host:warn", New With {.message = msg, .path = fullPath})
                End Try
            End If
        End Sub

        Private Sub HandleSwitchDocument(app As UIApplication, payload As Object)
            ' 더 이상 문서를 OpenAndActivateDocument로 다시 열지 않는다.
            ' 단순 안내 로그만 남긴다.
            Dim name As String = TryCast(GetProp(payload, "name"), String)
            If String.IsNullOrWhiteSpace(name) Then
                name = TryCast(GetProp(payload, "path"), String)
            End If

            SendToWeb("host:info", New With {
                .message = "문서 전환은 Revit 창에서 직접 선택해 주세요.",
                .target = name
            })
        End Sub

    End Class

    ' 외부이벤트 핸들러(큐를 비우는 역할만 수행)
    Friend Class BridgeHandler
        Implements IExternalEventHandler

        Private ReadOnly _run As Action(Of UIApplication)
        Public Sub New(run As Action(Of UIApplication))
            _run = run
        End Sub

        Public Sub Execute(uiApp As UIApplication) Implements IExternalEventHandler.Execute
            If _run IsNot Nothing Then
                _run.Invoke(uiApp)
            End If
        End Sub

        ' Revit API는 Function GetName() As String 을 요구합니다.

        ' -----------------------------
        ' [공통] 공종검토 엑셀 기본 파일명 생성
        ' -----------------------------
        Friend Shared Function BuildTradeReviewDefaultExcelName(rvtBaseName As String, issueCount As Integer) As String
            Try
                Dim baseName As String = If(rvtBaseName, "").Trim()
                If String.IsNullOrWhiteSpace(baseName) Then Return ""

                ' 확장자 제거(들어온 값이 Path 일 수도 있음)
                baseName = System.IO.Path.GetFileNameWithoutExtension(baseName)

                Dim prefix As String = ExtractTradePrefix(baseName)
                Dim nameCore As String

                If String.IsNullOrWhiteSpace(prefix) Then
                    ' 규칙 불일치: 파일명 전체_공종검토
                    nameCore = $"{baseName}_공종검토_({issueCount}건).xlsx"
                Else
                    nameCore = $"{prefix}-06_공종검토_0차_({issueCount}건).xlsx"
                End If

                Return SanitizeFileName(nameCore)
            Catch
                Return ""
            End Try
        End Function

        Friend Shared Function ExtractTradePrefix(rvtBaseName As String) As String
            Dim s As String = If(rvtBaseName, "").Trim()
            If String.IsNullOrWhiteSpace(s) Then Return ""

            Dim parts = s.Split("_"c)
            If parts Is Nothing OrElse parts.Length < 4 Then Return ""

            Dim token As String = parts(3)
            If String.IsNullOrWhiteSpace(token) Then Return ""

            Dim m = System.Text.RegularExpressions.Regex.Match(token, "^(.*?-\d+)")
            Dim cut As String = If(m.Success, m.Groups(1).Value, token)

            Return $"{parts(0)}_{parts(1)}_{parts(2)}_{cut}"
        End Function

        Friend Shared Function SanitizeFileName(fileName As String) As String
            Dim s As String = If(fileName, "")
            If String.IsNullOrWhiteSpace(s) Then Return ""
            For Each ch In System.IO.Path.GetInvalidFileNameChars()
                s = s.Replace(ch, "_"c)
            Next
            Return s
        End Function

        Public Function GetName() As String Implements IExternalEventHandler.GetName
            Return "KKY Hub Bridge"
        End Function
    End Class

End Namespace
