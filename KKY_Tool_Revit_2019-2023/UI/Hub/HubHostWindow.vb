Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Web.Script.Serialization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading

Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf

Namespace UI.Hub

    Public Class HubHostWindow
        Inherits Window

        Private Const BaseTitle As String = "KKY Tool Hub"

        Private Shared _instance As HubHostWindow
        Private Shared ReadOnly _gate As New Object()

        Private ReadOnly _web As New WebView2()
        Private ReadOnly _dropOverlayHideTimer As New DispatcherTimer()
        Private ReadOnly _serializer As New JavaScriptSerializer()

        Private _uiApp As UIApplication
        Private _currentDocName As String = String.Empty
        Private _currentDocPath As String = String.Empty
        Private _currentRouteKey As String = String.Empty
        Private _isNativeDropOverlayVisible As Boolean = False
        Private _dropOverlayWindow As NativeDropOverlayWindow = Nothing
        Private _pendingNativeDropPaths As New List(Of String)()
        Private _pendingNativeDropRouteKey As String = String.Empty
        Private _pendingNativeDropStampUtc As DateTime = DateTime.MinValue
        Private _initStarted As Boolean = False
        Private _isClosing As Boolean = False

        Public ReadOnly Property Web As WebView2
            Get
                Return _web
            End Get
        End Property

        Public ReadOnly Property IsClosing As Boolean
            Get
                Return _isClosing
            End Get
        End Property

        Public Shared Sub ShowSingleton(uiApp As UIApplication)
            If uiApp Is Nothing Then Return

            SyncLock _gate
                If _instance IsNot Nothing AndAlso Not _instance.IsClosing Then
                    _instance.AttachTo(uiApp)
                    UiBridgeExternalEvent.Initialize(_instance)

                    If _instance.WindowState = WindowState.Minimized Then
                        _instance.WindowState = WindowState.Normal
                    End If

                    _instance.Activate()
                    _instance.Focus()
                    Return
                End If

                Dim wnd As New HubHostWindow(uiApp)
                UiBridgeExternalEvent.Initialize(wnd)
                _instance = wnd
                wnd.Show()
            End SyncLock
        End Sub

        Public Shared Sub NotifyActiveDocumentChanged(doc As Document)
            Dim inst = _instance
            If inst Is Nothing OrElse inst.IsClosing Then Return
            inst.UpdateActiveDocument(doc)
        End Sub

        Public Shared Sub NotifyDocumentListChanged()
            Dim inst = _instance
            If inst Is Nothing OrElse inst.IsClosing Then Return
            inst.BroadcastDocumentList()
        End Sub

        Public Sub New(uiApp As UIApplication)
            _uiApp = uiApp

            Title = BaseTitle

            Dim workArea = SystemParameters.WorkArea
            Dim desiredWidth As Double = 1400
            Dim desiredHeight As Double = 900

            Width = Math.Min(desiredWidth, workArea.Width * 0.93)
            Height = Math.Min(desiredHeight, workArea.Height * 0.93)
            MinWidth = Math.Min(1100, Width)
            MinHeight = Math.Min(720, Height)

            WindowStartupLocation = WindowStartupLocation.CenterScreen
            AllowDrop = True
            _web.AllowDrop = True
            _web.AllowExternalDrop = True
            _dropOverlayHideTimer.Interval = TimeSpan.FromMilliseconds(180)
            AddHandler _dropOverlayHideTimer.Tick, AddressOf OnDropOverlayHideTimerTick

            Content = _web
            AddHandler Loaded, AddressOf OnLoaded
            AddHandler Closing, AddressOf OnWindowClosing
            AddHandler Closed, AddressOf OnWindowClosed
            AddHandler LocationChanged, AddressOf OnHostWindowBoundsChanged
            AddHandler SizeChanged, AddressOf OnHostWindowBoundsChanged
            AddHandler StateChanged, AddressOf OnHostWindowStateChanged
            AddHandler PreviewDragOver, AddressOf HandlePreviewDragOver
            AddHandler PreviewDrop, AddressOf HandlePreviewDrop
            AddHandler PreviewDragEnter, AddressOf HandlePreviewDragEnter
            AddHandler _web.PreviewDragOver, AddressOf HandlePreviewDragOver
            AddHandler _web.PreviewDrop, AddressOf HandlePreviewDrop
            AddHandler _web.PreviewDragEnter, AddressOf HandlePreviewDragEnter
            AddHandler _web.DragEnter, AddressOf HandlePreviewDragEnter
            AddHandler _web.DragOver, AddressOf HandlePreviewDragOver
            AddHandler _web.Drop, AddressOf HandlePreviewDrop

            UpdateActiveDocument(GetActiveDocument())
        End Sub

        Public Sub AttachTo(uiApp As UIApplication)
            _uiApp = uiApp
            UpdateActiveDocument(GetActiveDocument())
            BroadcastDocumentList()
        End Sub

        Private Function GetActiveDocument() As Document
            Try
                If _uiApp Is Nothing Then Return Nothing
                Dim uidoc = _uiApp.ActiveUIDocument
                If uidoc Is Nothing Then Return Nothing
                Return uidoc.Document
            Catch
                Return Nothing
            End Try
        End Function

        Private Sub UpdateActiveDocument(doc As Document)
            Dim name As String = String.Empty
            Dim path As String = String.Empty

            If doc IsNot Nothing Then
                Try : name = doc.Title : Catch : End Try
                Try : path = doc.PathName : Catch : End Try
            End If

            If String.IsNullOrWhiteSpace(path) Then path = name

            _currentDocName = name
            _currentDocPath = path

            UpdateWindowTitle()
            SendActiveDocument()
        End Sub

        Private Sub UpdateWindowTitle()
            If String.IsNullOrWhiteSpace(_currentDocName) Then
                Title = BaseTitle
            Else
                Title = $"{BaseTitle} - {_currentDocName}"
            End If
        End Sub

        Private Function ResolveUiFolder() As String
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim baseDir = Path.GetDirectoryName(asm.Location)
                Dim ui = Path.Combine(baseDir, "Resources", "HubUI")
                If Directory.Exists(ui) Then Return Path.GetFullPath(ui)
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function BuildWebView2UserDataFolder() As String
            ' ✅ 여러 Revit 프로세스(두 개의 Revit 실행) 동시 사용 시,
            '    WebView2 UserDataFolder를 공유하면 잠금 때문에 다른 Revit가 로딩에서 멈출 수 있음.
            '    프로세스별(UserDataFolder/pid_xxxx)로 분리해서 서로 독립적으로 동작하게 한다.
            Dim baseFolder As String =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                             "KKY_Tool_Revit",
                             "WebView2UserData")

            Dim pid As Integer = 0
            Try : pid = Process.GetCurrentProcess().Id : Catch : pid = 0 : End Try

            Dim folder As String = If(pid > 0, Path.Combine(baseFolder, "pid_" & pid.ToString()), baseFolder)
            Try : Directory.CreateDirectory(folder) : Catch : End Try
            Return folder
        End Function

        Private Async Sub OnLoaded(sender As Object, e As RoutedEventArgs)
            If _initStarted Then Return
            _initStarted = True

            Try
                Dim userData = BuildWebView2UserDataFolder()
                Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userData, Nothing)
                Await _web.EnsureCoreWebView2Async(env)

                Dim core = _web.CoreWebView2
                core.Settings.AreDefaultContextMenusEnabled = False
                core.Settings.IsStatusBarEnabled = False

#If DEBUG Then
                core.Settings.AreDevToolsEnabled = True
#Else
                core.Settings.AreDevToolsEnabled = False
#End If

                ' 가상 호스트 매핑
                Dim uiFolder = ResolveUiFolder()
                If String.IsNullOrEmpty(uiFolder) Then
                    Throw New DirectoryNotFoundException("Resources\HubUI 폴더를 찾을 수 없습니다.")
                End If

                core.SetVirtualHostNameToFolderMapping("hub.local", uiFolder, CoreWebView2HostResourceAccessKind.Allow)

                ' 메시지 브리지
                AddHandler core.WebMessageReceived, AddressOf OnWebMessage
                AddHandler core.NavigationStarting, AddressOf OnNavigationStarting
                AddHandler core.NewWindowRequested, AddressOf OnNewWindowRequested

                ' 허브 진입
                _web.Source = New Uri("https://hub.local/index.html")

                ' 초기 상태 알림
                SendToWeb("host:topmost", New With {.on = Me.Topmost})
                SendActiveDocument()
                BroadcastDocumentList()

            Catch ex As Exception
                Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
                MessageBox.Show($"WebView 초기화 실패 (0x{hr:X8}) : {ex.Message}",
                                "KKY Tool",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub OnWindowClosing(sender As Object, e As System.ComponentModel.CancelEventArgs)
            _isClosing = True
        End Sub

        Private Sub OnWindowClosed(sender As Object, e As EventArgs)
            Try
                If _dropOverlayWindow IsNot Nothing Then
                    _dropOverlayWindow.Close()
                    _dropOverlayWindow = Nothing
                End If
            Catch
            End Try
            SyncLock _gate
                If _instance Is Me Then _instance = Nothing
            End SyncLock
        End Sub

        Private Sub OnHostWindowBoundsChanged(sender As Object, e As EventArgs)
            SyncNativeDropOverlayBounds()
        End Sub

        Private Sub OnHostWindowStateChanged(sender As Object, e As EventArgs)
            If WindowState = WindowState.Minimized Then
                ApplyNativeDropOverlayVisibility(False)
                Return
            End If
            SyncNativeDropOverlayBounds()
        End Sub

        Private Sub OnWebMessage(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
            Try
                Dim root As Dictionary(Of String, Object) =
                    _serializer.Deserialize(Of Dictionary(Of String, Object))(e.WebMessageAsJson)

                Dim name As String = Nothing

                If root IsNot Nothing Then
                    If root.ContainsKey("ev") AndAlso root("ev") IsNot Nothing Then
                        name = Convert.ToString(root("ev"))
                    ElseIf root.ContainsKey("name") AndAlso root("name") IsNot Nothing Then
                        name = Convert.ToString(root("name"))
                    End If
                End If

                If String.IsNullOrEmpty(name) Then Return

                Dim payload As Object = Nothing
                If root IsNot Nothing AndAlso root.ContainsKey("payload") Then payload = root("payload")

                Select Case name
                    Case "ui:ping"
                        SendToWeb("host:pong", New With {.t = Date.Now.Ticks})

                    Case "ui:toggle-topmost"
                        Me.Topmost = Not Me.Topmost
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case "ui:query-topmost"
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case "ui:route-changed"
                        Dim routePayload = TryCast(payload, Dictionary(Of String, Object))
                        Dim route As String = String.Empty
                        If routePayload IsNot Nothing AndAlso routePayload.ContainsKey("route") AndAlso routePayload("route") IsNot Nothing Then
                            route = Convert.ToString(routePayload("route"))
                        End If
                        _currentRouteKey = NormalizeRouteKey(route)
                        ClearPendingNativeDropPaths()
                        SendHostDebug("route-changed", New With {
                            .route = _currentRouteKey
                        })

                    Case "ui:rvt-drop-overlay"
                        Dim overlayPayload = TryCast(payload, Dictionary(Of String, Object))
                        Dim active As Boolean = False
                        If overlayPayload IsNot Nothing AndAlso overlayPayload.ContainsKey("active") Then
                            active = SafeBoolObj(overlayPayload("active"), False)
                        End If
                        Dim routeKey = GetCurrentRouteKey()
                        SendHostDebug("overlay-request", New With {
                            .active = active,
                            .route = routeKey,
                            .supported = SupportsDroppedRvtRoute(routeKey),
                            .overlayVisible = _isNativeDropOverlayVisible
                        })
                        SetNativeDropOverlay(active)

                    Case "ui:rvt-drop-commit"
                        Dim routeKey = GetCurrentRouteKey()
                        Dim pending = ConsumePendingNativeDropPaths(routeKey)
                        Dim commitPayload = TryCast(payload, Dictionary(Of String, Object))
                        Dim fileNames As New List(Of String)()
                        Dim filesLength As Integer = 0
                        If commitPayload IsNot Nothing AndAlso commitPayload.ContainsKey("fileNames") AndAlso commitPayload("fileNames") IsNot Nothing Then
                            Dim rawNames = TryCast(commitPayload("fileNames"), System.Collections.IEnumerable)
                            If rawNames IsNot Nothing AndAlso Not TypeOf commitPayload("fileNames") Is String Then
                                fileNames = rawNames _
                                    .Cast(Of Object)() _
                                    .Select(Function(item) Convert.ToString(item)) _
                                    .Where(Function(nameText) Not String.IsNullOrWhiteSpace(nameText)) _
                                    .ToList()
                            End If
                        End If
                        If commitPayload IsNot Nothing AndAlso commitPayload.ContainsKey("filesLength") AndAlso commitPayload("filesLength") IsNot Nothing Then
                            Integer.TryParse(Convert.ToString(commitPayload("filesLength")), filesLength)
                        End If
                        SendHostDebug("overlay-commit", New With {
                            .route = routeKey,
                            .cachedPathCount = pending.Count,
                            .fileNames = fileNames,
                            .filesLength = filesLength
                        })
                        If pending.Count > 0 Then
                            DispatchDroppedRvts(routeKey, pending)
                        ElseIf fileNames.Count > 0 OrElse filesLength > 0 Then
                            SendToWeb("host:rvt-drop-invalid", New With {
                                .route = routeKey,
                                .fileNames = fileNames,
                                .message = "RVT 파일만 추가할 수 있습니다."
                            })
                        End If

                    Case Else
                        UiBridgeExternalEvent.Raise(name, payload)
                End Select

            Catch ex As Exception
                SendToWeb("host:error", New With {ex.Message})
            End Try
        End Sub

        Private Sub HandlePreviewDragEnter(sender As Object, e As DragEventArgs)
            If Not HasFileDrop(e.Data) Then Return
            Dim paths = ExtractDroppedRvtPaths(e.Data)
            CachePendingNativeDropPaths(GetCurrentRouteKey(), paths)
            SendHostDebug("drag-enter", New With {
                .sender = DescribeSender(sender),
                .route = GetCurrentRouteKey(),
                .formats = GetDataFormats(e.Data),
                .pathCount = paths.Count
            })
            e.Effects = DragDropEffects.Copy
            e.Handled = True
        End Sub

        Private Sub HandlePreviewDragOver(sender As Object, e As DragEventArgs)
            If Not HasFileDrop(e.Data) Then Return
            Dim paths = ExtractDroppedRvtPaths(e.Data)
            CachePendingNativeDropPaths(GetCurrentRouteKey(), paths)
            SendHostDebug("drag-over", New With {
                .sender = DescribeSender(sender),
                .route = GetCurrentRouteKey(),
                .overlayVisible = _isNativeDropOverlayVisible,
                .pathCount = paths.Count
            })
            e.Effects = DragDropEffects.Copy
            e.Handled = True
        End Sub

        Private Sub HandlePreviewDrop(sender As Object, e As DragEventArgs)
            Try
                Dim paths = ExtractDroppedRvtPaths(e.Data)
                If paths.Count = 0 Then
                    paths = ConsumePendingNativeDropPaths(GetCurrentRouteKey())
                End If
                SendHostDebug("drop-received", New With {
                    .sender = DescribeSender(sender),
                    .route = GetCurrentRouteKey(),
                    .formats = GetDataFormats(e.Data),
                    .pathCount = paths.Count,
                    .paths = paths
                })
                If paths.Count = 0 Then Return

                DispatchDroppedRvts(GetCurrentRouteKey(), paths)
                SetNativeDropOverlay(False, True)
                e.Handled = True
            Catch ex As Exception
                SetNativeDropOverlay(False, True)
                SendHostDebug("drop-error", New With {
                    .sender = DescribeSender(sender),
                    .message = ex.Message
                })
                SendToWeb("host:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub OnDropOverlayHideTimerTick(sender As Object, e As EventArgs)
            _dropOverlayHideTimer.Stop()
            SendHostDebug("overlay-hide-timer", New With {
                .route = GetCurrentRouteKey(),
                .overlayVisible = _isNativeDropOverlayVisible
            })
            ApplyNativeDropOverlayVisibility(False)
        End Sub

        Private Sub OnNavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
            Try
                Dim paths = ExtractDroppedRvtPathsFromUri(e.Uri)
                SendHostDebug("navigation-starting", New With {
                    .uri = TrimForLog(e.Uri),
                    .pathCount = paths.Count,
                    .route = GetCurrentRouteKey()
                })
                If paths.Count = 0 Then Return

                e.Cancel = True
                DispatchDroppedRvts(GetCurrentRouteKey(), paths)
            Catch ex As Exception
                SendHostDebug("navigation-error", New With {
                    .message = ex.Message,
                    .uri = TrimForLog(e.Uri)
                })
                SendToWeb("host:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub OnNewWindowRequested(sender As Object, e As CoreWebView2NewWindowRequestedEventArgs)
            Try
                Dim paths = ExtractDroppedRvtPathsFromUri(e.Uri)
                SendHostDebug("new-window-requested", New With {
                    .uri = TrimForLog(e.Uri),
                    .pathCount = paths.Count,
                    .route = GetCurrentRouteKey()
                })
                If paths.Count = 0 Then Return

                e.Handled = True
                DispatchDroppedRvts(GetCurrentRouteKey(), paths)
            Catch ex As Exception
                SendHostDebug("new-window-error", New With {
                    .message = ex.Message,
                    .uri = TrimForLog(e.Uri)
                })
                SendToWeb("host:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub BroadcastDocumentList()
            Dim docs As New List(Of Object)()

            Try
                If _uiApp IsNot Nothing AndAlso _uiApp.Application IsNot Nothing Then
                    For Each d As Document In _uiApp.Application.Documents
                        Try
                            Dim name = d.Title
                            Dim path = d.PathName
                            If String.IsNullOrWhiteSpace(path) Then path = name
                            docs.Add(New With {.name = name, .path = path})
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try

            SendToWeb("host:doc-list", docs)
        End Sub

        Private Sub SendActiveDocument()
            SendToWeb("host:doc-changed", New With {.name = _currentDocName, .path = _currentDocPath})
        End Sub

        Private Shared Function HasFileDrop(data As IDataObject) As Boolean
            Try
                Return data IsNot Nothing AndAlso data.GetDataPresent(DataFormats.FileDrop)
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ExtractDroppedRvtPaths(data As IDataObject) As List(Of String)
            If Not HasFileDrop(data) Then Return New List(Of String)()

            Dim raw = TryCast(data.GetData(DataFormats.FileDrop), String())
            If raw Is Nothing OrElse raw.Length = 0 Then Return New List(Of String)()

            Return raw _
                .Select(AddressOf NormalizeDroppedRvtPath) _
                .Where(Function(path) Not String.IsNullOrWhiteSpace(path)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
        End Function

        Private Shared Function ExtractDroppedRvtPathsFromUri(uriText As String) As List(Of String)
            Dim raw = If(uriText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(raw) Then Return New List(Of String)()

            Dim candidates As IEnumerable(Of String) =
                raw.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries) _
                    .Select(Function(item) item.Trim()) _
                    .Where(Function(item) item.Length > 0)

            Dim results As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each item In candidates
                Dim localPath As String = Nothing

                Dim uri As Uri = Nothing
                If Uri.TryCreate(item, UriKind.Absolute, uri) AndAlso uri.IsFile Then
                    localPath = uri.LocalPath
                ElseIf item.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        localPath = New Uri(item).LocalPath
                    Catch
                        localPath = Nothing
                    End Try
                End If

                Dim normalized = NormalizeDroppedRvtPath(localPath)
                If String.IsNullOrWhiteSpace(normalized) Then Continue For
                If seen.Add(normalized) Then results.Add(normalized)
            Next

            Return results
        End Function

        Private Shared Function NormalizeDroppedRvtPath(pathText As String) As String
            Dim text = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return Nothing
            If Directory.Exists(text) Then Return Nothing
            If Not text.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase) Then Return Nothing
            If Not File.Exists(text) Then Return Nothing

            Try
                Return System.IO.Path.GetFullPath(text)
            Catch
                Return text
            End Try
        End Function

        Private Function GetCurrentRouteKey() As String
            If Not String.IsNullOrWhiteSpace(_currentRouteKey) Then
                Return NormalizeRouteKey(_currentRouteKey)
            End If

            Dim uriText As String = Nothing

            Try
                If _web.Source IsNot Nothing Then uriText = _web.Source.AbsoluteUri
            Catch
            End Try

            If String.IsNullOrWhiteSpace(uriText) Then
                Try
                    Dim core = _web.CoreWebView2
                    If core IsNot Nothing Then uriText = core.Source
                Catch
                End Try
            End If

            If String.IsNullOrWhiteSpace(uriText) Then Return String.Empty

            Dim uri As Uri = Nothing
            If Uri.TryCreate(uriText, UriKind.Absolute, uri) Then
                Return NormalizeRouteKey(If(uri.Fragment, String.Empty).TrimStart("#"c).Trim())
            End If

            Dim hashIndex = uriText.IndexOf("#"c)
            If hashIndex < 0 OrElse hashIndex >= uriText.Length - 1 Then Return String.Empty
            Return NormalizeRouteKey(uriText.Substring(hashIndex + 1).Trim())
        End Function

        Private Shared Function NormalizeRouteKey(routeKey As String) As String
            Return If(routeKey, String.Empty).Trim().TrimStart("#"c).ToLowerInvariant()
        End Function

        Private Shared Function SafeBoolObj(value As Object, fallback As Boolean) As Boolean
            Try
                If value Is Nothing Then Return fallback
                If TypeOf value Is Boolean Then Return CBool(value)
                Dim text = Convert.ToString(value)
                If String.IsNullOrWhiteSpace(text) Then Return fallback
                Select Case text.Trim().ToLowerInvariant()
                    Case "1", "true", "y", "yes", "on"
                        Return True
                    Case "0", "false", "n", "no", "off"
                        Return False
                End Select
            Catch
            End Try
            Return fallback
        End Function

        Private Sub SetNativeDropOverlay(active As Boolean, Optional immediateHide As Boolean = False)
            Dim routeKey = GetCurrentRouteKey()
            Dim shouldShow = active AndAlso SupportsDroppedRvtRoute(routeKey)
            SendHostDebug("set-overlay", New With {
                .requested = active,
                .immediateHide = immediateHide,
                .route = routeKey,
                .shouldShow = shouldShow,
                .overlayVisible = _isNativeDropOverlayVisible
            })
            If shouldShow Then
                _dropOverlayHideTimer.Stop()
                ApplyNativeDropOverlayVisibility(True)
                Return
            End If

            If immediateHide Then
                _dropOverlayHideTimer.Stop()
                ApplyNativeDropOverlayVisibility(False)
                Return
            End If

            If Not _isNativeDropOverlayVisible Then Return
            _dropOverlayHideTimer.Stop()
            _dropOverlayHideTimer.Start()
        End Sub

        Private Sub ApplyNativeDropOverlayVisibility(active As Boolean)
            If _isNativeDropOverlayVisible = active Then Return
            SendHostDebug("apply-overlay-visibility", New With {
                .active = active,
                .wasVisible = _isNativeDropOverlayVisible,
                .hasOverlayWindow = (_dropOverlayWindow IsNot Nothing)
            })
            _isNativeDropOverlayVisible = active
            If active Then
                EnsureDropOverlayWindow()
                SyncNativeDropOverlayBounds()
                If _dropOverlayWindow IsNot Nothing AndAlso Not _dropOverlayWindow.IsVisible Then
                    SendHostDebug("overlay-show", New With {
                        .left = Left,
                        .top = Top,
                        .width = ActualWidth,
                        .height = ActualHeight
                    })
                    _dropOverlayWindow.Show()
                End If
            ElseIf _dropOverlayWindow IsNot Nothing Then
                SendHostDebug("overlay-hide", New With {
                    .route = GetCurrentRouteKey(),
                    .wasVisible = _dropOverlayWindow.IsVisible
                })
                Try
                    _dropOverlayWindow.Close()
                Catch
                End Try
                _dropOverlayWindow = Nothing
            End If
            SendToWeb("host:rvt-drop-overlay", New With {.active = active})
        End Sub

        Private Sub EnsureDropOverlayWindow()
            If _dropOverlayWindow IsNot Nothing Then Return

            _dropOverlayWindow = New NativeDropOverlayWindow(Me)
            SendHostDebug("overlay-window-created", New With {
                .ownerTitle = Me.Title
            })
            AddHandler _dropOverlayWindow.DragEnter, AddressOf HandlePreviewDragEnter
            AddHandler _dropOverlayWindow.DragOver, AddressOf HandlePreviewDragOver
            AddHandler _dropOverlayWindow.DragLeave, AddressOf HandlePreviewDragLeave
            AddHandler _dropOverlayWindow.Drop, AddressOf HandlePreviewDrop
        End Sub

        Private Sub SyncNativeDropOverlayBounds()
            If _dropOverlayWindow Is Nothing Then Return
            If WindowState = WindowState.Minimized Then Return

            _dropOverlayWindow.Owner = Me
            _dropOverlayWindow.Topmost = Topmost
            _dropOverlayWindow.Left = Left
            _dropOverlayWindow.Top = Top
            _dropOverlayWindow.Width = ActualWidth
            _dropOverlayWindow.Height = ActualHeight
            SendHostDebug("overlay-bounds-sync", New With {
                .left = Left,
                .top = Top,
                .width = ActualWidth,
                .height = ActualHeight,
                .topmost = Topmost
            })
        End Sub

        Private Sub HandlePreviewDragLeave(sender As Object, e As DragEventArgs)
            SendHostDebug("drag-leave", New With {
                .sender = DescribeSender(sender),
                .route = GetCurrentRouteKey(),
                .overlayVisible = _isNativeDropOverlayVisible
            })
            SetNativeDropOverlay(False, False)
            e.Handled = True
        End Sub

        Private Shared Function SupportsDroppedRvtRoute(routeKey As String) As Boolean
            Select Case NormalizeRouteKey(routeKey)
                Case "multi", "sharedparambatch", "deliverycleaner", "familylink", "guid", "segmentpms", "export"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Sub DispatchDroppedRvts(routeKey As String, paths As IList(Of String))
            If paths Is Nothing OrElse paths.Count = 0 Then Return
            ClearPendingNativeDropPaths()
            SendHostDebug("dispatch-dropped-rvts", New With {
                .route = NormalizeRouteKey(routeKey),
                .pathCount = paths.Count,
                .firstPath = paths(0)
            })

            Select Case NormalizeRouteKey(routeKey)
                Case "multi"
                    SendToWeb("hub:rvt-picked", New With {.paths = paths})
                Case "sharedparambatch"
                    SendToWeb("sharedparambatch:rvts-picked", New With {.ok = True, .paths = paths})
                Case "deliverycleaner"
                    SendToWeb("deliverycleaner:rvts-picked", New With {.ok = True, .paths = paths})
                Case "familylink"
                    SendToWeb("familylink:rvts-picked", New With {.paths = paths})
                Case "guid"
                    SendToWeb("guid:files", New With {.paths = paths})
                Case "segmentpms"
                    SendToWeb("segmentpms:rvt-picked-files", New With {.paths = paths})
                Case "export"
                    SendToWeb("export:rvt-files", New With {.files = paths})
            End Select
        End Sub

        ' .NET → JS (양쪽 호환: ev & name 둘 다 포함해서 송신)
        Public Sub SendToWeb(ev As String, payload As Object)
            Dim core = _web.CoreWebView2
            If core Is Nothing Then Return

            Dim msg As New Dictionary(Of String, Object) From {
                {"ev", ev},
                {"name", ev},
                {"payload", payload}
            }

            Dim json = _serializer.Serialize(msg)
            core.PostWebMessageAsJson(json)
        End Sub

        Private Sub SendHostDebug(message As String, Optional payload As Object = Nothing)
            SendToWeb("host:debug", New With {
                .message = message,
                .payload = payload
            })
        End Sub

        Private Shared Function DescribeSender(sender As Object) As String
            If sender Is Nothing Then Return "(null)"
            Return sender.GetType().FullName
        End Function

        Private Shared Function GetDataFormats(data As IDataObject) As String()
            Try
                If data Is Nothing Then Return Array.Empty(Of String)()
                Return data.GetFormats()
            Catch
                Return Array.Empty(Of String)()
            End Try
        End Function

        Private Shared Function TrimForLog(text As String, Optional maxLength As Integer = 220) As String
            Dim value = If(text, String.Empty)
            If value.Length <= maxLength Then Return value
            Return value.Substring(0, maxLength) & "..."
        End Function

        Private Sub CachePendingNativeDropPaths(routeKey As String, paths As IList(Of String))
            If paths Is Nothing OrElse paths.Count = 0 Then Return
            _pendingNativeDropPaths = paths _
                .Where(Function(path) Not String.IsNullOrWhiteSpace(path)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
            _pendingNativeDropRouteKey = NormalizeRouteKey(routeKey)
            _pendingNativeDropStampUtc = DateTime.UtcNow
        End Sub

        Private Function ConsumePendingNativeDropPaths(routeKey As String) As List(Of String)
            Dim normalizedRoute = NormalizeRouteKey(routeKey)
            Dim isFresh = _pendingNativeDropStampUtc <> DateTime.MinValue AndAlso
                (DateTime.UtcNow - _pendingNativeDropStampUtc) <= TimeSpan.FromSeconds(5)
            If Not isFresh Then
                ClearPendingNativeDropPaths()
                Return New List(Of String)()
            End If
            If Not String.Equals(_pendingNativeDropRouteKey, normalizedRoute, StringComparison.OrdinalIgnoreCase) Then
                Return New List(Of String)()
            End If

            Dim result = _pendingNativeDropPaths.ToList()
            ClearPendingNativeDropPaths()
            Return result
        End Function

        Private Sub ClearPendingNativeDropPaths()
            _pendingNativeDropPaths.Clear()
            _pendingNativeDropRouteKey = String.Empty
            _pendingNativeDropStampUtc = DateTime.MinValue
        End Sub

    End Class

    Friend Class NativeDropOverlayWindow
        Inherits Window

        Public Sub New(owner As Window)
            Owner = owner
            WindowStyle = WindowStyle.None
            ResizeMode = ResizeMode.NoResize
            ShowInTaskbar = False
            ShowActivated = False
            AllowsTransparency = True
            Background = Brushes.Transparent
            Opacity = 0.01
            AllowDrop = True
            Content = New Border With {
                .Background = Brushes.Transparent
            }
        End Sub
    End Class

End Namespace
