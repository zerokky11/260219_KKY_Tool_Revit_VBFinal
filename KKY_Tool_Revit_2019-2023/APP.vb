Imports System.Diagnostics
Imports System.Reflection
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.IO
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI.Events
Imports System.Linq
Imports KKY_Tool_Revit.UI
Imports KKY_Tool_Revit.UI.Hub
'커밋 확인용'
Namespace KKY_Tool_Revit
    Public Class App
        Implements IExternalApplication

        Public Function OnStartup(a As UIControlledApplication) As Result Implements IExternalApplication.OnStartup
            Const TAB_NAME As String = "KKY Tools"
            Const PANEL_NAME As String = "Hub"

            ' 탭 생성(이미 있으면 무시)
            Try : a.CreateRibbonTab(TAB_NAME) : Catch : End Try

            ' 패널 찾기/생성
            Dim ribbonPanel As RibbonPanel = a.GetRibbonPanels(TAB_NAME).FirstOrDefault(Function(p) p.Name = PANEL_NAME)
            If ribbonPanel Is Nothing Then
                ribbonPanel = a.CreateRibbonPanel(TAB_NAME, PANEL_NAME)
            End If

            ' 허브 버튼
            Dim asmPath = Assembly.GetExecutingAssembly().Location
            Dim cmdFullName As String = GetType(DuplicateExport).FullName
            Dim pbd As New PushButtonData("KKY_Hub_Button", "KKY Tools", asmPath, cmdFullName)

            Dim btn = TryCast(ribbonPanel.AddItem(pbd), PushButton)
            If btn IsNot Nothing Then
                btn.ToolTip = "KKY Tool 허브 열기"
                'Dim resNames = String.Join(vbCrLf, Assembly.GetExecutingAssembly().GetManifestResourceNames())
                'TaskDialog.Show("RES", resNames)

                Dim smallImg = LoadHubIcon(16,
                    "KKY_Tool_Revit.KKY_TOOL_64_source.png",
                    "KKY_Tool_Revit.KKY_Tool_16.png",
                    "KKY_Tool_Revit.KKY_Hub_16.png",
                    "KKY_Tool_Revit.hub_16.png")

                Dim largeImg = LoadHubIcon(32,
                    "KKY_Tool_Revit.KKY_TOOL_64_source.png",
                    "KKY_Tool_Revit.KKY_Tool_32.png",
                    "KKY_Tool_Revit.KKY_Hub_32.png",
                    "KKY_Tool_Revit.hub_32.png")

                btn.Image = smallImg
                btn.LargeImage = largeImg
            End If

            ' 활성 문서/뷰 변경 이벤트 구독
            AddHandler a.ViewActivated, AddressOf OnViewActivated
            AddHandler a.ControlledApplication.DocumentOpened, AddressOf OnDocumentListChanged
            AddHandler a.ControlledApplication.DocumentClosed, AddressOf OnDocumentListChanged
            AddHandler a.ControlledApplication.DocumentClosed, AddressOf OnDocumentClosedForActiveLinkWorksetReopen

            Return Result.Succeeded
        End Function

        Public Function OnShutdown(a As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
            Try
                RemoveHandler a.ViewActivated, AddressOf OnViewActivated
                RemoveHandler a.ControlledApplication.DocumentOpened, AddressOf OnDocumentListChanged
                RemoveHandler a.ControlledApplication.DocumentClosed, AddressOf OnDocumentListChanged
                RemoveHandler a.ControlledApplication.DocumentClosed, AddressOf OnDocumentClosedForActiveLinkWorksetReopen
            Catch
            End Try
            Return Result.Succeeded
        End Function

        Private Sub OnViewActivated(sender As Object, e As ViewActivatedEventArgs)
            Try
                HubHostWindow.NotifyActiveDocumentChanged(e.Document)
            Catch
            End Try
        End Sub

        Private Sub OnDocumentListChanged(sender As Object, e As EventArgs)
            Try
                HubHostWindow.NotifyDocumentListChanged()
            Catch
            End Try
        End Sub

        Private Sub OnDocumentClosedForActiveLinkWorksetReopen(sender As Object, e As Autodesk.Revit.DB.Events.DocumentClosedEventArgs)
            Try
                UiBridgeExternalEvent.NotifyActiveLinkWorksetDocumentClosed()
            Catch
            End Try
        End Sub

        Private Function LoadHubIcon(decodePixelWidth As Integer, ParamArray resourceNames() As String) As ImageSource
            For Each name In resourceNames
                Dim img = LoadPngImageSource(name, decodePixelWidth)
                If img IsNot Nothing Then Return img
            Next
            Return Nothing
        End Function

        Private Function LoadPngImageSource(resName As String, Optional decodePixelWidth As Integer = 0) As ImageSource
            Dim asm = Assembly.GetExecutingAssembly()
            Dim resolvedName = ResolveResourceName(asm, resName)
            If String.IsNullOrEmpty(resolvedName) Then
                Debug.WriteLine($"[KKY_Tool_Revit] 아이콘 리소스 로드 실패: {resName}")
                Return Nothing
            End If

            Using s = asm.GetManifestResourceStream(resolvedName)
                If s Is Nothing Then
                    Debug.WriteLine($"[KKY_Tool_Revit] 아이콘 리소스 로드 실패: {resName}")
                    Return Nothing
                End If
                Dim bmp As New BitmapImage()
                bmp.BeginInit()
                bmp.StreamSource = s
                bmp.CacheOption = BitmapCacheOption.OnLoad
                If decodePixelWidth > 0 Then
                    bmp.DecodePixelWidth = decodePixelWidth
                    bmp.DecodePixelHeight = decodePixelWidth
                End If
                bmp.EndInit()
                bmp.Freeze()
                Return bmp
            End Using
        End Function

        Private Function ResolveResourceName(asm As Assembly, desired As String) As String
            Dim names = asm.GetManifestResourceNames()

            Dim exact = names.FirstOrDefault(Function(n) String.Equals(n, desired, StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrEmpty(exact) Then Return exact

            Dim fileName = ExtractEmbeddedFileName(desired)
            Dim byFile = names.FirstOrDefault(Function(n) n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrEmpty(byFile) Then Return byFile

            Return Nothing
        End Function

        Private Function ExtractEmbeddedFileName(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return value

            Dim extIndex = value.LastIndexOf("."c)
            If extIndex <= 0 Then Return value

            Dim nameIndex = value.LastIndexOf("."c, extIndex - 1)
            If nameIndex < 0 Then Return value

            Return value.Substring(nameIndex + 1)
        End Function
    End Class
End Namespace
