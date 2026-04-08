Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

Namespace Infrastructure

    Friend NotInheritable Class ThirdPartyWarningSuppressor

        Private Sub New()
        End Sub

        Private Shared ReadOnly SyncRoot As New Object()
        Private Shared _timer As Timer
        Private Shared _disposed As Boolean
        Private Shared _isTicking As Integer
        Private Shared _lastHandledHandle As IntPtr = IntPtr.Zero
        Private Shared _lastHandledAtUtc As DateTime = DateTime.MinValue

        Public Shared Sub Start()
            SyncLock SyncRoot
                If _disposed OrElse _timer IsNot Nothing Then Return
                _timer = New Timer(AddressOf OnTimerTick, Nothing, 1000, 800)
            End SyncLock
        End Sub

        Public Shared Sub [Stop]()
            SyncLock SyncRoot
                _disposed = True

                If _timer IsNot Nothing Then
                    _timer.Dispose()
                    _timer = Nothing
                End If
            End SyncLock
        End Sub

        Private Shared Sub OnTimerTick(state As Object)
            If Interlocked.Exchange(_isTicking, 1) = 1 Then Return

            Try
                If _disposed Then Return

                Dim targetDialog = FindTargetDialog()
                If targetDialog = IntPtr.Zero Then Return

                If targetDialog = _lastHandledHandle AndAlso (DateTime.UtcNow - _lastHandledAtUtc).TotalSeconds < 5 Then
                    Return
                End If

                If TryClickDoNotWarnButton(targetDialog) Then
                    _lastHandledHandle = targetDialog
                    _lastHandledAtUtc = DateTime.UtcNow
                    Debug.WriteLine("[KKY_Tool_Revit] Auto-confirmed third-party warning with 'Do not warn'.")
                End If
            Catch ex As Exception
                Debug.WriteLine("[KKY_Tool_Revit] ThirdPartyWarningSuppressor error: " & ex.Message)
            Finally
                Interlocked.Exchange(_isTicking, 0)
            End Try
        End Sub

        Private Shared Function FindTargetDialog() As IntPtr
            Dim found As IntPtr = IntPtr.Zero
            Dim currentProcessId = Process.GetCurrentProcess().Id

            EnumWindows(Function(hWnd, lParam)
                            If found <> IntPtr.Zero Then Return False
                            If Not IsWindowVisible(hWnd) Then Return True

                            Dim processId As UInteger = 0UI
                            GetWindowThreadProcessId(hWnd, processId)
                            If processId <> CUInt(currentProcessId) Then Return True

                            Dim className = GetWindowClassName(hWnd)
                            If Not String.Equals(className, "#32770", StringComparison.Ordinal) Then Return True

                            Dim dialogTitle = GetWindowTextSafe(hWnd)
                            Dim dialogText = String.Join(Environment.NewLine, GetDescendantTexts(hWnd))

                            If Not IsThirdPartyWarning(dialogTitle, dialogText) Then Return True
                            If Not HasDoNotWarnButton(hWnd) Then Return True

                            found = hWnd
                            Return False
                        End Function, IntPtr.Zero)

            Return found
        End Function

        Private Shared Function IsThirdPartyWarning(title As String, body As String) As Boolean
            Dim combined = (If(title, String.Empty) & vbLf & If(body, String.Empty)).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(combined) Then Return False

            Dim hasThirdParty = combined.Contains("third-party") OrElse
                                combined.Contains("third party") OrElse
                                combined.Contains("thirdparty") OrElse
                                combined.Contains("서드파티")
            If Not hasThirdParty Then Return False

            Return combined.Contains("warn") OrElse
                   combined.Contains("warning") OrElse
                   combined.Contains("do not warn") OrElse
                   combined.Contains("경고")
        End Function

        Private Shared Function HasDoNotWarnButton(dialogHandle As IntPtr) As Boolean
            Return GetChildButtons(dialogHandle).Exists(Function(btn) IsDoNotWarnLabel(btn.Text))
        End Function

        Private Shared Function TryClickDoNotWarnButton(dialogHandle As IntPtr) As Boolean
            For Each button In GetChildButtons(dialogHandle)
                If Not IsDoNotWarnLabel(button.Text) Then Continue For

                SendMessage(button.Handle, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                Return True
            Next

            Return False
        End Function

        Private Shared Function IsDoNotWarnLabel(text As String) As Boolean
            Dim normalized = NormalizeText(text)
            If String.IsNullOrWhiteSpace(normalized) Then Return False

            Return normalized.Contains("do not warn") OrElse
                   normalized.Contains("don't warn") OrElse
                   normalized.Contains("dont warn") OrElse
                   normalized.Contains("do not show again") OrElse
                   normalized.Contains("do not show this message again")
        End Function

        Private Shared Function NormalizeText(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return String.Empty
            Return text.Replace(vbCr, " ").Replace(vbLf, " ").Trim().ToLowerInvariant()
        End Function

        Private Shared Function GetDescendantTexts(parentHandle As IntPtr) As List(Of String)
            Dim texts As New List(Of String)()

            EnumChildWindows(parentHandle,
                             Function(childHandle, lParam)
                                 Dim text = GetWindowTextSafe(childHandle)
                                 If Not String.IsNullOrWhiteSpace(text) Then
                                     texts.Add(text.Trim())
                                 End If
                                 Return True
                             End Function,
                             IntPtr.Zero)

            Return texts
        End Function

        Private Shared Function GetChildButtons(parentHandle As IntPtr) As List(Of WindowButton)
            Dim buttons As New List(Of WindowButton)()

            EnumChildWindows(parentHandle,
                             Function(childHandle, lParam)
                                 If Not IsWindowVisible(childHandle) Then Return True
                                 If Not String.Equals(GetWindowClassName(childHandle), "Button", StringComparison.OrdinalIgnoreCase) Then Return True

                                 buttons.Add(New WindowButton(childHandle, GetWindowTextSafe(childHandle)))
                                 Return True
                             End Function,
                             IntPtr.Zero)

            Return buttons
        End Function

        Private Shared Function GetWindowClassName(hWnd As IntPtr) As String
            Dim buffer As New StringBuilder(256)
            GetClassName(hWnd, buffer, buffer.Capacity)
            Return buffer.ToString()
        End Function

        Private Shared Function GetWindowTextSafe(hWnd As IntPtr) As String
            Dim length = GetWindowTextLength(hWnd)
            If length <= 0 Then Return String.Empty

            Dim buffer As New StringBuilder(length + 1)
            GetWindowText(hWnd, buffer, buffer.Capacity)
            Return buffer.ToString()
        End Function

        Private NotInheritable Class WindowButton

            Public Sub New(handle As IntPtr, text As String)
                Me.Handle = handle
                Me.Text = If(text, String.Empty)
            End Sub

            Public ReadOnly Property Handle As IntPtr
            Public ReadOnly Property Text As String
        End Class

        Private Const BM_CLICK As Integer = &HF5

        <DllImport("user32.dll")>
        Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function EnumChildWindows(hWndParent As IntPtr, lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

        <DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef processId As UInteger) As UInteger
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function GetClassName(hWnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function
    End Class
End Namespace
