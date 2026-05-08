Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports WinForms = System.Windows.Forms

Namespace Infrastructure

    Friend NotInheritable Class ExplorerFolderPicker

        Private Const HResultCancelled As Integer = -2147023673

        Private Sub New()
        End Sub

        Public Shared Function TryPickFolder(title As String,
                                             initialFolder As String,
                                             ByRef selectedPath As String) As Boolean
            selectedPath = String.Empty

            Try
                If TryPickFolderWithExplorerDialog(title, initialFolder, selectedPath) Then Return True
                Return False
            Catch
                Return TryPickFolderWithFallbackDialog(title, initialFolder, selectedPath)
            End Try
        End Function

        Private Shared Function TryPickFolderWithExplorerDialog(title As String,
                                                               initialFolder As String,
                                                               ByRef selectedPath As String) As Boolean
            selectedPath = String.Empty

            Dim dialogObject As Object = Nothing
            Dim folderItem As IShellItem = Nothing
            Dim resultItem As IShellItem = Nothing

            Try
                dialogObject = New FileOpenDialog()
                Dim dialog = DirectCast(dialogObject, IFileDialog)

                Dim options As FileOpenOptions = FileOpenOptions.None
                ThrowIfFailed(dialog.GetOptions(options))
                options = options Or FileOpenOptions.PickFolders Or FileOpenOptions.ForceFileSystem Or FileOpenOptions.PathMustExist
                ThrowIfFailed(dialog.SetOptions(options))

                If Not String.IsNullOrWhiteSpace(title) Then
                    ThrowIfFailed(dialog.SetTitle(title.Trim()))
                End If
                ThrowIfFailed(dialog.SetOkButtonLabel("폴더 선택"))

                Dim resolvedInitialFolder = ResolveInitialFolder(initialFolder)
                If TryCreateShellItem(resolvedInitialFolder, folderItem) Then
                    ThrowIfFailed(dialog.SetFolder(folderItem))
                End If

                Dim showResult = dialog.Show(IntPtr.Zero)
                If showResult = HResultCancelled Then Return False
                ThrowIfFailed(showResult)

                ThrowIfFailed(dialog.GetResult(resultItem))
                selectedPath = GetShellItemPath(resultItem)
                Return Not String.IsNullOrWhiteSpace(selectedPath)
            Finally
                ReleaseComObject(resultItem)
                ReleaseComObject(folderItem)
                ReleaseComObject(dialogObject)
            End Try
        End Function

        Private Shared Function TryPickFolderWithFallbackDialog(title As String,
                                                               initialFolder As String,
                                                               ByRef selectedPath As String) As Boolean
            selectedPath = String.Empty

            Using dlg As New WinForms.FolderBrowserDialog()
                dlg.Description = If(String.IsNullOrWhiteSpace(title), "폴더 선택", title.Trim())

                Dim resolvedInitialFolder = ResolveInitialFolder(initialFolder)
                If Not String.IsNullOrWhiteSpace(resolvedInitialFolder) Then
                    dlg.SelectedPath = resolvedInitialFolder
                End If

                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return False
                selectedPath = dlg.SelectedPath
                Return Not String.IsNullOrWhiteSpace(selectedPath)
            End Using
        End Function

        Private Shared Function ResolveInitialFolder(pathText As String) As String
            Dim value = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            Try
                If Directory.Exists(value) Then Return value
                If File.Exists(value) Then
                    Dim parent = Path.GetDirectoryName(value)
                    If Not String.IsNullOrWhiteSpace(parent) AndAlso Directory.Exists(parent) Then Return parent
                End If

                Dim candidateParent = Path.GetDirectoryName(value)
                If Not String.IsNullOrWhiteSpace(candidateParent) AndAlso Directory.Exists(candidateParent) Then Return candidateParent
            Catch
            End Try

            Return String.Empty
        End Function

        Private Shared Function TryCreateShellItem(folderPath As String, ByRef shellItem As IShellItem) As Boolean
            shellItem = Nothing
            If String.IsNullOrWhiteSpace(folderPath) OrElse Not Directory.Exists(folderPath) Then Return False

            Dim shellItemId = GetType(IShellItem).GUID
            Dim hr = SHCreateItemFromParsingName(folderPath, IntPtr.Zero, shellItemId, shellItem)
            If hr <> 0 OrElse shellItem Is Nothing Then Return False
            Return True
        End Function

        Private Shared Function GetShellItemPath(item As IShellItem) As String
            If item Is Nothing Then Return String.Empty

            Dim pathPointer As IntPtr = IntPtr.Zero
            Try
                ThrowIfFailed(item.GetDisplayName(SigDn.FileSysPath, pathPointer))
                Return If(Marshal.PtrToStringUni(pathPointer), String.Empty)
            Finally
                If pathPointer <> IntPtr.Zero Then Marshal.FreeCoTaskMem(pathPointer)
            End Try
        End Function

        Private Shared Sub ThrowIfFailed(hr As Integer)
            If hr <> 0 Then Marshal.ThrowExceptionForHR(hr)
        End Sub

        Private Shared Sub ReleaseComObject(instance As Object)
            If instance Is Nothing Then Return
            Try
                If Marshal.IsComObject(instance) Then Marshal.FinalReleaseComObject(instance)
            Catch
            End Try
        End Sub

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=True)>
        Private Shared Function SHCreateItemFromParsingName(
            <MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
            pbc As IntPtr,
            ByRef riid As Guid,
            <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IShellItem) As Integer
        End Function

        <ComImport>
        <Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")>
        Private Class FileOpenDialog
        End Class

        <ComImport>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        <Guid("42F85136-DB7E-439C-85F1-E4075D135FC8")>
        Private Interface IFileDialog
            <PreserveSig>
            Function Show(hwndOwner As IntPtr) As Integer

            <PreserveSig>
            Function SetFileTypes(cFileTypes As UInteger, rgFilterSpec As IntPtr) As Integer

            <PreserveSig>
            Function SetFileTypeIndex(iFileType As UInteger) As Integer

            <PreserveSig>
            Function GetFileTypeIndex(ByRef piFileType As UInteger) As Integer

            <PreserveSig>
            Function Advise(pfde As IntPtr, ByRef pdwCookie As UInteger) As Integer

            <PreserveSig>
            Function Unadvise(dwCookie As UInteger) As Integer

            <PreserveSig>
            Function SetOptions(fos As FileOpenOptions) As Integer

            <PreserveSig>
            Function GetOptions(ByRef pfos As FileOpenOptions) As Integer

            <PreserveSig>
            Function SetDefaultFolder(psi As IShellItem) As Integer

            <PreserveSig>
            Function SetFolder(psi As IShellItem) As Integer

            <PreserveSig>
            Function GetFolder(ByRef ppsi As IShellItem) As Integer

            <PreserveSig>
            Function GetCurrentSelection(ByRef ppsi As IShellItem) As Integer

            <PreserveSig>
            Function SetFileName(<MarshalAs(UnmanagedType.LPWStr)> pszName As String) As Integer

            <PreserveSig>
            Function GetFileName(ByRef pszName As IntPtr) As Integer

            <PreserveSig>
            Function SetTitle(<MarshalAs(UnmanagedType.LPWStr)> pszTitle As String) As Integer

            <PreserveSig>
            Function SetOkButtonLabel(<MarshalAs(UnmanagedType.LPWStr)> pszText As String) As Integer

            <PreserveSig>
            Function SetFileNameLabel(<MarshalAs(UnmanagedType.LPWStr)> pszLabel As String) As Integer

            <PreserveSig>
            Function GetResult(ByRef ppsi As IShellItem) As Integer

            <PreserveSig>
            Function AddPlace(psi As IShellItem, fdap As Integer) As Integer

            <PreserveSig>
            Function SetDefaultExtension(<MarshalAs(UnmanagedType.LPWStr)> pszDefaultExtension As String) As Integer

            <PreserveSig>
            Function Close(hr As Integer) As Integer

            <PreserveSig>
            Function SetClientGuid(ByRef guid As Guid) As Integer

            <PreserveSig>
            Function ClearClientData() As Integer

            <PreserveSig>
            Function SetFilter(pFilter As IntPtr) As Integer
        End Interface

        <ComImport>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        <Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")>
        Private Interface IShellItem
            <PreserveSig>
            Function BindToHandler(pbc As IntPtr, ByRef bhid As Guid, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer

            <PreserveSig>
            Function GetParent(ByRef ppsi As IShellItem) As Integer

            <PreserveSig>
            Function GetDisplayName(sigdnName As SigDn, ByRef ppszName As IntPtr) As Integer

            <PreserveSig>
            Function GetAttributes(sfgaoMask As UInteger, ByRef psfgaoAttribs As UInteger) As Integer

            <PreserveSig>
            Function Compare(psi As IShellItem, hint As UInteger, ByRef piOrder As Integer) As Integer
        End Interface

        <Flags>
        Private Enum FileOpenOptions As UInteger
            None = 0UI
            PickFolders = &H20UI
            ForceFileSystem = &H40UI
            PathMustExist = &H800UI
        End Enum

        Private Enum SigDn As UInteger
            FileSysPath = &H80058000UI
        End Enum

    End Class

End Namespace
