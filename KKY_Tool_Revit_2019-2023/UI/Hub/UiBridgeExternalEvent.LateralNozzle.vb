Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared ReadOnly _lateralNozzleLock As New Object()
        Private Shared _lateralNozzleSettings As UtilityLateralNozzleExtractService.Settings
        Private Shared _lateralNozzleLastResult As UtilityLateralNozzleExtractService.RunResult

        Private Sub HandleLateralNozzleInit(app As UIApplication, payload As Object)
            SendToWeb("lateralnozzle:init", BuildLateralNozzleStatePayload())
        End Sub

        Private Sub HandleLateralNozzlePickExcels(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
                dlg.Multiselect = True
                dlg.Title = "엑셀 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                SendToWebAfterDialog("lateralnozzle:excels-picked", New With {
                    .ok = True,
                    .paths = dlg.FileNames
                })
            End Using
        End Sub

        Private Sub HandleLateralNozzleRun(app As UIApplication, payload As Object)
            Dim settings = ParseLateralNozzleSettings(payload)
            SyncLock _lateralNozzleLock
                _lateralNozzleSettings = settings
                _lateralNozzleLastResult = Nothing
            End SyncLock

            Dim progress As IProgress(Of Object) = New Progress(Of Object)(Sub(snapshot) SendToWeb("lateralnozzle:progress", snapshot))

            Try
                Dim result = UtilityLateralNozzleExtractService.Run(settings, progress)
                SyncLock _lateralNozzleLock
                    _lateralNozzleLastResult = result
                End SyncLock

                SendToWeb("lateralnozzle:done", New With {
                    .ok = result.Ok,
                    .message = result.Message,
                    .outputFolder = result.OutputFolder,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .summary = result.Summary,
                    .fileCount = result.Files.Count
                })
            Catch ex As Exception
                SyncLock _lateralNozzleLock
                    _lateralNozzleLastResult = Nothing
                End SyncLock
                SendToWeb("lateralnozzle:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleLateralNozzleOpenFolder(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim pathText = Convert.ToString(GetProp(pd, "path"))

            If String.IsNullOrWhiteSpace(pathText) Then
                SyncLock _lateralNozzleLock
                    If _lateralNozzleLastResult IsNot Nothing Then pathText = _lateralNozzleLastResult.OutputFolder
                    If String.IsNullOrWhiteSpace(pathText) AndAlso _lateralNozzleSettings IsNot Nothing Then pathText = _lateralNozzleSettings.OutputFolder
                End SyncLock
            End If

            If String.IsNullOrWhiteSpace(pathText) Then
                SendToWeb("lateralnozzle:error", New With {.message = "열 폴더 경로가 없습니다."})
                Return
            End If

            Dim targetPath = pathText
            If File.Exists(targetPath) Then targetPath = Path.GetDirectoryName(targetPath)
            If String.IsNullOrWhiteSpace(targetPath) OrElse Not Directory.Exists(targetPath) Then
                SendToWeb("lateralnozzle:error", New With {.message = "열 폴더 경로를 찾을 수 없습니다."})
                Return
            End If

            Dim psi As New ProcessStartInfo("explorer.exe", """" & targetPath & """")
            psi.UseShellExecute = True
            Process.Start(psi)
            SendToWeb("lateralnozzle:folder-opened", New With {.ok = True, .path = targetPath})
        End Sub

        Private Shared Function ParseLateralNozzleSettings(payload As Object) As UtilityLateralNozzleExtractService.Settings
            Dim pd = ParsePayloadDict(payload)
            Dim settings As New UtilityLateralNozzleExtractService.Settings()
            settings.ExcelPaths = ParseStringList(pd, "excelPaths")
            settings.OutputFolder = Convert.ToString(GetProp(pd, "outputFolder"))
            If settings.OutputFolder Is Nothing Then settings.OutputFolder = String.Empty
            settings.OutputFolder = settings.OutputFolder.Trim()
            Return settings
        End Function

        Private Shared Function BuildLateralNozzleStatePayload() As Object
            Dim settings As UtilityLateralNozzleExtractService.Settings = Nothing
            Dim result As UtilityLateralNozzleExtractService.RunResult = Nothing

            SyncLock _lateralNozzleLock
                settings = _lateralNozzleSettings
                result = _lateralNozzleLastResult
            End SyncLock

            If settings Is Nothing Then
                settings = New UtilityLateralNozzleExtractService.Settings()
            End If

            Return New With {
                .settings = New With {
                    .excelPaths = If(settings.ExcelPaths, New List(Of String)()),
                    .outputFolder = settings.OutputFolder
                },
                .result = If(result Is Nothing, Nothing, New With {
                    .ok = result.Ok,
                    .message = result.Message,
                    .outputFolder = result.OutputFolder,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .summary = result.Summary,
                    .fileCount = result.Files.Count
                })
            }
        End Function

    End Class

End Namespace
