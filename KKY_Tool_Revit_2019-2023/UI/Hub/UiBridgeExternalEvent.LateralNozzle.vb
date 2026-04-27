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
                    .fileCount = result.Files.Count,
                    .canExport = result.Files IsNot Nothing AndAlso result.Files.Count > 0
                })
            Catch ex As Exception
                SyncLock _lateralNozzleLock
                    _lateralNozzleLastResult = Nothing
                End SyncLock
                SendToWeb("lateralnozzle:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleLateralNozzleExport(app As UIApplication, payload As Object)
            Dim result As UtilityLateralNozzleExtractService.RunResult = Nothing
            Dim doAutoFit As Boolean = ParseExcelMode(payload)
            Dim excelMode As String = If(doAutoFit, "normal", "fast")
            Dim exportLocale As String = ParseExcelLocale(payload)

            SyncLock _lateralNozzleLock
                result = _lateralNozzleLastResult
            End SyncLock

            If result Is Nothing OrElse result.Files Is Nothing OrElse result.Files.Count = 0 Then
                SendToWeb("lateralnozzle:error", New With {.message = "내보낼 결과가 없습니다. 먼저 추출을 실행해주세요."})
                Return
            End If

            Using dlg As New WinForms.SaveFileDialog()
                dlg.Filter = "Excel (*.xlsx)|*.xlsx"
                dlg.FileName = If(String.IsNullOrWhiteSpace(result.ResultWorkbookPath),
                                  UtilityLateralNozzleExtractService.GetDefaultExportFileName(),
                                  Path.GetFileName(result.ResultWorkbookPath))
                dlg.AddExtension = True
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then
                    SendToWeb("lateralnozzle:exported", New With {.ok = False, .cancelled = True})
                    Return
                End If

                Try
                    Dim exportPath = UtilityLateralNozzleExtractService.ExportWorkbook(result, dlg.FileName, doAutoFit, exportLocale)
                    SyncLock _lateralNozzleLock
                        _lateralNozzleLastResult = result
                    End SyncLock

                    SendToWeb("lateralnozzle:exported", New With {
                        .ok = True,
                        .path = exportPath,
                        .excelMode = excelMode
                    })
                Catch ex As Exception
                    SendToWeb("lateralnozzle:error", New With {.message = "엑셀 저장 중 오류가 발생했습니다: " & ex.Message})
                End Try
            End Using
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
                    .fileCount = result.Files.Count,
                    .canExport = result.Files IsNot Nothing AndAlso result.Files.Count > 0
                })
            }
        End Function

    End Class

End Namespace
