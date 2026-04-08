Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Models
Imports KKY_Tool_Revit.Services
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared ReadOnly _conditionExtractLock As New Object()
        Private Shared _conditionExtractSettings As ConditionExtractService.Settings
        Private Shared _conditionExtractLastResult As ConditionExtractService.RunResult

        Private Sub HandleConditionExtractInit(app As UIApplication, payload As Object)
            SendToWeb("conditionextract:init", BuildConditionExtractStatePayload(app))
        End Sub

        Private Sub HandleConditionExtractPickRvts(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                SendToWebAfterDialog("conditionextract:rvts-picked", New With {.ok = True, .paths = dlg.FileNames})
            End Using
        End Sub

        Private Sub HandleConditionExtractRun(app As UIApplication, payload As Object)
            Dim settings = ParseConditionExtractSettings(payload)
            SyncLock _conditionExtractLock
                _conditionExtractSettings = settings
                _conditionExtractLastResult = Nothing
            End SyncLock

            Dim progress As IProgress(Of Object) = New Progress(Of Object)(Sub(snapshot) SendToWeb("conditionextract:progress", snapshot))

            Try
                Dim result = ConditionExtractService.Run(app, settings, progress)
                SyncLock _conditionExtractLock
                    _conditionExtractLastResult = result
                End SyncLock

                SendToWeb("conditionextract:done", New With {
                    .ok = result.Ok,
                    .message = result.Message,
                    .outputFolder = result.OutputFolder,
                    .summary = result.Summary,
                    .fileCount = result.Files.Count,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .logTextPath = result.LogTextPath,
                    .activeDocument = New With {
                        .title = app.ActiveUIDocument?.Document?.Title,
                        .isWorkshared = If(app.ActiveUIDocument?.Document IsNot Nothing, app.ActiveUIDocument.Document.IsWorkshared, False)
                    }
                })
            Catch ex As Exception
                SyncLock _conditionExtractLock
                    _conditionExtractLastResult = Nothing
                End SyncLock
                SendToWeb("conditionextract:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleConditionExtractExportResults(app As UIApplication, payload As Object)
            Try
                Dim settings As ConditionExtractService.Settings = Nothing
                Dim result As ConditionExtractService.RunResult = Nothing

                SyncLock _conditionExtractLock
                    settings = _conditionExtractSettings
                    result = _conditionExtractLastResult
                End SyncLock

                If result Is Nothing Then
                    SendToWeb("conditionextract:artifacts-exported", New With {.ok = False, .message = "최근 실행 결과가 없습니다."})
                    Return
                End If

                ConditionExtractService.ExportArtifacts(app, settings, result)
                SendToWeb("conditionextract:artifacts-exported", New With {
                    .ok = True,
                    .message = "결과 엑셀과 로그를 저장했습니다.",
                    .outputFolder = result.OutputFolder,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .logTextPath = result.LogTextPath
                })
            Catch ex As Exception
                SendToWeb("conditionextract:artifacts-exported", New With {.ok = False, .message = ex.Message})
            End Try
        End Sub

        Private Sub HandleConditionExtractOpenFolder(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim pathText = Convert.ToString(GetProp(pd, "path"))
            If String.IsNullOrWhiteSpace(pathText) Then
                SyncLock _conditionExtractLock
                    If _conditionExtractLastResult IsNot Nothing Then pathText = _conditionExtractLastResult.OutputFolder
                    If String.IsNullOrWhiteSpace(pathText) AndAlso _conditionExtractSettings IsNot Nothing Then pathText = _conditionExtractSettings.OutputFolder
                End SyncLock
            End If

            If String.IsNullOrWhiteSpace(pathText) Then
                SendToWeb("conditionextract:error", New With {.message = "열 폴더 경로가 없습니다."})
                Return
            End If

            Dim targetPath = pathText
            If File.Exists(targetPath) Then targetPath = Path.GetDirectoryName(targetPath)
            If String.IsNullOrWhiteSpace(targetPath) OrElse Not Directory.Exists(targetPath) Then
                SendToWeb("conditionextract:error", New With {.message = "열 폴더 경로를 찾을 수 없습니다."})
                Return
            End If

            Dim psi As New ProcessStartInfo("explorer.exe", """" & targetPath & """")
            psi.UseShellExecute = True
            Process.Start(psi)
            SendToWeb("conditionextract:folder-opened", New With {.ok = True, .path = targetPath})
        End Sub

        Private Shared Function ParseConditionExtractSettings(payload As Object) As ConditionExtractService.Settings
            Dim pd = ParsePayloadDict(payload)
            Dim settings As New ConditionExtractService.Settings()
            settings.UseActiveDocument = SafeBoolObj(GetProp(pd, "useActiveDocument"), True)
            settings.RvtPaths = ParseStringList(pd, "rvtPaths")
            settings.OutputFolder = NormalizeDeliveryCleanerText(GetProp(pd, "outputFolder"))
            settings.CloseAllWorksetsOnOpen = SafeBoolObj(GetProp(pd, "closeAllWorksetsOnOpen"), True)
            settings.ElementParameterUpdate = ParseDeliveryCleanerElementUpdate(GetProp(pd, "elementParameterUpdate"))
            settings.IncludeCoordinates = SafeBoolObj(GetProp(pd, "includeCoordinates"), False)
            settings.IncludeLinearMetrics = SafeBoolObj(GetProp(pd, "includeLinearMetrics"), False)
            settings.ExtractParameterNames = ParseConditionExtractNames(GetProp(pd, "extractParameterNamesCsv"))
            settings.LengthUnit = NormalizeConditionExtractUnit(GetProp(pd, "lengthUnit"), "mm")
            settings.AreaUnit = NormalizeConditionExtractUnit(GetProp(pd, "areaUnit"), "mm2")
            settings.VolumeUnit = NormalizeConditionExtractUnit(GetProp(pd, "volumeUnit"), "mm3")
            Return settings
        End Function

        Private Shared Function BuildConditionExtractStatePayload(app As UIApplication) As Object
            Dim settings As ConditionExtractService.Settings = Nothing
            Dim result As ConditionExtractService.RunResult = Nothing

            SyncLock _conditionExtractLock
                settings = _conditionExtractSettings
                result = _conditionExtractLastResult
            End SyncLock

            If settings Is Nothing Then
                settings = New ConditionExtractService.Settings() With {
                    .UseActiveDocument = True,
                    .CloseAllWorksetsOnOpen = True,
                    .ElementParameterUpdate = New ElementParameterUpdateSettings()
                }
            End If

            Return New With {
                .settings = New With {
                    .useActiveDocument = settings.UseActiveDocument,
                    .rvtPaths = If(settings.RvtPaths, New List(Of String)()),
                    .outputFolder = settings.OutputFolder,
                    .closeAllWorksetsOnOpen = settings.CloseAllWorksetsOnOpen,
                    .elementParameterUpdate = SerializeElementParameterUpdateSettings(settings.ElementParameterUpdate),
                    .extractParameterNamesCsv = String.Join(", ", If(settings.ExtractParameterNames, New List(Of String)())),
                    .includeCoordinates = settings.IncludeCoordinates,
                    .includeLinearMetrics = settings.IncludeLinearMetrics,
                    .lengthUnit = settings.LengthUnit,
                    .areaUnit = settings.AreaUnit,
                    .volumeUnit = settings.VolumeUnit
                },
                .result = If(result Is Nothing, Nothing, New With {
                    .ok = result.Ok,
                    .message = result.Message,
                    .outputFolder = result.OutputFolder,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .logTextPath = result.LogTextPath,
                    .summary = result.Summary,
                    .fileCount = result.Files.Count
                }),
                .activeDocument = New With {
                    .title = app.ActiveUIDocument?.Document?.Title,
                    .isWorkshared = If(app.ActiveUIDocument?.Document IsNot Nothing, app.ActiveUIDocument.Document.IsWorkshared, False)
                }
            }
        End Function

        Private Shared Function ParseConditionExtractNames(raw As Object) As List(Of String)
            Return Convert.ToString(raw) _
                .Split(New String() {",", ";", vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries) _
                .Select(Function(x) x.Trim()) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()
        End Function

        Private Shared Function NormalizeConditionExtractUnit(raw As Object, fallback As String) As String
            Dim text = NormalizeDeliveryCleanerText(raw).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(text) Then Return fallback
            Return text
        End Function

    End Class

End Namespace
