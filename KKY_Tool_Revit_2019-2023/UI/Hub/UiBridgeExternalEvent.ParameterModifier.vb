Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Models
Imports KKY_Tool_Revit.Services
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared ReadOnly _parameterModifierLock As New Object()
        Private Shared _parameterModifierSettings As ParameterModifierService.Settings
        Private Shared _parameterModifierLastResult As ParameterModifierService.RunResult

        Private Sub HandleParameterModifierInit(app As UIApplication, payload As Object)
            SendToWeb("parammodifier:init", BuildParameterModifierStatePayload(app))
        End Sub

        Private Sub HandleParameterModifierPickRvts(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "RVT 파일 선택"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                SendToWebAfterDialog("parammodifier:rvts-picked", New With {.ok = True, .paths = dlg.FileNames})
            End Using
        End Sub

        Private Sub HandleParameterModifierBrowseOutputFolder(app As UIApplication, payload As Object)
            Using dlg As New WinForms.FolderBrowserDialog()
                dlg.Description = "파라미터 수정기 결과 폴더 선택"
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return
                SendToWebAfterDialog("parammodifier:output-folder-picked", New With {.ok = True, .path = dlg.SelectedPath})
            End Using
        End Sub

        Private Sub HandleParameterModifierFilterImport(app As UIApplication, payload As Object)
            Using dlg As New WinForms.OpenFileDialog()
                dlg.Filter = "XML (*.xml)|*.xml"
                dlg.Title = "View Filter XML 불러오기"
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                Dim profile = RevitViewFilterProfileService.LoadFromXml(dlg.FileName)
                SendToWebAfterDialog("parammodifier:filter-loaded", New With {
                    .ok = True,
                    .profile = SerializeFilterProfile(profile),
                    .source = dlg.FileName
                })
            End Using
        End Sub

        Private Sub HandleParameterModifierFilterSave(app As UIApplication, payload As Object)
            Dim profile = ParseDeliveryCleanerFilterProfile(GetProp(ParsePayloadDict(payload), "filterProfile"))
            If profile Is Nothing OrElse Not profile.IsConfigured() Then
                SendToWeb("parammodifier:error", New With {.message = "저장할 필터 설정이 올바르지 않습니다."})
                Return
            End If

            Using dlg As New WinForms.SaveFileDialog()
                dlg.Filter = "XML (*.xml)|*.xml"
                dlg.Title = "View Filter XML 저장"
                dlg.FileName = If(String.IsNullOrWhiteSpace(profile.FilterName), "ViewFilterProfile.xml", profile.FilterName & ".xml")
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                RevitViewFilterProfileService.SaveToXml(profile, dlg.FileName)
                SendToWeb("parammodifier:filter-saved", New With {.ok = True, .path = dlg.FileName})
            End Using
        End Sub

        Private Sub HandleParameterModifierFilterDocList(app As UIApplication, payload As Object)
            Dim doc = app.ActiveUIDocument?.Document
            If doc Is Nothing Then
                SendToWeb("parammodifier:filter-doc-list", New With {
                    .ok = False,
                    .message = "현재 열려 있는 문서를 찾을 수 없습니다.",
                    .items = New List(Of Object)()
                })
                Return
            End If

            Dim items = New FilteredElementCollector(doc) _
                .OfClass(GetType(ParameterFilterElement)) _
                .Cast(Of ParameterFilterElement)() _
                .OrderBy(Function(x) x.Name) _
                .Select(Function(x) New With {.id = x.Id.IntegerValue, .name = x.Name}) _
                .ToList()

            SendToWeb("parammodifier:filter-doc-list", New With {.ok = True, .items = items, .docTitle = doc.Title})
        End Sub

        Private Sub HandleParameterModifierFilterDocExtract(app As UIApplication, payload As Object)
            Dim doc = app.ActiveUIDocument?.Document
            If doc Is Nothing Then
                SendToWeb("parammodifier:error", New With {.message = "현재 열려 있는 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim pd = ParsePayloadDict(payload)
            Dim filterIdInt = SafeIntObj(GetProp(pd, "filterId"), Integer.MinValue)
            If filterIdInt = Integer.MinValue Then
                SendToWeb("parammodifier:error", New With {.message = "추출할 필터를 선택하세요."})
                Return
            End If

            Dim filterEl = TryCast(doc.GetElement(New ElementId(filterIdInt)), ParameterFilterElement)
            If filterEl Is Nothing Then
                SendToWeb("parammodifier:error", New With {.message = "선택한 필터를 찾을 수 없습니다."})
                Return
            End If

            Dim profile = RevitViewFilterProfileService.ExtractProfileFromFilter(doc, filterEl.Id)
            SendToWeb("parammodifier:filter-loaded", New With {
                .ok = True,
                .profile = SerializeFilterProfile(profile),
                .source = doc.Title
            })
        End Sub

        Private Sub HandleParameterModifierRun(app As UIApplication, payload As Object)
            Dim settings = ParseParameterModifierSettings(payload)
            SyncLock _parameterModifierLock
                _parameterModifierSettings = settings
            End SyncLock

            Dim progress As IProgress(Of Object) = New Progress(Of Object)(
                Sub(snapshot)
                    SendToWeb("parammodifier:progress", snapshot)
                End Sub)

            Try
                Dim result = ParameterModifierService.Run(app, settings, progress)
                SyncLock _parameterModifierLock
                    _parameterModifierLastResult = result
                End SyncLock

                SendToWeb("parammodifier:done", New With {
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
                SendToWeb("parammodifier:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleParameterModifierExportResults(app As UIApplication, payload As Object)
            Try
                Dim settings As ParameterModifierService.Settings = Nothing
                Dim result As ParameterModifierService.RunResult = Nothing

                SyncLock _parameterModifierLock
                    settings = _parameterModifierSettings
                    result = _parameterModifierLastResult
                End SyncLock

                If result Is Nothing Then
                    SendToWeb("parammodifier:artifacts-exported", New With {
                        .ok = False,
                        .message = "최근 실행 결과가 없습니다."
                    })
                    Return
                End If

                ParameterModifierService.ExportArtifacts(app, settings, result)

                SendToWeb("parammodifier:artifacts-exported", New With {
                    .ok = True,
                    .message = "결과 엑셀과 로그를 저장했습니다.",
                    .outputFolder = result.OutputFolder,
                    .resultWorkbookPath = result.ResultWorkbookPath,
                    .logTextPath = result.LogTextPath
                })
            Catch ex As Exception
                SendToWeb("parammodifier:artifacts-exported", New With {
                    .ok = False,
                    .message = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleParameterModifierOpenFolder(app As UIApplication, payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim pathText = Convert.ToString(GetProp(pd, "path"))
            If String.IsNullOrWhiteSpace(pathText) Then
                SyncLock _parameterModifierLock
                    If _parameterModifierLastResult IsNot Nothing Then pathText = _parameterModifierLastResult.OutputFolder
                    If String.IsNullOrWhiteSpace(pathText) AndAlso _parameterModifierSettings IsNot Nothing Then pathText = _parameterModifierSettings.OutputFolder
                End SyncLock
            End If

            If String.IsNullOrWhiteSpace(pathText) Then
                SendToWeb("parammodifier:error", New With {.message = "열 폴더 경로가 없습니다."})
                Return
            End If

            Dim targetPath = pathText
            If File.Exists(targetPath) Then targetPath = Path.GetDirectoryName(targetPath)
            If String.IsNullOrWhiteSpace(targetPath) OrElse Not Directory.Exists(targetPath) Then
                SendToWeb("parammodifier:error", New With {.message = "폴더 경로를 찾을 수 없습니다."})
                Return
            End If

            Dim psi As New ProcessStartInfo("explorer.exe", """" & targetPath & """")
            psi.UseShellExecute = True
            Process.Start(psi)
            SendToWeb("parammodifier:folder-opened", New With {.ok = True, .path = targetPath})
        End Sub

        Private Shared Function ParseParameterModifierSettings(payload As Object) As ParameterModifierService.Settings
            Dim pd = ParsePayloadDict(payload)
            Dim settings As New ParameterModifierService.Settings()
            settings.UseActiveDocument = SafeBoolObj(GetProp(pd, "useActiveDocument"), True)
            settings.RvtPaths = ParseStringList(pd, "rvtPaths")
            settings.OutputFolder = NormalizeDeliveryCleanerText(GetProp(pd, "outputFolder"))
            settings.CloseAllWorksetsOnOpen = SafeBoolObj(GetProp(pd, "closeAllWorksetsOnOpen"), True)
            settings.SynchronizeAfterProcessing = If(settings.UseActiveDocument, SafeBoolObj(GetProp(pd, "synchronizeAfterProcessing"), False), True)
            settings.SyncComment = NormalizeDeliveryCleanerText(GetProp(pd, "syncComment"))
            settings.FilterProfile = ParseDeliveryCleanerFilterProfile(GetProp(pd, "filterProfile"))
            settings.ElementParameterUpdate = ParseDeliveryCleanerElementUpdate(GetProp(pd, "elementParameterUpdate"))
            Return settings
        End Function

        Private Shared Function BuildParameterModifierStatePayload(app As UIApplication) As Object
            Dim settings As ParameterModifierService.Settings = Nothing
            Dim result As ParameterModifierService.RunResult = Nothing

            SyncLock _parameterModifierLock
                settings = _parameterModifierSettings
                result = _parameterModifierLastResult
            End SyncLock

            If settings Is Nothing Then
                settings = New ParameterModifierService.Settings() With {
                    .UseActiveDocument = True,
                    .CloseAllWorksetsOnOpen = True,
                    .SynchronizeAfterProcessing = False,
                    .ElementParameterUpdate = New ElementParameterUpdateSettings()
                }
            End If

            Return New With {
                .settings = New With {
                    .useActiveDocument = settings.UseActiveDocument,
                    .rvtPaths = If(settings.RvtPaths, New List(Of String)()),
                    .outputFolder = settings.OutputFolder,
                    .closeAllWorksetsOnOpen = settings.CloseAllWorksetsOnOpen,
                    .synchronizeAfterProcessing = settings.SynchronizeAfterProcessing,
                    .syncComment = settings.SyncComment,
                    .filterProfile = SerializeFilterProfile(settings.FilterProfile),
                    .elementParameterUpdate = SerializeElementParameterUpdateSettings(settings.ElementParameterUpdate)
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

    End Class

End Namespace
