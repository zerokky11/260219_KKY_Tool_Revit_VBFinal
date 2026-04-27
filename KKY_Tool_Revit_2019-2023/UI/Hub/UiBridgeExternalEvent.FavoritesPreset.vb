Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Text
Imports System.Web.Script.Serialization
Imports WinForms = System.Windows.Forms

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Const FavoritePresetFilter As String = "KKY 즐겨찾기 프리셋 (*.kkyfav.json)|*.kkyfav.json|JSON 파일 (*.json)|*.json"

        Private Sub HandleFavoritesPresetSave(payload As Object)
            Dim pd = ParsePayloadDict(payload)
            Dim jsonText As String = Convert.ToString(GetProp(pd, "json"))
            If String.IsNullOrWhiteSpace(jsonText) Then
                SendToWeb("favorites:preset-error", New With {.message = "저장할 프리셋 데이터가 없습니다."})
                Return
            End If

            Try
                Dim serializer As New JavaScriptSerializer()
                serializer.MaxJsonLength = Integer.MaxValue
                serializer.Deserialize(Of Object)(jsonText)
            Catch ex As Exception
                SendToWeb("favorites:preset-error", New With {.message = "프리셋 데이터 형식이 올바르지 않습니다: " & ex.Message})
                Return
            End Try

            Dim defaultName As String = NormalizeFavoritePresetFileName(Convert.ToString(GetProp(pd, "defaultName")))

            Try
                Using dlg As New WinForms.SaveFileDialog()
                    dlg.Filter = FavoritePresetFilter
                    dlg.Title = "즐겨찾기 프리셋 저장"
                    dlg.RestoreDirectory = True
                    dlg.AddExtension = True
                    dlg.DefaultExt = "json"
                    dlg.OverwritePrompt = True
                    dlg.ValidateNames = True
                    dlg.FileName = defaultName

                    Dim docDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    If Not String.IsNullOrWhiteSpace(docDir) AndAlso Directory.Exists(docDir) Then
                        dlg.InitialDirectory = docDir
                    End If

                    If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                    Dim savePath As String = dlg.FileName
                    If String.IsNullOrWhiteSpace(savePath) Then Return

                    File.WriteAllText(savePath, jsonText, New UTF8Encoding(False))

                    SendToWebAfterDialog("favorites:preset-saved", New With {
                        .path = savePath,
                        .fileName = Path.GetFileName(savePath)
                    })
                End Using
            Catch ex As Exception
                SendToWebAfterDialog("favorites:preset-error", New With {.message = "프리셋 저장 실패: " & ex.Message})
            End Try
        End Sub

        Private Sub HandleFavoritesPresetLoad()
            Try
                Using dlg As New WinForms.OpenFileDialog()
                    dlg.Filter = FavoritePresetFilter
                    dlg.Title = "즐겨찾기 프리셋 불러오기"
                    dlg.Multiselect = False
                    dlg.RestoreDirectory = True
                    dlg.CheckFileExists = True
                    dlg.CheckPathExists = True

                    Dim docDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    If Not String.IsNullOrWhiteSpace(docDir) AndAlso Directory.Exists(docDir) Then
                        dlg.InitialDirectory = docDir
                    End If

                    If dlg.ShowDialog() <> WinForms.DialogResult.OK Then Return

                    Dim filePath As String = dlg.FileName
                    If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                        SendToWebAfterDialog("favorites:preset-error", New With {.message = "선택한 프리셋 파일을 찾을 수 없습니다."})
                        Return
                    End If

                    Dim jsonText As String = File.ReadAllText(filePath, Encoding.UTF8)
                    If String.IsNullOrWhiteSpace(jsonText) Then
                        SendToWebAfterDialog("favorites:preset-error", New With {.message = "선택한 프리셋 파일이 비어 있습니다."})
                        Return
                    End If

                    Dim serializer As New JavaScriptSerializer()
                    serializer.MaxJsonLength = Integer.MaxValue
                    serializer.Deserialize(Of Object)(jsonText)

                    SendToWebAfterDialog("favorites:preset-loaded", New With {
                        .path = filePath,
                        .fileName = Path.GetFileName(filePath),
                        .json = jsonText
                    })
                End Using
            Catch ex As Exception
                SendToWebAfterDialog("favorites:preset-error", New With {.message = "프리셋 불러오기 실패: " & ex.Message})
            End Try
        End Sub

        Private Shared Function NormalizeFavoritePresetFileName(value As String) As String
            Dim name As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(name) Then
                name = $"favorites-preset-{Date.Now:yyyyMMdd-HHmm}.kkyfav.json"
            End If

            For Each ch As Char In Path.GetInvalidFileNameChars()
                name = name.Replace(ch, "_"c)
            Next

            If Not name.EndsWith(".json", StringComparison.OrdinalIgnoreCase) Then
                name &= ".kkyfav.json"
            End If

            Return name
        End Function

    End Class

End Namespace
