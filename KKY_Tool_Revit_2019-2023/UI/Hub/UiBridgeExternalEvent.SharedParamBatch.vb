Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Web.Script.Serialization
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared _lastSharedParamBatchResult As SharedParamBatchService.RunResult
        Private Shared _lastSharedParamBatchPayloadJson As String

        Private Sub HandleSharedParamBatchInit(app As UIApplication, payload As Object)
            Try
                Dim res = SharedParamBatchService.Init(app)
                SendToWeb("sharedparambatch:init", res)
            Catch ex As Exception
                SendToWeb("sharedparambatch:init", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleSharedParamBatchBrowseRvts(app As UIApplication, payload As Object)
            Try
                Dim res = SharedParamBatchService.BrowseRvts()
                SendToWeb("sharedparambatch:rvts-picked", res)
            Catch ex As Exception
                SendToWeb("sharedparambatch:rvts-picked", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleSharedParamBatchBrowseFolder(app As UIApplication, payload As Object)
            Try
                Dim res = SharedParamBatchService.BrowseRvtFolder()
                SendToWeb("sharedparambatch:rvts-picked", res)
            Catch ex As Exception
                SendToWeb("sharedparambatch:rvts-picked", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Sub HandleSharedParamBatchRun(app As UIApplication, payload As Object)
            Try
                Dim serializer As New JavaScriptSerializer()
                Dim payloadJson As String = serializer.Serialize(payload)
                _lastSharedParamBatchPayloadJson = payloadJson

                Dim progress As IProgress(Of Object) = New Progress(Of Object)(Sub(p)
                                                                                  SendToWeb("sharedparambatch:progress", p)
                                                                              End Sub)

                Dim res As SharedParamBatchService.RunResult = TryCast(SharedParamBatchService.Run(app, payloadJson, progress), SharedParamBatchService.RunResult)
                If res IsNot Nothing Then
                    res.Logs = FilterSharedParamBatchIssueLogs(res.Logs)
                End If
                _lastSharedParamBatchResult = res

                If res Is Nothing Then
                    SendToWeb("sharedparambatch:done", New With {.ok = False, .message = "결과를 확인할 수 없습니다."})
                    Return
                End If

                Dim responsePayload As New With {
                    .ok = res.Ok,
                    .message = res.Message,
                    .summary = If(res.Summary, New SharedParamBatchService.RunSummary()),
                    .logs = If(res.Logs, New List(Of SharedParamBatchService.LogEntry)()).Select(Function(l) New With {
                        .level = l.Level,
                        .file = l.File,
                        .msg = l.Message
                    }).ToList(),
                    .logTextPath = res.LogTextPath
                }

                SendToWeb("sharedparambatch:done", responsePayload)

                If Not res.Ok Then
                    SendToWeb("revit:error", New With {.message = "Shared Parameter Batch 실패: " & res.Message})
                End If
            Catch ex As Exception
                SendToWeb("sharedparambatch:done", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "Shared Parameter Batch 실패: " & ex.Message})
            End Try
        End Sub

        Private Shared Function FilterSharedParamBatchIssueLogs(logs As List(Of SharedParamBatchService.LogEntry)) As List(Of SharedParamBatchService.LogEntry)
            Dim source As List(Of SharedParamBatchService.LogEntry) = If(logs, New List(Of SharedParamBatchService.LogEntry)())
            Dim dt As New Data.DataTable("SharedParamBatchLogs")
            dt.Columns.Add("__idx", GetType(Integer))
            dt.Columns.Add("성공여부")
            dt.Columns.Add("메시지")

            Dim index As Integer = 0
            For Each l In source
                Dim row = dt.NewRow()
                row("__idx") = index
                row("성공여부") = If(l Is Nothing, "", If(l.Level, ""))
                row("메시지") = If(l Is Nothing, "", If(l.Message, ""))
                dt.Rows.Add(row)
                index += 1
            Next

            Dim filtered As Data.DataTable = FilterIssueRowsCopy("sharedparambatch", dt)
            Dim result As New List(Of SharedParamBatchService.LogEntry)()
            If filtered Is Nothing Then Return result

            For Each row As Data.DataRow In filtered.Rows
                Dim sourceIdx As Integer = -1
                Try
                    sourceIdx = Convert.ToInt32(row("__idx"))
                Catch
                    sourceIdx = -1
                End Try
                If sourceIdx >= 0 AndAlso sourceIdx < source.Count Then
                    result.Add(source(sourceIdx))
                End If
            Next

            Return result
        End Function

        Private Sub HandleSharedParamBatchExportExcel(app As UIApplication, payload As Object)
            Try
                If _lastSharedParamBatchResult Is Nothing Then
                    SendToWeb("sharedparambatch:exported", New With {.ok = False, .message = "최근 실행 결과가 없습니다."})
                    Return
                End If

                Dim mode As String = Nothing
                Try
                    Dim prop = GetProp(payload, "excelMode")
                    If prop IsNot Nothing Then mode = Convert.ToString(prop)
                Catch
                End Try

                Dim logsPayload = _lastSharedParamBatchResult.Logs.Select(Function(l) New With {
                    .level = l.Level,
                    .file = l.File,
                    .msg = l.Message
                }).ToList()

                Dim serializer As New JavaScriptSerializer()
                Dim runPayload As Dictionary(Of String, Object) = Nothing
                If Not String.IsNullOrWhiteSpace(_lastSharedParamBatchPayloadJson) Then
                    Try
                        runPayload = serializer.Deserialize(Of Dictionary(Of String, Object))(_lastSharedParamBatchPayloadJson)
                    Catch
                        runPayload = Nothing
                    End Try
                End If

                Dim rvtPaths As Object = Nothing
                Dim parameters As Object = Nothing
                If runPayload IsNot Nothing Then
                    runPayload.TryGetValue("rvtPaths", rvtPaths)
                    runPayload.TryGetValue("parameters", parameters)
                End If

                Dim payloadJson As String = serializer.Serialize(New With {
                    .logs = logsPayload,
                    .excelMode = mode,
                    .rvtPaths = rvtPaths,
                    .parameters = parameters
                })

                Dim res = SharedParamBatchService.ExportExcel(payloadJson)
                SendToWeb("sharedparambatch:exported", res)
            Catch ex As Exception
                SendToWeb("sharedparambatch:exported", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

        Private Sub HandleSharedParamBatchOpenFolder(app As UIApplication, payload As Object)
            Dim inputPath As String = TryCast(GetProp(payload, "path"), String)
            If String.IsNullOrWhiteSpace(inputPath) Then
                SendToWeb("sharedparambatch:open-folder", New With {.ok = False, .message = "경로가 비어 있습니다."})
                Return
            End If

            Dim targetPath As String = inputPath
            Try
                If File.Exists(inputPath) Then
                    targetPath = Path.GetDirectoryName(inputPath)
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(targetPath) Then
                SendToWeb("sharedparambatch:open-folder", New With {.ok = False, .message = "폴더 경로를 확인할 수 없습니다."})
                Return
            End If

            Try
                Dim targetPathText As String = targetPath.ToString()
                Dim psi As New ProcessStartInfo("explorer.exe", """" & targetPathText & """")
                psi.UseShellExecute = True
                Process.Start(psi)
                SendToWeb("sharedparambatch:open-folder", New With {.ok = True, .path = targetPathText})
            Catch ex As Exception
                SendToWeb("sharedparambatch:open-folder", New With {.ok = False, .message = ex.Message})
            End Try
        End Sub

    End Class

End Namespace
