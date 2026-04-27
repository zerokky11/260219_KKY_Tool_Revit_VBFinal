Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Exports
Imports KKY_Tool_Revit.Services

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private _linkPathLastRows As List(Of RevitLinkPathRow) = Nothing
        Private _linkPathLastWorkbookPath As String = ""

        Private Sub HandleLinkPathPickRvts()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.Title = "링크 추출 대상 RVT 선택"
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return
                SendToWebAfterDialog("linkpath:rvts-picked", New With {.paths = dlg.FileNames})
            End Using
        End Sub

        Private Sub HandleLinkPathPickExcel()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                dlg.Multiselect = False
                dlg.Title = "링크 수정용 엑셀 선택"
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return
                SendToWebAfterDialog("linkpath:excel-picked", New With {.path = dlg.FileName})
            End Using
        End Sub

        Private Sub HandleLinkPathExtract(app As UIApplication, payload As Object)
            Dim rvtPaths As List(Of String) = ExtractStringList(payload, "rvtPaths")
            If rvtPaths.Count = 0 Then
                SendToWeb("linkpath:error", New With {
                    .message = "추출할 RVT 파일이 없습니다.",
                    .detail = "RVT 목록을 먼저 추가해 주세요."
                })
                Return
            End If

            Try
                Dim rows As List(Of RevitLinkPathRow) = RevitLinkPathBatchService.Extract(app.Application, rvtPaths, AddressOf ReportLinkPathProgress)
                _linkPathLastRows = rows
                SendLinkPathRows("extract", rows, rvtPaths.Count, _linkPathLastWorkbookPath)
            Catch ex As Exception
                SendToWeb("linkpath:error", New With {
                    .message = "링크 경로 추출 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleLinkPathExport(payload As Object)
            If _linkPathLastRows Is Nothing OrElse _linkPathLastRows.Count = 0 Then
                SendToWeb("linkpath:error", New With {
                    .message = "내보낼 링크 데이터가 없습니다.",
                    .detail = "먼저 링크를 추출해 주세요."
                })
                Return
            End If

            Dim autoFit As Boolean = ParseExcelMode(payload)
            Dim fastExport As Boolean = Not autoFit
            Dim exportLocale As String = ParseExcelLocale(payload)

            Try
                ExcelProgressReporter.Reset("linkpath:progress")
                Dim savedPath As String = RevitLinkPathExport.Export(_linkPathLastRows, fastExport, autoFit, "linkpath:progress", exportLocale)
                If String.IsNullOrWhiteSpace(savedPath) Then
                    SendToWeb("linkpath:exported", New With {
                        .ok = False,
                        .message = "엑셀 저장이 취소되었습니다."
                    })
                    Return
                End If

                _linkPathLastWorkbookPath = savedPath
                SendToWeb("linkpath:exported", New With {
                    .ok = True,
                    .path = savedPath
                })
            Catch ex As Exception
                ExcelProgressReporter.Report("linkpath:progress", "ERROR", ex.Message, 0, 0, Nothing, True)
                SendToWeb("linkpath:error", New With {
                    .message = "엑셀 내보내기 실패",
                    .detail = ex.Message
                })
                SendToWeb("linkpath:exported", New With {
                    .ok = False,
                    .message = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleLinkPathImport(app As UIApplication, payload As Object)
            Dim excelPath As String = NormalizeLinkPathText(TryCast(GetProp(payload, "path"), String))
            If String.IsNullOrWhiteSpace(excelPath) Then
                SendToWeb("linkpath:error", New With {
                    .message = "불러올 엑셀 경로가 없습니다.",
                    .detail = "엑셀 파일을 먼저 선택해 주세요."
                })
                Return
            End If

            Try
                Dim rows As List(Of RevitLinkPathRow) = RevitLinkPathBatchService.ImportWorkbook(excelPath)
                _linkPathLastRows = rows
                _linkPathLastWorkbookPath = excelPath
                SendLinkPathRows("import", rows, DistinctHostCount(rows), excelPath)
            Catch ex As Exception
                SendToWeb("linkpath:error", New With {
                    .message = "엑셀 불러오기 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub HandleLinkPathApply(app As UIApplication, payload As Object)
            If _linkPathLastRows Is Nothing OrElse _linkPathLastRows.Count = 0 Then
                SendToWeb("linkpath:error", New With {
                    .message = "적용할 링크 데이터가 없습니다.",
                    .detail = "링크 추출 또는 엑셀 불러오기를 먼저 실행해 주세요."
                })
                Return
            End If

            If _linkPathLastRows.All(Function(x) x Is Nothing OrElse
                                                 (String.IsNullOrWhiteSpace(x.TargetLinkPath) AndAlso
                                                  String.IsNullOrWhiteSpace(x.ReferenceElementId))) Then
                SendToWeb("linkpath:error", New With {
                    .message = "적용할 링크 변경/삭제 데이터가 없습니다.",
                    .detail = "Reload From은 TargetLinkPath를 입력하고, 삭제는 기존 링크 행의 TargetLinkPath를 비워 주세요."
                })
                Return
            End If

            Try
                Dim applyOptions As New RevitLinkPathApplyOptions With {
                    .NewLinkPlacement = ParseLinkPathNewLinkPlacement(payload),
                    .HostPathHints = ExtractStringList(payload, "rvtPaths")
                }
                Dim appliedRows As List(Of RevitLinkPathRow) = RevitLinkPathBatchService.Apply(app.Application, _linkPathLastRows, AddressOf ReportLinkPathProgress, applyOptions)
                _linkPathLastRows = appliedRows
                SendLinkPathRows("apply", appliedRows, DistinctHostCount(appliedRows), _linkPathLastWorkbookPath)
                SendToWeb("linkpath:applied", New With {
                    .ok = True,
                    .summary = BuildLinkPathSummaryPayload(appliedRows, DistinctHostCount(appliedRows))
                })
            Catch ex As Exception
                SendToWeb("linkpath:error", New With {
                    .message = "링크 경로 적용 실패",
                    .detail = ex.Message
                })
            End Try
        End Sub

        Private Sub SendLinkPathRows(source As String,
                                     rows As IList(Of RevitLinkPathRow),
                                     hostCount As Integer,
                                     workbookPath As String)
            Dim payloadRows As List(Of Dictionary(Of String, Object)) =
                If(rows, New List(Of RevitLinkPathRow)()).
                    Where(Function(x) x IsNot Nothing).
                    Select(AddressOf LinkPathRowToDict).
                    ToList()

            SendToWeb("linkpath:rows", New With {
                .source = source,
                .schema = RevitLinkPathExport.Schema,
                .rows = payloadRows,
                .workbookPath = workbookPath,
                .summary = BuildLinkPathSummaryPayload(rows, hostCount)
            })
        End Sub

        Private Function LinkPathRowToDict(row As RevitLinkPathRow) As Dictionary(Of String, Object)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            d("HostFileName") = row.HostFileName
            d("HostFilePath") = row.HostFilePath
            d("ReferenceElementId") = row.ReferenceElementId
            d("LinkName") = row.LinkName
            d("LinkFileName") = row.LinkFileName
            d("TypeWorksetNames") = row.TypeWorksetNames
            d("InstanceWorksetNames") = row.InstanceWorksetNames
            d("ApplyTypeWorksetNames") = row.ApplyTypeWorksetNames
            d("ApplyInstanceWorksetNames") = row.ApplyInstanceWorksetNames
            d("CurrentLinkPath") = row.CurrentLinkPath
            d("StoredLinkPath") = row.StoredLinkPath
            d("CurrentPathType") = row.CurrentPathType
            d("TargetLinkPath") = row.TargetLinkPath
            d("TargetPathType") = row.TargetPathType
            d("ApplyStatus") = row.ApplyStatus
            d("ApplyMessage") = row.ApplyMessage
            Return d
        End Function

        Private Function BuildLinkPathSummaryPayload(rows As IList(Of RevitLinkPathRow), hostCount As Integer) As Object
            Dim safeRows As List(Of RevitLinkPathRow) = If(rows, New List(Of RevitLinkPathRow)()).Where(Function(x) x IsNot Nothing).ToList()
            Return New With {
                .hostCount = Math.Max(0, hostCount),
                .rowCount = safeRows.Count,
                .targetCount = safeRows.Where(Function(x) Not String.IsNullOrWhiteSpace(x.TargetLinkPath)).Count(),
                .deleteCandidateCount = safeRows.Where(Function(x) String.IsNullOrWhiteSpace(x.TargetLinkPath) AndAlso Not String.IsNullOrWhiteSpace(x.ReferenceElementId)).Count(),
                .deletedCount = safeRows.Where(Function(x) StringEquals(x.ApplyStatus, "changed") AndAlso If(x.ApplyMessage, "").IndexOf("삭제", StringComparison.OrdinalIgnoreCase) >= 0).Count(),
                .changedCount = safeRows.Where(Function(x) StringEquals(x.ApplyStatus, "changed")).Count(),
                .skipCount = safeRows.Where(Function(x) StringEquals(x.ApplyStatus, "skip") OrElse StringEquals(x.ApplyStatus, "info")).Count(),
                .errorCount = safeRows.Where(Function(x) StringEquals(x.ApplyStatus, "error")).Count()
            }
        End Function

        Private Function DistinctHostCount(rows As IEnumerable(Of RevitLinkPathRow)) As Integer
            If rows Is Nothing Then Return 0
            Return rows.
                Where(Function(x) x IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(x.HostFilePath)).
                Select(Function(x) x.HostFilePath).
                Distinct(StringComparer.OrdinalIgnoreCase).
                Count()
        End Function

        Private Sub ReportLinkPathProgress(percent As Integer, message As String)
            Try
                SendToWeb("linkpath:progress", New With {
                    .percent = Math.Max(0, Math.Min(100, percent)),
                    .message = If(message, ""),
                    .detail = If(message, ""),
                    .stage = ResolveLinkPathProgressStage(percent, message)
                })
            Catch
            End Try
        End Sub

        Private Function ResolveLinkPathProgressStage(percent As Integer, detail As String) As String
            Dim message As String = If(detail, String.Empty).Trim()

            If message.IndexOf("확인", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "RVT 확인 중"
            End If
            If message.IndexOf("읽는 중", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 데이터 읽는 중"
            End If
            If message.IndexOf("스캔", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 스캔 중"
            End If
            If message.IndexOf("열기", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "RVT 열기 중"
            End If
            If message.IndexOf("신규", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "신규 링크 생성 중"
            End If
            If message.IndexOf("삭제", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 삭제 중"
            End If
            If message.IndexOf("재로드", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("Reload", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 재로드 중"
            End If
            If message.IndexOf("동기화", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "센트럴 동기화 중"
            End If
            If message.IndexOf("저장", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "RVT 저장 중"
            End If
            If message.IndexOf("추출", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 추출 중"
            End If
            If message.IndexOf("적용", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "링크 적용 중"
            End If
            If message.IndexOf("완료", StringComparison.OrdinalIgnoreCase) >= 0 OrElse percent >= 100 Then
                Return "링크 작업 완료"
            End If

            Return "링크 작업 진행 중"
        End Function

        Private Function NormalizeLinkPathText(value As String) As String
            Dim s As String = NormalizeWrappedQuotesText(value)
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Try
                If s.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
                    Dim u As New Uri(s)
                    If u IsNot Nothing AndAlso u.IsFile Then s = u.LocalPath
                End If
            Catch
            End Try
            Return s.Replace("/"c, "\"c).Trim()
        End Function

        Private Function ParseLinkPathNewLinkPlacement(payload As Object) As String
            Dim raw As String = Convert.ToString(GetProp(payload, "newLinkPlacement"))
            Dim text As String = If(raw, "").Trim().Replace("-"c, "_"c).Replace(" "c, "_"c).ToLowerInvariant()

            Select Case text
                Case "center", "centered", "center_to_center", "centertocenter"
                    Return "Centered"
                Case "shared", "shared_coordinates", "sharedcoordinates"
                    Return "Shared"
                Case "site", "project_base_point", "projectbasepoint"
                    Return "Site"
                Case Else
                    Return "Origin"
            End Select
        End Function

        Private Function StringEquals(left As String, right As String) As Boolean
            Return String.Equals(If(left, "").Trim(), If(right, "").Trim(), StringComparison.OrdinalIgnoreCase)
        End Function

    End Class

End Namespace
