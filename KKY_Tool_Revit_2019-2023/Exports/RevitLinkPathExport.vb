Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Exports

    Public Module RevitLinkPathExport

        Public ReadOnly Schema As String() = {
            "HostFileName",
            "HostFilePath",
            "ReferenceElementId",
            "LinkName",
            "LinkFileName",
            "TypeWorksetNames",
            "InstanceWorksetNames",
            "ApplyTypeWorksetNames",
            "ApplyInstanceWorksetNames",
            "CurrentLinkPath",
            "StoredLinkPath",
            "CurrentPathType",
            "TargetLinkPath",
            "TargetPathType",
            "ApplyStatus",
            "ApplyMessage"
        }

        Private ReadOnly HiddenHeaders As HashSet(Of String) =
            New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "ReferenceElementId",
                "StoredLinkPath",
                "TargetPathType",
                "ApplyStatus",
                "ApplyMessage"
            }

        Public Function Export(rows As IEnumerable(Of RevitLinkPathRow),
                               Optional fastExport As Boolean = True,
                               Optional autoFit As Boolean = False,
                               Optional progressChannel As String = Nothing,
                               Optional exportLocale As String = "ko") As String
            If rows Is Nothing Then Return String.Empty

            Dim table As DataTable = ToDataTable(rows)
            ExcelCore.EnsureMessageRow(table, "추출된 Revit 링크가 없습니다.")
            If Not ValidateSchema(table) Then
                Throw New InvalidOperationException("스키마 검증 실패: RevitLinkPath")
            End If

            Dim defaultName As String = $"RevitLinkPath_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                dlg.FileName = defaultName
                dlg.AddExtension = True
                dlg.DefaultExt = "xlsx"
                dlg.OverwritePrompt = True
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return String.Empty

                Dim filePath As String = dlg.FileName
                If Not String.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                    filePath = Path.ChangeExtension(filePath, ".xlsx")
                End If

                Dim doAutoFit As Boolean = (Not fastExport) AndAlso autoFit
                ReportProgress(progressChannel, "EXCEL_INIT", "엑셀 저장 준비 중...", 0, Math.Max(1, table.Rows.Count), 0.0R, True)
                ExcelCore.SaveXlsx(filePath, "RevitLinkPath", table, doAutoFit, progressKey:=progressChannel, exportKind:="linkpath", exportLocale:=exportLocale)
                HideInternalColumns(filePath, "RevitLinkPath")
                ReportProgress(progressChannel, "EXCEL_SAVE", "엑셀 파일 저장 중...", Math.Max(1, table.Rows.Count), Math.Max(1, table.Rows.Count), 0.95R, True)
                If doAutoFit Then
                    ReportProgress(progressChannel, "AUTOFIT", "열 너비 자동 조정 중...", 1, 1, 1.0R, True)
                End If
                ReportProgress(progressChannel, "DONE", "엑셀 내보내기 완료", 1, 1, 1.0R, True)
                Return filePath
            End Using
        End Function

        Public Function ToDataTable(rows As IEnumerable(Of RevitLinkPathRow)) As DataTable
            Dim dt As New DataTable("RevitLinkPath")
            For Each h In Schema
                dt.Columns.Add(h)
            Next

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim dr = dt.NewRow()
                dr("HostFileName") = SafeStr(row.HostFileName)
                dr("HostFilePath") = SafeStr(row.HostFilePath)
                dr("ReferenceElementId") = SafeStr(row.ReferenceElementId)
                dr("LinkName") = SafeStr(row.LinkName)
                dr("LinkFileName") = SafeStr(row.LinkFileName)
                dr("TypeWorksetNames") = SafeStr(row.TypeWorksetNames)
                dr("InstanceWorksetNames") = SafeStr(row.InstanceWorksetNames)
                dr("ApplyTypeWorksetNames") = SafeStr(row.ApplyTypeWorksetNames)
                dr("ApplyInstanceWorksetNames") = SafeStr(row.ApplyInstanceWorksetNames)
                dr("CurrentLinkPath") = SafeStr(row.CurrentLinkPath)
                dr("StoredLinkPath") = SafeStr(row.StoredLinkPath)
                dr("CurrentPathType") = SafeStr(row.CurrentPathType)
                dr("TargetLinkPath") = SafeStr(row.TargetLinkPath)
                dr("TargetPathType") = SafeStr(row.TargetPathType)
                dr("ApplyStatus") = SafeStr(row.ApplyStatus)
                dr("ApplyMessage") = SafeStr(row.ApplyMessage)
                dt.Rows.Add(dr)
            Next

            Return dt
        End Function

        Private Function ValidateSchema(table As DataTable) As Boolean
            If table Is Nothing Then Return False
            If table.Columns.Count <> Schema.Length Then Return False
            For i As Integer = 0 To Schema.Length - 1
                If Not String.Equals(table.Columns(i).ColumnName, Schema(i), StringComparison.Ordinal) Then
                    Return False
                End If
            Next
            Return True
        End Function

        Private Sub ReportProgress(channel As String,
                                   phase As String,
                                   message As String,
                                   current As Integer,
                                   total As Integer,
                                   Optional percentOverride As Double? = Nothing,
                                   Optional force As Boolean = False)
            If String.IsNullOrWhiteSpace(channel) Then Return
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(channel, phase, message, current, total, percentOverride, force)
        End Sub

        Private Sub HideInternalColumns(filePath As String, sheetName As String)
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return

            Dim workbook As IWorkbook = Nothing
            Using readStream As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                workbook = New XSSFWorkbook(readStream)
            End Using
            If workbook Is Nothing Then Return

            Try
                Dim sheet As ISheet = workbook.GetSheet(sheetName)
                If sheet Is Nothing AndAlso workbook.NumberOfSheets > 0 Then
                    sheet = workbook.GetSheetAt(0)
                End If
                If sheet Is Nothing Then Return

                Dim headerRow As IRow = sheet.GetRow(sheet.FirstRowNum)
                If headerRow Is Nothing Then Return

                Dim changed As Boolean = False
                For col As Integer = 0 To CInt(headerRow.LastCellNum) - 1
                    Dim headerText As String = SafeStr(headerRow.GetCell(col)?.ToString()).Trim()
                    If HiddenHeaders.Contains(headerText) Then
                        sheet.SetColumnHidden(col, True)
                        changed = True
                    End If
                Next

                If Not changed Then Return

                Using writeStream As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    workbook.Write(writeStream)
                End Using
            Finally
                Try
                    workbook.Close()
                Catch
                End Try
            End Try
        End Sub

        Private Function SafeStr(value As String) As String
            Return If(value, "")
        End Function

    End Module

End Namespace
