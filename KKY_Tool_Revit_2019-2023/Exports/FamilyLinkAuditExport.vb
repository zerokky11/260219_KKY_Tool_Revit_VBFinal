Option Explicit On
Option Strict On

Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports KKY_Tool_Revit.Infrastructure
Imports KKY_Tool_Revit.Services

Namespace Exports

    Public Module FamilyLinkAuditExport

        Public ReadOnly Schema As String() = {
            "FileName",
            "HostFamilyName",
            "HostFamilyCategory",
            "NestedFamilyName",
            "NestedTypeName",
            "NestedInstanceId",
            "NestedPath",
            "NestingLevel",
            "NestedCategory",
            "NestedParamName",
            "TargetParamName",
            "ExpectedGuid",
            "FoundScope",
            "NestedParamGuid",
            "NestedParamDataType",
            "AssocHostParamName",
            "HostParamIsShared",
            "Issue",
            "Notes"
        }

        Public Function Export(rows As IEnumerable(Of FamilyLinkAuditRow),
                               Optional fastExport As Boolean = True,
                               Optional autoFit As Boolean = False,
                               Optional progressChannel As String = Nothing) As String
            If rows Is Nothing Then Return String.Empty
            Dim table As DataTable = ToDataTable(rows)
            ' Global.KKY_Tool_Revit.Infrastructure.ResultTableFilter.KeepOnlyIssues("familylink", table)
            ExcelCore.EnsureMessageRow(table, "오류가 없습니다.")
            If Not ValidateSchema(table) Then
                Throw New InvalidOperationException("스키마 검증 실패: 컬럼 순서/헤더가 규격과 다릅니다.")
            End If

            Dim defaultName As String = $"FamilyLinkAudit_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            Using dlg As New SaveFileDialog()
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                dlg.FileName = defaultName
                dlg.AddExtension = True
                dlg.DefaultExt = "xlsx"
                dlg.OverwritePrompt = True
                dlg.RestoreDirectory = True

                If dlg.ShowDialog() <> DialogResult.OK Then Return String.Empty

                Dim filePath As String = dlg.FileName
                Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()
                If Not String.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                    filePath = Path.ChangeExtension(filePath, ".xlsx")
                End If

                Dim doAutoFit As Boolean = (Not fastExport) AndAlso autoFit
                Dim excelMode As String = If(fastExport, "fast", "normal")
                ReportProgress(progressChannel, "EXCEL_INIT", "엑셀 저장 준비 중...", 0, Math.Max(1, table.Rows.Count), 0.0R, True)
                ExcelCore.SaveXlsx(filePath, "FamilyLinkAudit", table, doAutoFit, progressKey:=progressChannel, exportKind:="familylink")
                ReportProgress(progressChannel, "EXCEL_SAVE", "엑셀 파일 저장 중...", Math.Max(1, table.Rows.Count), Math.Max(1, table.Rows.Count), 0.95R, True)
                ExcelExportStyleRegistry.ApplyStylesForKey("familylink", filePath, autoFit:=doAutoFit, excelMode:=excelMode)
                If doAutoFit Then
                    ReportProgress(progressChannel, "AUTOFIT", "열 너비 자동 조정 중...", 1, 1, 1.0R, True)
                End If
                ReportProgress(progressChannel, "DONE", "내보내기 완료", 1, 1, 1.0R, True)
                Return filePath
            End Using
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

        Public Function ToDataTable(rows As IEnumerable(Of FamilyLinkAuditRow)) As DataTable
            Return BuildTable(rows)
        End Function

        Private Function BuildTable(rows As IEnumerable(Of FamilyLinkAuditRow)) As DataTable
            Dim dt As New DataTable("FamilyLinkAudit")
            For Each h In Schema
                dt.Columns.Add(h)
            Next

            For Each r In rows
                Dim dr = dt.NewRow()
                dr("FileName") = SafeStr(r.FileName)
                dr("HostFamilyName") = SafeStr(r.HostFamilyName)
                dr("HostFamilyCategory") = SafeStr(r.HostFamilyCategory)
                dr("NestedFamilyName") = SafeStr(r.NestedFamilyName)
                dr("NestedTypeName") = SafeStr(r.NestedTypeName)
                dr("NestedInstanceId") = SafeStr(r.NestedInstanceId)
                dr("NestedPath") = SafeStr(r.NestedPath)
                dr("NestingLevel") = SafeStr(r.NestingLevel)
                dr("NestedCategory") = SafeStr(r.NestedCategory)
                dr("NestedParamName") = SafeStr(r.NestedParamName)
                dr("TargetParamName") = SafeStr(r.TargetParamName)
                dr("ExpectedGuid") = SafeStr(r.ExpectedGuid)
                dr("FoundScope") = SafeStr(r.FoundScope)
                dr("NestedParamGuid") = SafeStr(r.NestedParamGuid)
                dr("NestedParamDataType") = SafeStr(r.NestedParamDataType)
                dr("AssocHostParamName") = SafeStr(r.AssocHostParamName)
                dr("HostParamIsShared") = SafeStr(r.HostParamIsShared)
                dr("Issue") = SafeStr(r.Issue)
                dr("Notes") = SafeStr(r.Notes)
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

        Private Function SafeStr(s As String) As String
            Return If(s, "")
        End Function

    End Module

End Namespace
