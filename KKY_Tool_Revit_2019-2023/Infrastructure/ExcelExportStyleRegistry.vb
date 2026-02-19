Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Public Module ExcelExportStyleRegistry

        Public Delegate Function RowStatusResolver(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus

        Private ReadOnly _resolvers As New Dictionary(Of String, RowStatusResolver)(StringComparer.OrdinalIgnoreCase)

        Sub New()
            Register("connector", AddressOf ResolveConnector)
            Register("guid", AddressOf ResolveResultLike)
            Register("paramprop", AddressOf ResolveParamProp)
            Register("points", AddressOf ResolveResultLike)
            Register("pms", AddressOf ResolveResultLike)
            Register("familylink", AddressOf ResolveIssueLike)
            Register("sharedparambatch", AddressOf ResolveSharedParamBatch)

            Register("connector diagnostics", AddressOf ResolveConnector)
            Register("familylinkaudit", AddressOf ResolveIssueLike)
            Register("family link audit", AddressOf ResolveIssueLike)
            Register("pms vs segment size검토", AddressOf ResolveResultLike)
            Register("pipe segment class검토", AddressOf ResolveResultLike)
            Register("routing class검토", AddressOf ResolveResultLike)
        End Sub

        Public Sub Register(key As String, resolver As RowStatusResolver)
            If String.IsNullOrWhiteSpace(key) OrElse resolver Is Nothing Then Return
            _resolvers(key.Trim()) = resolver
        End Sub

        Public Function Resolve(sheetNameOrKey As String, row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            If row Is Nothing OrElse table Is Nothing Then Return ExcelStyleHelper.RowStatus.None

            Dim key = NormalizeKey(sheetNameOrKey)
            Dim fn As RowStatusResolver = Nothing

            If Not String.IsNullOrWhiteSpace(key) AndAlso _resolvers.TryGetValue(key, fn) Then
                Return SafeResolve(fn, row, table)
            End If

            If Not String.IsNullOrWhiteSpace(sheetNameOrKey) AndAlso _resolvers.TryGetValue(sheetNameOrKey.Trim(), fn) Then
                Return SafeResolve(fn, row, table)
            End If

            Return ResolveGeneric(row, table)
        End Function

        Public Function FilterIssueRows(styleKey As String, table As DataTable) As DataTable
            If table Is Nothing Then Return Nothing

            Dim filtered As DataTable = table.Clone()
            For Each row As DataRow In table.Rows
                Dim status As ExcelStyleHelper.RowStatus = Resolve(styleKey, row, table)
                If status = ExcelStyleHelper.RowStatus.Warning OrElse status = ExcelStyleHelper.RowStatus.[Error] Then
                    filtered.ImportRow(row)
                End If
            Next
            Return filtered
        End Function

        Public Sub ApplyStylesForKey(styleKey As String,
                                     xlsxPath As String,
                                     Optional sheetName As String = Nothing,
                                     Optional autoFit As Boolean = True,
                                     Optional excelMode As String = "normal")
            If String.IsNullOrWhiteSpace(xlsxPath) Then Return
            If Not File.Exists(xlsxPath) Then Return

            Using fs As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim wb As IWorkbook = New XSSFWorkbook(fs)

                Dim targetSheets As New List(Of ISheet)()
                If Not String.IsNullOrWhiteSpace(sheetName) Then
                    Dim namedSheet = wb.GetSheet(sheetName)
                    If namedSheet IsNot Nothing Then targetSheets.Add(namedSheet)
                End If

                If targetSheets.Count = 0 Then
                    For i As Integer = 0 To wb.NumberOfSheets - 1
                        Dim sh = wb.GetSheetAt(i)
                        If sh IsNot Nothing Then targetSheets.Add(sh)
                    Next
                End If

                For Each sh In targetSheets
                    ApplyStandardSheetStyle(wb, sh)
                    ApplyBorders(wb, sh)
                    ApplyKeySpecificFormats(styleKey, sh)
                Next

                fs.Close()
                Using outFs As New FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(outFs)
                End Using
            End Using
        End Sub

        Private Sub ApplyStandardSheetStyle(wb As IWorkbook, sh As ISheet)
            If wb Is Nothing OrElse sh Is Nothing Then Return

            Dim header = sh.GetRow(0)
            If header Is Nothing Then Return

            Dim headerStyle As ICellStyle = ExcelStyleHelper.GetHeaderStyle(wb)
            If headerStyle Is Nothing Then Return

            Dim lastCol As Integer = CInt(header.LastCellNum) - 1
            If lastCol < 0 Then Return

            For c As Integer = 0 To lastCol
                Dim cell = header.GetCell(c)
                If cell Is Nothing Then
                    cell = header.CreateCell(c)
                    cell.SetCellValue("")
                End If
                cell.CellStyle = headerStyle
            Next
        End Sub

        Private Sub ApplyBorders(wb As IWorkbook, sh As ISheet)
            If wb Is Nothing OrElse sh Is Nothing Then Return
            Dim lastRow As Integer = sh.LastRowNum
            If lastRow < 0 Then Return

            For r As Integer = 0 To lastRow
                Dim row = sh.GetRow(r)
                If row Is Nothing Then Continue For

                Dim lastCol As Integer = CInt(row.LastCellNum) - 1
                If lastCol < 0 Then Continue For

                For c As Integer = 0 To lastCol
                    Dim cell = row.GetCell(c)
                    If cell Is Nothing Then
                        cell = row.CreateCell(c)
                        cell.SetCellValue("")
                    End If

                    Dim src = cell.CellStyle
                    Dim dst = wb.CreateCellStyle()
                    If src IsNot Nothing Then
                        dst.CloneStyleFrom(src)
                    End If
                    dst.BorderBottom = BorderStyle.Thin
                    dst.BorderTop = BorderStyle.Thin
                    dst.BorderLeft = BorderStyle.Thin
                    dst.BorderRight = BorderStyle.Thin
                    cell.CellStyle = dst
                Next
            Next
        End Sub

        Private Sub ApplyKeySpecificFormats(styleKey As String, sh As ISheet)
            If sh Is Nothing Then Return
            Dim key As String = NormalizeKey(styleKey)
            If String.IsNullOrWhiteSpace(key) Then Return

            ' key별 세부 포맷 확장 포인트 (현재는 중복 처리 방지를 위해 최소 동작 유지)
            If key = "pms" Then
                Return
            End If
        End Sub

        Private Function SafeResolve(fn As RowStatusResolver, row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Try
                Return fn(row, table)
            Catch
                Return ExcelStyleHelper.RowStatus.None
            End Try
        End Function

        Private Function NormalizeKey(sheetNameOrKey As String) As String
            If String.IsNullOrWhiteSpace(sheetNameOrKey) Then Return ""
            Dim s = sheetNameOrKey.Trim().ToLowerInvariant()
            If s.Contains("connector") Then Return "connector"
            If s.Contains("guid") Then Return "guid"
            If s.Contains("param") Then Return "paramprop"
            If s.Contains("point") Then Return "points"
            If s.Contains("sharedparambatch") Then Return "sharedparambatch"
            If s.Contains("familylink") OrElse s.Contains("family link") Then Return "familylink"
            If s.Contains("pms") OrElse s.Contains("segment") Then Return "pms"
            Return s
        End Function

        Private Function ResolveConnector(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim statusText = GetFirstExistingText(row, table, "Status", "Result", "검토결과")
            If IsOkLike(statusText) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(statusText) Then Return ExcelStyleHelper.RowStatus.[Error]
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveIssueLike(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim issue = GetFirstExistingText(row, table, "Issue", "Result", "Status")
            If IsOkLike(issue) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(issue) Then Return ExcelStyleHelper.RowStatus.[Error]
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveResultLike(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim result = GetFirstExistingText(row, table, "Result", "Res", "Status", "성공여부", "검토결과", "Class검토결과", "Class검토")
            If IsOkLike(result) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(result) Then Return ExcelStyleHelper.RowStatus.[Error]
            If String.IsNullOrWhiteSpace(result) Then Return ExcelStyleHelper.RowStatus.Info
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveParamProp(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim detail = GetFirstExistingText(row, table, "Detail", "Result", "Status", "메시지")
            If String.IsNullOrWhiteSpace(detail) Then Return ExcelStyleHelper.RowStatus.None
            If IsOkLike(detail) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(detail) Then Return ExcelStyleHelper.RowStatus.[Error]
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveSharedParamBatch(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim status = GetFirstExistingText(row, table, "성공여부", "Level", "Status", "Result", "메시지")
            If IsOkLike(status) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(status) Then Return ExcelStyleHelper.RowStatus.[Error]
            If String.IsNullOrWhiteSpace(status) Then Return ExcelStyleHelper.RowStatus.None
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveGeneric(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim txt = GetFirstExistingText(row, table, "Status", "Result", "Issue", "Error", "ErrorMessage", "Notes")
            If String.IsNullOrWhiteSpace(txt) Then Return ExcelStyleHelper.RowStatus.None
            If IsOkLike(txt) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(txt) Then Return ExcelStyleHelper.RowStatus.[Error]
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function GetFirstExistingText(row As DataRow, table As DataTable, ParamArray colNames() As String) As String
            For Each c In colNames
                Dim t = GetColText(row, table, c)
                If Not String.IsNullOrWhiteSpace(t) Then Return t
            Next
            Return ""
        End Function

        Private Function GetColText(row As DataRow, table As DataTable, colName As String) As String
            For Each col As DataColumn In table.Columns
                If String.Equals(col.ColumnName, colName, StringComparison.OrdinalIgnoreCase) Then
                    Dim v = row(col)
                    If v Is Nothing OrElse v Is DBNull.Value Then Return ""
                    Return v.ToString().Trim()
                End If
            Next
            Return ""
        End Function

        Private Function IsOkLike(s As String) As Boolean
            If String.IsNullOrWhiteSpace(s) Then Return False
            Dim t = s.Trim().ToLowerInvariant()
            If t = "ok" OrElse t = "pass" OrElse t = "success" Then Return True
            If t.StartsWith("ok(") OrElse t.StartsWith("ok[") OrElse t.StartsWith("ok_") Then Return True
            If t.Contains("오류 없음") OrElse t.Contains("정상") OrElse t.Contains("이상 없음") Then Return True
            Return False
        End Function

        Private Function LooksError(s As String) As Boolean
            If String.IsNullOrWhiteSpace(s) Then Return False
            Dim t = s.Trim().ToLowerInvariant()
            If t.Contains("mismatch") OrElse t.Contains("불일치") Then Return True
            If t.Contains("error") OrElse t.Contains("fail") Then Return True
            If t.Contains("실패") OrElse t.Contains("오류") Then Return True
            Return False
        End Function

    End Module

End Namespace
