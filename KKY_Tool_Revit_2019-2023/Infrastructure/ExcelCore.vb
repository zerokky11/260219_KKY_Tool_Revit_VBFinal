Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Public Module ExcelCore

        Public Function PickAndSaveXlsx(title As String,
                                        table As DataTable,
                                        defaultFileName As String,
                                        Optional autoFit As Boolean = False,
                                        Optional progressKey As String = Nothing,
                                        Optional exportKind As String = Nothing) As String

            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, title)
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsx(path, If(String.IsNullOrWhiteSpace(table.TableName), title, table.TableName), table, autoFit, sheetKey:=title, progressKey:=progressKey, exportKind:=exportKind)
            Return path
        End Function

        Public Function PickAndSaveXlsxMulti(sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                             defaultFileName As String,
                                             Optional autoFit As Boolean = False,
                                             Optional progressKey As String = Nothing) As String

            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentException("Sheets is empty.", NameOf(sheets))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, "엑셀 저장")
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsxMulti(path, sheets, autoFit, progressKey)
            Return path
        End Function

        Public Sub SaveXlsx(filePath As String,
                            sheetName As String,
                            table As DataTable,
                            Optional autoFit As Boolean = False,
                            Optional sheetKey As String = Nothing,
                            Optional progressKey As String = Nothing,
                            Optional exportKind As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            EnsureDir(filePath)
            If ShouldEnsureNoDataRow(sheetName, sheetKey, exportKind) Then
                EnsureNoDataRow(table)
            End If

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim safeSheet = NormalizeSheetName(If(sheetName, "Sheet1"))
                Dim sheet = wb.CreateSheet(safeSheet)

                WriteTableToSheet(wb, sheet, safeSheet, table, sheetKey, autoFit, progressKey, exportKind)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub

        Public Sub SaveXlsxMulti(filePath As String,
                                 sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                 Optional autoFit As Boolean = False,
                                 Optional progressKey As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentException("Sheets is empty.", NameOf(sheets))

            EnsureDir(filePath)

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim usedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For i As Integer = 0 To sheets.Count - 1
                    Dim name = If(sheets(i).Key, $"Sheet{i + 1}")
                    Dim table = sheets(i).Value
                    If table Is Nothing Then Continue For

                    Dim safe = MakeUniqueSheetName(NormalizeSheetName(name), usedNames)
                    usedNames.Add(safe)

                    If ShouldEnsureNoDataRow(safe, name, Nothing) Then
                        EnsureNoDataRow(table)
                    End If
                    Dim sheet = wb.CreateSheet(safe)
                    WriteTableToSheet(wb, sheet, safe, table, sheetKey:=name, autoFit:=autoFit, progressKey:=progressKey, exportKind:=Nothing)
                Next

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub


        Public Sub SaveStyledSimple(filePath As String,
                                    title As String,
                                    table As DataTable,
                                    groupHeader As String,
                                    Optional autoFit As Boolean = False,
                                    Optional progressKey As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            EnsureDir(filePath)

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim baseName As String = If(String.IsNullOrWhiteSpace(title),
                                            If(String.IsNullOrWhiteSpace(table.TableName), "Sheet1", table.TableName),
                                            title)

                Dim safeSheet = NormalizeSheetName(baseName)
                Dim sh = wb.CreateSheet(safeSheet)

                ' 1) 값 쓰기 (AutoFit은 스타일 적용 후 마지막에)
                WriteTableToSheet(wb, sh, safeSheet, table, sheetKey:=title, autoFit:=False, progressKey:=progressKey, exportKind:=Nothing)

                ' 2) 그룹 밴딩 (DuplicateExport: "Group" 컬럼)
                If Not String.IsNullOrWhiteSpace(groupHeader) Then
                    TryApplyGroupBanding(wb, sh, table, groupHeader)
                End If

                ' 3) 기본 시트 스타일 (Freeze/Filter/Border/AutoFit)
                ApplyStandardSheetStyle(wb, sh, headerRowIndex:=0, autoFilter:=True, freezeTopRow:=True, borderAll:=True, autoFit:=autoFit)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub

        Private Sub TryApplyGroupBanding(wb As IWorkbook, sh As ISheet, table As DataTable, groupHeader As String)
            If wb Is Nothing OrElse sh Is Nothing OrElse table Is Nothing Then Return
            If String.IsNullOrWhiteSpace(groupHeader) Then Return
            If table.Columns.Count = 0 OrElse table.Rows.Count = 0 Then Return

            Dim groupCol As Integer = -1
            For c As Integer = 0 To table.Columns.Count - 1
                If String.Equals(table.Columns(c).ColumnName, groupHeader, StringComparison.OrdinalIgnoreCase) Then
                    groupCol = c
                    Exit For
                End If
            Next
            If groupCol < 0 Then Return

            Dim cache As New Dictionary(Of Integer, ICellStyle)()
            Dim lastKey As String = Nothing
            Dim band As Boolean = False

            For r As Integer = 0 To table.Rows.Count - 1
                Dim v = table.Rows(r)(groupCol)
                Dim keyText As String = If(v Is Nothing OrElse v Is DBNull.Value, "", v.ToString())

                If lastKey Is Nothing Then
                    lastKey = keyText
                ElseIf Not String.Equals(lastKey, keyText, StringComparison.Ordinal) Then
                    band = Not band
                    lastKey = keyText
                End If

                If Not band Then Continue For

                Dim row = sh.GetRow(r + 1) ' header=0, data starts at 1
                If row Is Nothing Then Continue For

                Dim lastCol As Integer = table.Columns.Count - 1
                For c As Integer = 0 To lastCol
                    Dim cell = row.GetCell(c)
                    If cell Is Nothing Then Continue For

                    Dim baseStyle = cell.CellStyle
                    Dim styleKey As Integer = If(baseStyle Is Nothing, -1, CInt(baseStyle.Index))

                    Dim st As ICellStyle = Nothing
                    If Not cache.TryGetValue(styleKey, st) Then
                        st = wb.CreateCellStyle()
                        If baseStyle IsNot Nothing Then st.CloneStyleFrom(baseStyle)
                        st.FillForegroundColor = IndexedColors.Grey25Percent.Index
                        st.FillPattern = FillPattern.SolidForeground
                        cache(styleKey) = st
                    End If

                    cell.CellStyle = st
                Next
            Next
        End Sub


        Private ReadOnly ReviewExportKeys As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "guid",
            "familylink",
            "paramprop",
            "pms",
            "sharedparambatch",
            "connector"
        }

        Private Function ShouldEnsureNoDataRow(sheetName As String,
                                               sheetKey As String,
                                               exportKind As String) As Boolean
            Dim key As String = NormalizeExportPolicyKey(exportKind, sheetKey, sheetName)
            If String.IsNullOrWhiteSpace(key) Then Return False

            If key.Equals("points", StringComparison.OrdinalIgnoreCase) OrElse
               key.Equals("export", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Return ReviewExportKeys.Contains(key)
        End Function

        Private Function NormalizeExportPolicyKey(exportKind As String,
                                                  sheetKey As String,
                                                  sheetName As String) As String
            Dim raw As String = ""
            If Not String.IsNullOrWhiteSpace(exportKind) Then
                raw = exportKind
            ElseIf Not String.IsNullOrWhiteSpace(sheetKey) Then
                raw = sheetKey
            Else
                raw = If(sheetName, "")
            End If

            If String.IsNullOrWhiteSpace(raw) Then Return ""
            Dim s As String = raw.Trim().ToLowerInvariant()

            If s.Contains("point") Then Return "points"
            If s = "export" Then Return "export"
            If s.Contains("guid") Then Return "guid"
            If s.Contains("familylink") OrElse s.Contains("family link") Then Return "familylink"
            If s.Contains("param") Then Return "paramprop"
            If s.Contains("pms") OrElse s.Contains("segment") Then Return "pms"
            If s.Contains("sharedparambatch") Then Return "sharedparambatch"
            If s.Contains("connector") Then Return "connector"

            Return s
        End Function

        Public Sub EnsureNoDataRow(table As DataTable,
                                   Optional message As String = "오류가 없습니다.")
            If table Is Nothing Then Return

            If table.Columns.Count = 0 Then
                table.Columns.Add("Message", GetType(String))
            End If

            If table.Rows.Count > 0 Then Return

            Dim finalMessage As String = If(String.IsNullOrWhiteSpace(message), "오류가 없습니다.", message)
            Dim row = table.NewRow()
            row(0) = finalMessage
            table.Rows.Add(row)
        End Sub

        Public Sub EnsureMessageRow(table As DataTable,
                                    Optional message As String = "오류가 없습니다.")
            EnsureNoDataRow(table, message)
        End Sub

        Public Sub SaveCsv(filePath As String, table As DataTable)
            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))
            EnsureDir(filePath)

            Using sw As New StreamWriter(filePath, False, New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                ' header
                For c As Integer = 0 To table.Columns.Count - 1
                    If c > 0 Then sw.Write(",")
                    sw.Write(EscapeCsv(table.Columns(c).ColumnName))
                Next
                sw.WriteLine()

                ' rows
                For r As Integer = 0 To table.Rows.Count - 1
                    Dim dr = table.Rows(r)
                    For c As Integer = 0 To table.Columns.Count - 1
                        If c > 0 Then sw.Write(",")
                        Dim v = dr(c)
                        Dim s = If(v Is Nothing OrElse v Is DBNull.Value, "", v.ToString())
                        sw.Write(EscapeCsv(s))
                    Next
                    sw.WriteLine()
                Next
            End Using
        End Sub

        ' ---------------- internal ----------------

        Private Sub WriteTableToSheet(wb As IWorkbook,
                                      sheet As ISheet,
                                      sheetName As String,
                                      table As DataTable,
                                      sheetKey As String,
                                      autoFit As Boolean,
                                      progressKey As String,
                                      exportKind As String)

            Dim colCount As Integer = table.Columns.Count
            If colCount = 0 Then Return

            ' header
            Dim headerRow = sheet.CreateRow(0)
            Dim isConnector As Boolean = String.Equals(exportKind, "connector", StringComparison.OrdinalIgnoreCase)
            Dim headerStyle As ICellStyle = If(isConnector, ExcelStyleHelper.GetHeaderStyleNoWrap(wb), ExcelStyleHelper.GetHeaderStyle(wb))
            If isConnector Then
                headerRow.Height = -1
            End If

            For c As Integer = 0 To colCount - 1
                Dim cell = headerRow.CreateCell(c)
                cell.SetCellValue(table.Columns(c).ColumnName)
                cell.CellStyle = headerStyle
            Next

            sheet.CreateFreezePane(0, 1)

            Dim total As Integer = table.Rows.Count
            For r As Integer = 0 To total - 1
                Dim dr = table.Rows(r)
                Dim row = sheet.CreateRow(r + 1)
                If isConnector Then
                    row.Height = -1
                End If

                For c As Integer = 0 To colCount - 1
                    WriteCell(row, c, dr(c))
                Next

                ' ---- 핵심: 저장하면서 행 상태를 판정해서 배경색 적용 ----
                Dim status = ExcelExportStyleRegistry.Resolve(If(sheetKey, sheetName), dr, table)
                If status <> ExcelStyleHelper.RowStatus.None Then
                    Dim style = If(isConnector, ExcelStyleHelper.GetRowStyleNoWrap(wb, status), ExcelStyleHelper.GetRowStyle(wb, status))
                    ExcelStyleHelper.ApplyStyleToRow(row, colCount, style)
                End If

                If (r Mod 200) = 0 Then
                    TryReportProgress(progressKey, r, total, sheetName)
                End If
            Next

            If isConnector AndAlso colCount > 0 Then
                Try
                    Dim lastRowIndex As Integer = Math.Max(0, total)
                    Dim range As New CellRangeAddress(0, lastRowIndex, 0, colCount - 1)
                    sheet.SetAutoFilter(range)
                Catch
                End Try
            End If

            If autoFit Then
                TryTrackAllColumnsForAutoSizing(sheet)
                For c As Integer = 0 To colCount - 1
                    Try
                        sheet.AutoSizeColumn(c)
                    Catch
                    End Try
                Next
            End If
        End Sub

        Private Sub WriteCell(row As IRow, colIndex As Integer, value As Object)
            Dim cell = row.CreateCell(colIndex)

            If value Is Nothing OrElse value Is DBNull.Value Then
                cell.SetCellValue("")
                Return
            End If

            If TypeOf value Is Boolean Then
                cell.SetCellValue(CBool(value))
                Return
            End If

            If TypeOf value Is DateTime Then
                cell.SetCellValue(DirectCast(value, DateTime))
                Return
            End If

            If TypeOf value Is Byte OrElse TypeOf value Is Short OrElse TypeOf value Is Integer OrElse
               TypeOf value Is Long OrElse TypeOf value Is Single OrElse TypeOf value Is Double OrElse
               TypeOf value Is Decimal Then

                Dim d As Double
                If Double.TryParse(value.ToString(), d) Then
                    cell.SetCellValue(d)
                Else
                    cell.SetCellValue(value.ToString())
                End If
                Return
            End If

            cell.SetCellValue(value.ToString())
        End Sub

        Private Function PickSavePath(filter As String, defaultFileName As String, title As String) As String
            Using dlg As New SaveFileDialog()
                dlg.Filter = filter
                dlg.Title = If(String.IsNullOrWhiteSpace(title), "저장", title)
                dlg.FileName = If(String.IsNullOrWhiteSpace(defaultFileName), "export.xlsx", defaultFileName)
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then Return ""
                Return dlg.FileName
            End Using
        End Function

        Private Sub EnsureDir(filePath As String)
            Dim dir = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
        End Sub

        Private Function NormalizeSheetName(name As String) As String
            Dim s = If(name, "Sheet1").Trim()
            If s.Length = 0 Then s = "Sheet1"

            ' Excel 금지 문자: : \ / ? * [ ]
            Dim bad = New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
            For Each ch In bad
                s = s.Replace(ch, "_"c)
            Next

            If s.Length > 31 Then s = s.Substring(0, 31)
            Return s
        End Function

        Private Function MakeUniqueSheetName(baseName As String, used As HashSet(Of String)) As String
            Dim s = baseName
            Dim i As Integer = 1
            While used.Contains(s)
                Dim suffix = $"({i})"
                Dim cut = Math.Min(31 - suffix.Length, baseName.Length)
                s = baseName.Substring(0, cut) & suffix
                i += 1
            End While
            Return s
        End Function

        Private Function EscapeCsv(s As String) As String
            If s Is Nothing Then Return ""
            Dim needs = s.Contains(","c) OrElse s.Contains(""""c) OrElse s.Contains(vbCr) OrElse s.Contains(vbLf)
            Dim t = s.Replace("""", """""")
            If needs Then Return $"""{t}"""
            Return t
        End Function

        ' ---------------- 추가: SegmentPms/Connector 등에서 호출되는 스타일 헬퍼 ----------------

        Public Sub ApplyStandardSheetStyle(wb As IWorkbook,
                                           sh As ISheet,
                                           Optional headerRowIndex As Integer = 0,
                                           Optional autoFilter As Boolean = True,
                                           Optional freezeTopRow As Boolean = True,
                                           Optional borderAll As Boolean = False,
                                           Optional autoFit As Boolean = False)

            If wb Is Nothing OrElse sh Is Nothing Then Return

            If freezeTopRow Then
                Try
                    sh.CreateFreezePane(0, headerRowIndex + 1)
                Catch
                End Try
            End If

            If autoFilter Then
                Try
                    Dim headerRow = sh.GetRow(headerRowIndex)
                    If headerRow IsNot Nothing Then
                        Dim lastCol As Integer = CInt(headerRow.LastCellNum) - 1
                        If lastCol >= 0 Then
                            Dim lastRow As Integer = Math.Max(headerRowIndex, sh.LastRowNum)
                            Dim range As New CellRangeAddress(headerRowIndex, lastRow, 0, lastCol)
                            TrySetAutoFilter(sh, range)
                        End If
                    End If
                Catch
                End Try
            End If

            If borderAll Then
                TryApplyThinBorderToUsedRange(wb, sh)
            End If

            If autoFit Then
                TryTrackAllColumnsForAutoSizing(sh)
                Dim headerRow = sh.GetRow(headerRowIndex)
                Dim lastCol As Integer = If(headerRow Is Nothing, -1, CInt(headerRow.LastCellNum) - 1)
                If lastCol >= 0 Then
                    For c As Integer = 0 To lastCol
                        Try
                            sh.AutoSizeColumn(c)
                        Catch
                        End Try
                    Next
                End If
            End If
        End Sub

        Public Sub ApplyNumberFormatByHeader(wb As IWorkbook,
                                            sh As ISheet,
                                            headerRowIndex As Integer,
                                            headers As IEnumerable(Of String),
                                            numberFormat As String)

            If wb Is Nothing OrElse sh Is Nothing OrElse headers Is Nothing Then Return

            Dim headerRow = sh.GetRow(headerRowIndex)
            If headerRow Is Nothing Then Return

            Dim headerSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each h In headers
                If Not String.IsNullOrWhiteSpace(h) Then headerSet.Add(h.Trim())
            Next
            If headerSet.Count = 0 Then Return

            Dim targetCols As New List(Of Integer)()
            Dim fmt As New DataFormatter()
            Dim eval = wb.GetCreationHelper().CreateFormulaEvaluator()

            Dim lastCol As Integer = CInt(headerRow.LastCellNum) - 1
            For c As Integer = 0 To lastCol
                Dim cell = headerRow.GetCell(c)
                Dim text As String = ""
                Try
                    text = fmt.FormatCellValue(cell, eval).Trim()
                Catch
                End Try

                If headerSet.Contains(text) Then
                    targetCols.Add(c)
                End If
            Next
            If targetCols.Count = 0 Then Return

            Dim fmtIdx As Short = wb.CreateDataFormat().GetFormat(If(String.IsNullOrWhiteSpace(numberFormat), "0.###############", numberFormat))

            Dim styleCache As New Dictionary(Of Integer, ICellStyle)()

            For r As Integer = headerRowIndex + 1 To sh.LastRowNum
                Dim row = sh.GetRow(r)
                If row Is Nothing Then Continue For

                For Each c In targetCols
                    Dim cell = row.GetCell(c)
                    If cell Is Nothing Then Continue For

                    If cell.CellType = CellType.Numeric OrElse cell.CellType = CellType.Formula Then
                        Dim baseStyle = cell.CellStyle
                        Dim key As Integer = If(baseStyle Is Nothing, -1, CInt(baseStyle.Index))

                        Dim newStyle As ICellStyle = Nothing
                        If Not styleCache.TryGetValue(key, newStyle) Then
                            newStyle = wb.CreateCellStyle()
                            If baseStyle IsNot Nothing Then newStyle.CloneStyleFrom(baseStyle)
                            newStyle.DataFormat = fmtIdx
                            styleCache(key) = newStyle
                        End If

                        cell.CellStyle = newStyle
                    End If
                Next
            Next
        End Sub

        Public Sub ApplyResultFillByHeader(wb As IWorkbook, sh As ISheet, headerRowIndex As Integer)
            If wb Is Nothing OrElse sh Is Nothing Then Return

            Dim headerRow = sh.GetRow(headerRowIndex)
            If headerRow Is Nothing Then Return

            Dim fmt As New DataFormatter()
            Dim eval = wb.GetCreationHelper().CreateFormulaEvaluator()

            Dim resultCol As Integer = -1
            Dim lastCol As Integer = CInt(headerRow.LastCellNum) - 1

            For c As Integer = 0 To lastCol
                Dim h As String = ""
                Try
                    h = fmt.FormatCellValue(headerRow.GetCell(c), eval).Trim()
                Catch
                End Try

                Dim norm = NormalizeHeader(h)
                If norm = "result" OrElse norm = "status" Then
                    resultCol = c
                    Exit For
                End If
            Next

            If resultCol < 0 Then Return

            Dim warnCache As New Dictionary(Of Integer, ICellStyle)()
            Dim errCache As New Dictionary(Of Integer, ICellStyle)()

            For r As Integer = headerRowIndex + 1 To sh.LastRowNum
                Dim row = sh.GetRow(r)
                If row Is Nothing Then Continue For

                Dim cell = row.GetCell(resultCol)
                If cell Is Nothing Then Continue For

                Dim text As String = ""
                Try
                    text = fmt.FormatCellValue(cell, eval)
                Catch
                End Try

                Dim cls As Integer = ClassifyResult(text) ' 0=ok, 1=warn, 2=err
                If cls = 0 Then Continue For

                Dim baseStyle = cell.CellStyle
                Dim key As Integer = If(baseStyle Is Nothing, -1, CInt(baseStyle.Index))

                If cls = 2 Then
                    Dim st As ICellStyle = Nothing
                    If Not errCache.TryGetValue(key, st) Then
                        st = wb.CreateCellStyle()
                        If baseStyle IsNot Nothing Then st.CloneStyleFrom(baseStyle)
                        st.FillForegroundColor = IndexedColors.Rose.Index
                        st.FillPattern = FillPattern.SolidForeground
                        errCache(key) = st
                    End If
                    cell.CellStyle = st
                ElseIf cls = 1 Then
                    Dim st As ICellStyle = Nothing
                    If Not warnCache.TryGetValue(key, st) Then
                        st = wb.CreateCellStyle()
                        If baseStyle IsNot Nothing Then st.CloneStyleFrom(baseStyle)
                        st.FillForegroundColor = IndexedColors.LightYellow.Index
                        st.FillPattern = FillPattern.SolidForeground
                        warnCache(key) = st
                    End If
                    cell.CellStyle = st
                End If
            Next
        End Sub

        Public Sub TryAutoFitWithExcel(xlsxPath As String)
            If String.IsNullOrWhiteSpace(xlsxPath) Then Return
            If Not File.Exists(xlsxPath) Then Return

            Dim excelApp As Object = Nothing
            Dim wbs As Object = Nothing
            Dim wb As Object = Nothing

            Try
                excelApp = CreateObject("Excel.Application")
                If excelApp Is Nothing Then Return

                excelApp.DisplayAlerts = False
                excelApp.Visible = False

                wbs = excelApp.Workbooks
                wb = wbs.Open(xlsxPath)

                Dim sheets As Object = Nothing
                Try
                    sheets = wb.Worksheets
                    For Each ws As Object In sheets
                        Try
                            ws.Cells.EntireColumn.AutoFit()
                        Catch
                        Finally
                            ReleaseCom(ws)
                        End Try
                    Next
                Catch
                Finally
                    ReleaseCom(sheets)
                End Try

                Try
                    wb.Save()
                Catch
                End Try

            Catch
                ' ignore (Excel 미설치/권한/보안정책 등)
            Finally
                Try
                    If wb IsNot Nothing Then wb.Close(SaveChanges:=True)
                Catch
                End Try
                Try
                    If excelApp IsNot Nothing Then excelApp.Quit()
                Catch
                End Try

                ReleaseCom(wb)
                ReleaseCom(wbs)
                ReleaseCom(excelApp)
            End Try
        End Sub

        Private Sub ReleaseCom(o As Object)
            Try
                If o Is Nothing Then Return
                If Marshal.IsComObject(o) Then
                    Marshal.FinalReleaseComObject(o)
                End If
            Catch
            End Try
        End Sub

        Private Sub TryTrackAllColumnsForAutoSizing(sheet As ISheet)
            If sheet Is Nothing Then Return
            Try
                Dim mi = sheet.GetType().GetMethod("TrackAllColumnsForAutoSizing", Type.EmptyTypes)
                If mi IsNot Nothing Then mi.Invoke(sheet, Nothing)
            Catch
            End Try
        End Sub

        Private Sub TrySetAutoFilter(sheet As ISheet, range As CellRangeAddress)
            If sheet Is Nothing OrElse range Is Nothing Then Return
            Try
                Dim mi = sheet.GetType().GetMethod("SetAutoFilter", New Type() {GetType(CellRangeAddress)})
                If mi IsNot Nothing Then mi.Invoke(sheet, New Object() {range})
            Catch
            End Try
        End Sub

        Private Sub TryApplyThinBorderToUsedRange(wb As IWorkbook, sh As ISheet)
            If wb Is Nothing OrElse sh Is Nothing Then Return

            Dim maxCol As Integer = GetMaxUsedColumnIndex(sh)
            If maxCol < 0 Then Return

            Dim cache As New Dictionary(Of Integer, ICellStyle)()

            For r As Integer = 0 To sh.LastRowNum
                Dim row = sh.GetRow(r)
                If row Is Nothing Then Continue For

                For c As Integer = 0 To maxCol
                    Dim cell = row.GetCell(c)
                    If cell Is Nothing Then
                        cell = row.CreateCell(c)
                        cell.SetCellValue("")
                    End If

                    Dim baseStyle = cell.CellStyle
                    Dim key As Integer = If(baseStyle Is Nothing, -1, CInt(baseStyle.Index))

                    Dim st As ICellStyle = Nothing
                    If Not cache.TryGetValue(key, st) Then
                        st = wb.CreateCellStyle()
                        If baseStyle IsNot Nothing Then st.CloneStyleFrom(baseStyle)
                        st.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
                        st.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
                        st.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
                        st.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
                        cache(key) = st
                    End If

                    cell.CellStyle = st
                Next
            Next
        End Sub

        Private Function GetMaxUsedColumnIndex(sh As ISheet) As Integer
            Dim maxCol As Integer = -1
            For r As Integer = 0 To sh.LastRowNum
                Dim row = sh.GetRow(r)
                If row Is Nothing Then Continue For

                Dim lastCellNum As Integer = CInt(row.LastCellNum)
                If lastCellNum <= 0 Then Continue For

                Dim lastIdx As Integer = lastCellNum - 1
                If lastIdx > maxCol Then maxCol = lastIdx
            Next
            Return maxCol
        End Function

        Private Function NormalizeHeader(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Return s.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "")
        End Function

        Private Function ClassifyResult(s As String) As Integer
            If String.IsNullOrWhiteSpace(s) Then Return 0
            Dim t = s.Trim().ToLowerInvariant()

            If t = "ok" OrElse t = "pass" OrElse t = "success" Then Return 0
            If t.Contains("오류 없음") OrElse t.Contains("정상") OrElse t.Contains("이상 없음") Then Return 0

            If t.Contains("error") OrElse t.Contains("fail") OrElse t.Contains("mismatch") Then Return 2
            If t.Contains("실패") OrElse t.Contains("오류") OrElse t.Contains("불일치") Then Return 2

            If t.Contains("na") OrElse t.Contains("n/a") OrElse t.Contains("missing") OrElse t.Contains("없음") Then Return 1
            Return 1
        End Function

        ' progressKey는 UiBridge에서 "hub:multi-progress" 같은 채널로 쓰는 구조가 있어서
        ' 여기서는 있으면 최대한 조용히 반영(리플렉션)하고, 없어도 기능은 정상 동작하게 처리
        Private Sub TryReportProgress(progressKey As String, current As Integer, total As Integer, sheetName As String)
            If String.IsNullOrWhiteSpace(progressKey) Then Return
            Try
                Dim t = Type.GetType("KKY_Tool_Revit.UI.Hub.ExcelProgressReporter, " & GetType(ExcelCore).Assembly.FullName, throwOnError:=False)
                If t Is Nothing Then Return

                Dim mi = t.GetMethod("Report", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If mi Is Nothing Then Return

                Dim percent As Double = 0
                If total > 0 Then percent = (CDbl(current) / CDbl(total)) * 100.0R
                mi.Invoke(Nothing, New Object() {progressKey, percent, $"Exporting {sheetName}...", $"{current}/{total}"})
            Catch
            End Try
        End Sub

    End Module

End Namespace
