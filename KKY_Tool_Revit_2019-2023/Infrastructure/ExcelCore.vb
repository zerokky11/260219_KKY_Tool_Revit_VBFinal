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

        Private Const HiddenColumnFlag As String = "ExcelHidden"

        Public Function PickAndSaveXlsx(title As String,
                                        table As DataTable,
                                        defaultFileName As String,
                                        Optional autoFit As Boolean = False,
                                        Optional progressKey As String = Nothing,
                                        Optional exportKind As String = Nothing,
                                        Optional exportLocale As String = Nothing) As String

            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, title)
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsx(path, If(String.IsNullOrWhiteSpace(table.TableName), title, table.TableName), table, autoFit, sheetKey:=title, progressKey:=progressKey, exportKind:=exportKind, exportLocale:=exportLocale)
            Return path
        End Function

        Public Function PickAndSaveXlsxMulti(sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                             defaultFileName As String,
                                             Optional autoFit As Boolean = False,
                                             Optional progressKey As String = Nothing,
                                             Optional sheetKeyOverride As String = Nothing,
                                             Optional exportKind As String = Nothing,
                                             Optional exportLocale As String = Nothing) As String

            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentException("Sheets is empty.", NameOf(sheets))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, "엑셀 저장")
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsxMulti(path, sheets, autoFit, progressKey, sheetKeyOverride:=sheetKeyOverride, exportKind:=exportKind, exportLocale:=exportLocale)
            Return path
        End Function

        Public Function PickAndSaveStyledSimple(title As String,
                                                table As DataTable,
                                                defaultFileName As String,
                                                groupHeader As String,
                                                Optional autoFit As Boolean = False,
                                                Optional progressKey As String = Nothing,
                                                Optional exportKind As String = Nothing,
                                                Optional exportLocale As String = Nothing) As String

            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, title)
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveStyledSimple(path, title, table, groupHeader, autoFit, progressKey, exportKind, exportLocale)
            Return path
        End Function

        Public Sub SaveXlsx(filePath As String,
                            sheetName As String,
                            table As DataTable,
                            Optional autoFit As Boolean = False,
                            Optional sheetKey As String = Nothing,
                            Optional progressKey As String = Nothing,
                            Optional exportKind As String = Nothing,
                            Optional exportLocale As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            EnsureDir(filePath)
            If ShouldEnsureNoDataRow(sheetName, sheetKey, exportKind) Then
                EnsureNoDataRow(table)
            End If

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim safeSheet = NormalizeSheetName(If(sheetName, "Sheet1"))
                Dim sheet = wb.CreateSheet(safeSheet)

                Dim exportTable As DataTable = PrepareTableForExport(table, exportKind, exportLocale)
                WriteTableToSheet(wb, sheet, safeSheet, exportTable, sheetKey, autoFit, progressKey, exportKind)
                ApplyStandardSheetStyle(wb, sheet, headerRowIndex:=0, autoFilter:=True, freezeTopRow:=True, borderAll:=True, autoFit:=False)
                ReportExcelPhase(progressKey, "EXCEL_SAVE", "엑셀 파일 저장 중...", 0, 1, 0, True)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using

            If autoFit Then
                ReportExcelPhase(progressKey, "AUTOFIT", "열 너비 자동 조정 중...", 0, 1, 0, True)
                TryAutoFitWithExcel(filePath)
            End If
            ReportExcelPhase(progressKey, "DONE", "엑셀 저장 완료", 1, 1, 100, True)
        End Sub

        Public Sub SaveXlsxMulti(filePath As String,
                                 sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                 Optional autoFit As Boolean = False,
                                 Optional progressKey As String = Nothing,
                                 Optional sheetKeyOverride As String = Nothing,
                                 Optional exportKind As String = Nothing,
                                 Optional exportLocale As String = Nothing)

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

                    If ShouldEnsureNoDataRow(safe, name, exportKind) Then
                        EnsureNoDataRow(table)
                    End If
                    Dim sheet = wb.CreateSheet(safe)
                    Dim keyForStyle As String = If(String.IsNullOrWhiteSpace(sheetKeyOverride), name, sheetKeyOverride)
                    Dim exportTable As DataTable = PrepareTableForExport(table, exportKind, exportLocale)
                    WriteTableToSheet(wb, sheet, safe, exportTable, sheetKey:=keyForStyle, autoFit:=autoFit, progressKey:=progressKey, exportKind:=exportKind)
                    ApplyStandardSheetStyle(wb, sheet, headerRowIndex:=0, autoFilter:=True, freezeTopRow:=True, borderAll:=True, autoFit:=False)
                Next

                ReportExcelPhase(progressKey, "EXCEL_SAVE", "엑셀 파일 저장 중...", 0, 1, 0, True)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using

            If autoFit Then
                ReportExcelPhase(progressKey, "AUTOFIT", "열 너비 자동 조정 중...", 0, 1, 0, True)
                TryAutoFitWithExcel(filePath)
            End If
            ReportExcelPhase(progressKey, "DONE", "엑셀 저장 완료", 1, 1, 100, True)
        End Sub


        Public Sub SaveStyledSimple(filePath As String,
                                    title As String,
                                    table As DataTable,
                                    groupHeader As String,
                                    Optional autoFit As Boolean = False,
                                    Optional progressKey As String = Nothing,
                                    Optional exportKind As String = Nothing,
                                    Optional exportLocale As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            EnsureDir(filePath)

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim baseName As String = If(String.IsNullOrWhiteSpace(title),
                                            If(String.IsNullOrWhiteSpace(table.TableName), "Sheet1", table.TableName),
                                            title)

                Dim safeSheet = NormalizeSheetName(baseName)
                Dim sh = wb.CreateSheet(safeSheet)
                Dim exportKey As String = NormalizeExportPolicyKey(exportKind, table.TableName, table.TableName)
                Dim exportLocaleKey As String = NormalizeExcelLocale(exportLocale)
                Dim exportTable As DataTable = PrepareTableForExport(table, exportKind, exportLocale)
                Dim useFixedDuplicateHeaders As Boolean = String.Equals(exportKey, "dupclash", StringComparison.OrdinalIgnoreCase)
                Dim defaultGroupHeader As String =
                    If(useFixedDuplicateHeaders,
                       groupHeader,
                       If(String.Equals(exportLocaleKey, "en", StringComparison.OrdinalIgnoreCase),
                          ResolveEnglishExportHeaderName(groupHeader, exportKey),
                          groupHeader))
                Dim exportGroupHeader As String =
                    If(useFixedDuplicateHeaders,
                       groupHeader,
                       ReviewExportMessageOverrideService.ResolveHeader(exportKey, groupHeader, defaultGroupHeader, exportLocaleKey))

                ' 1) 값 쓰기 (AutoFit은 스타일 적용 후 마지막에)
                WriteTableToSheet(wb, sh, safeSheet, exportTable, sheetKey:=title, autoFit:=False, progressKey:=progressKey, exportKind:=exportKind)

                ' 2) 그룹 밴딩 (DuplicateExport: "Group" 컬럼)
                If Not String.IsNullOrWhiteSpace(exportGroupHeader) Then
                    TryApplyGroupBanding(wb, sh, exportTable, exportGroupHeader)
                End If

                ' 3) 기본 시트 스타일 (Freeze/Filter/Border/AutoFit)
                ApplyStandardSheetStyle(wb, sh, headerRowIndex:=0, autoFilter:=True, freezeTopRow:=True, borderAll:=True, autoFit:=autoFit)
                ReportExcelPhase(progressKey, "EXCEL_SAVE", "엑셀 파일 저장 중...", 0, 1, 0, True)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using

            If autoFit Then
                ReportExcelPhase(progressKey, "AUTOFIT", "열 너비 자동 조정 중...", 0, 1, 0, True)
                TryAutoFitWithExcel(filePath)
            End If
            ReportExcelPhase(progressKey, "DONE", "엑셀 저장 완료", 1, 1, 100, True)
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
            "familysuitability",
            "floorinfo",
            "linkworkset",
            "paramprop",
            "pms",
            "sharedparambatch",
            "tapalign",
            "worksetassignment",
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
            If s.Contains("floorinfo") Then Return "floorinfo"
            If s.Contains("familysuitability") Then Return "familysuitability"
            If s.Contains("parametermissing") OrElse s.Contains("parameter missing") Then Return "parametermissing"
            If s.Contains("linkworkset") OrElse s.Contains("link workset") Then Return "linkworkset"
            If s.Contains("worksetassignment") Then Return "worksetassignment"
            If s.Contains("tapalign") Then Return "tapalign"
            If s.Contains("sharedparambatch") Then Return "sharedparambatch"
            If s.Contains("param") Then Return "paramprop"
            If s.Contains("pms") OrElse s.Contains("segment") Then Return "pms"
            If s.Contains("connector") Then Return "connector"

            Return s
        End Function

        Private Function NormalizeExcelLocale(exportLocale As String) As String
            Dim normalized As String = If(exportLocale, String.Empty).Trim().ToLowerInvariant()
            If normalized = "en" OrElse normalized = "eng" OrElse normalized = "english" Then Return "en"
            Return "ko"
        End Function

        Private Function PrepareTableForExport(table As DataTable,
                                               exportKind As String,
                                               exportLocale As String) As DataTable
            If table Is Nothing Then Return Nothing

            Dim exportKey As String = NormalizeExportPolicyKey(exportKind, table.TableName, table.TableName)
            Dim locale As String = NormalizeExcelLocale(exportLocale)
            Dim result As New DataTable(If(table.TableName, String.Empty))
            Dim mappedNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim canonicalNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim useFixedDuplicateHeaders As Boolean = String.Equals(exportKey, "dupclash", StringComparison.OrdinalIgnoreCase)

            For Each sourceColumn As DataColumn In table.Columns
                If ShouldSkipExportHelperColumn(sourceColumn.ColumnName) Then Continue For
                Dim canonicalName As String =
                    If(useFixedDuplicateHeaders,
                       sourceColumn.ColumnName,
                       ResolveEnglishExportHeaderName(sourceColumn.ColumnName, exportKey))
                Dim defaultHeader As String =
                    If(useFixedDuplicateHeaders,
                       sourceColumn.ColumnName,
                       If(String.Equals(locale, "en", StringComparison.OrdinalIgnoreCase),
                          canonicalName,
                          sourceColumn.ColumnName))
                Dim mappedName As String =
                    If(useFixedDuplicateHeaders,
                       sourceColumn.ColumnName,
                       ReviewExportMessageOverrideService.ResolveHeader(exportKey, sourceColumn.ColumnName, defaultHeader, locale))
                mappedName = MakeUniqueColumnName(mappedName, result)

                Dim added As DataColumn = result.Columns.Add(mappedName, sourceColumn.DataType)
                If sourceColumn.ExtendedProperties.Contains(HiddenColumnFlag) Then
                    added.ExtendedProperties(HiddenColumnFlag) = sourceColumn.ExtendedProperties(HiddenColumnFlag)
                End If
                mappedNames(sourceColumn.ColumnName) = mappedName
                canonicalNames(sourceColumn.ColumnName) = canonicalName
            Next

            For Each sourceRow As DataRow In table.Rows
                Dim targetRow As DataRow = result.NewRow()
                For Each sourceColumn As DataColumn In table.Columns
                    If ShouldSkipExportHelperColumn(sourceColumn.ColumnName) Then Continue For
                    Dim targetName As String = mappedNames(sourceColumn.ColumnName)
                    Dim canonicalName As String = canonicalNames(sourceColumn.ColumnName)
                    Dim sourceValue As Object = ResolveExportCellValue(sourceRow, sourceColumn, exportKey, locale)
                    targetRow(targetName) = TranslateExportCellValue(sourceValue, sourceColumn.ColumnName, canonicalName, exportKey, locale)
                Next
                result.Rows.Add(targetRow)
            Next

            Return result
        End Function

        Private Function ShouldSkipExportHelperColumn(columnName As String) As Boolean
            Dim text As String = If(columnName, String.Empty).Trim()
            Return text.Equals("__ReviewEn", StringComparison.OrdinalIgnoreCase) OrElse
                   text.Equals("__ReviewKo", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function ResolveExportCellValue(row As DataRow,
                                                sourceColumn As DataColumn,
                                                exportKey As String,
                                                locale As String) As Object
            If row Is Nothing OrElse sourceColumn Is Nothing Then Return Nothing

            Dim relocatedNoDataValue As Object = Nothing
            If TryResolveLocalizedNoDataCellValue(row, sourceColumn, exportKey, locale, relocatedNoDataValue) Then
                Return relocatedNoDataValue
            End If

            If String.Equals(exportKey, "familysuitability", StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(sourceColumn.ColumnName, "Review", StringComparison.OrdinalIgnoreCase) AndAlso
               row.Table IsNot Nothing Then
                Dim localizedColumnName As String = If(locale = "en", "__ReviewEn", "__ReviewKo")
                If row.Table.Columns.Contains(localizedColumnName) Then
                    Return row(localizedColumnName)
                End If
            End If

            Return row(sourceColumn)
        End Function

        Private Function TryResolveLocalizedNoDataCellValue(row As DataRow,
                                                            sourceColumn As DataColumn,
                                                            exportKey As String,
                                                            locale As String,
                                                            ByRef relocatedValue As Object) As Boolean
            relocatedValue = Nothing

            If row Is Nothing OrElse sourceColumn Is Nothing Then Return False
            If Not String.Equals(locale, "en", StringComparison.OrdinalIgnoreCase) Then Return False
            If row.Table Is Nothing Then Return False

            Select Case NormalizeExportPolicyKey(exportKey, row.Table.TableName, row.Table.TableName)
                Case "familysuitability"
                    If Not String.Equals(sourceColumn.ColumnName, "Review", StringComparison.OrdinalIgnoreCase) Then Return False
                    If Not row.Table.Columns.Contains("Category") Then Return False
                    If Not String.IsNullOrWhiteSpace(ReadRowText(row, "Review")) Then Return False

                    Dim noDataText As String = ReadRowText(row, "Category")
                    If Not IsNoDataMessage(noDataText, "집계 가능한 객체가 없습니다.", "No rows to export.") Then Return False

                    relocatedValue = noDataText
                    Return True

                Case "floorinfo"
                    If Not String.Equals(sourceColumn.ColumnName, "Note", StringComparison.OrdinalIgnoreCase) Then Return False
                    If Not row.Table.Columns.Contains("File") Then Return False
                    If Not String.IsNullOrWhiteSpace(ReadRowText(row, "Note")) Then Return False

                    Dim noDataText As String = ReadRowText(row, "File")
                    If Not IsNoDataMessage(noDataText, "오류가 없습니다.", "No issues.") Then Return False

                    relocatedValue = noDataText
                    Return True
            End Select

            Return False
        End Function

        Private Function ReadRowText(row As DataRow, columnName As String) As String
            If row Is Nothing OrElse row.Table Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return String.Empty
            If Not row.Table.Columns.Contains(columnName) Then Return String.Empty

            Dim value As Object = row(columnName)
            If value Is Nothing OrElse value Is DBNull.Value Then Return String.Empty

            Return Convert.ToString(value).Trim()
        End Function

        Private Function IsNoDataMessage(value As String, ParamArray candidates() As String) As Boolean
            Dim text As String = If(value, String.Empty).Trim()
            If text = String.Empty OrElse candidates Is Nothing Then Return False

            For Each candidate In candidates
                If String.Equals(text, If(candidate, String.Empty).Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function MakeUniqueColumnName(baseName As String, table As DataTable) As String
            Dim name As String = If(String.IsNullOrWhiteSpace(baseName), "Column", baseName.Trim())
            If table Is Nothing OrElse Not table.Columns.Contains(name) Then Return name

            Dim seed As String = name
            Dim suffix As Integer = 2
            While table.Columns.Contains(name)
                name = seed & "_" & suffix.ToString()
                suffix += 1
            End While
            Return name
        End Function

        Private Function ResolveEnglishExportHeaderName(header As String,
                                                       Optional exportKind As String = Nothing) As String
            Dim text As String = If(header, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return "Column"

            Dim exportKey As String = NormalizeExportPolicyKey(exportKind, text, text)

            If String.Equals(exportKey, "connector", StringComparison.OrdinalIgnoreCase) Then
                Select Case text
                    Case "검토내용"
                        Return "Review Details"
                    Case "비고", "비고(답변)"
                        Return "Comments"
                    Case "ConnectionType"
                        Return "Connection Type"
                    Case "Distance (inch)", "Distance (mm)", "Distance"
                        Return "Distance (inch or mm)"
                End Select
            End If

            Select Case text
                Case "파일"
                    Return "File"
                Case "파일명"
                    Return "FileName"
                Case "요소 ID"
                    Return "ElementId"
                Case "카테고리"
                    Return "Category"
                Case "패밀리"
                    Return "Family"
                Case "타입"
                    Return "Type"
                Case "연결 라인 ID"
                    Return "Connected Host Id"
                Case "연결 라인 카테고리"
                    Return "Connected Host Category"
                Case "연결 라인 타입"
                    Return "Connected Host Type"
                Case "공종"
                    Return "Domain"
                Case "메시지"
                    Return "Message"
                Case "검토내용"
                    Return "Review"
                Case "비고", "비고(답변)"
                    Return "Notes"
                Case "상세"
                    Return "Detail"
                Case "작업"
                    Return "Task"
                Case "항목"
                    Return "Item"
                Case "상태"
                    Return "Status"
                Case "기록"
                    Return "Record"
                Case "정리 전 객체수"
                    Return "Before Count"
                Case "정리 후 객체수"
                    Return "After Count"
                Case "증감"
                    Return "Delta"
                Case "삭제여부"
                    Return "Delete"
                Case "파라미터명"
                    Return "ParamName"
                Case "구분"
                    Return "Scope"
                Case "적용카테고리"
                    Return "BoundCategories"
                Case "현재 GUID"
                    Return "CurrentGuid"
                Case "기준 GUID"
                    Return "ExpectedGuid"
                Case "검토결과"
                    Return "Result"
                Case "패밀리명"
                    Return "FamilyName"
                Case "패밀리카테고리"
                    Return "FamilyCategory"
                Case "ConnectionType"
                    Return "Connection Type"
            End Select

            If text.StartsWith("중심축으로부터 거리 (", StringComparison.OrdinalIgnoreCase) Then
                Return "Distance From Center (" & text.Substring("중심축으로부터 거리 (".Length)
            End If
            If String.Equals(text, "모델링된 각도 (deg)", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(text, "XY 평면 기준 각도 (deg)", StringComparison.OrdinalIgnoreCase) Then
                Return "Angle To XY Plane (deg)"
            End If
            If text.EndsWith(" (분기 객체)", StringComparison.OrdinalIgnoreCase) Then
                Return text.Substring(0, text.Length - " (분기 객체)".Length) & " (Branch Element)"
            End If
            If text.EndsWith(" (연결 라인)", StringComparison.OrdinalIgnoreCase) Then
                Return text.Substring(0, text.Length - " (연결 라인)".Length) & " (Connected Host)"
            End If
            If text.EndsWith("검토결과", StringComparison.OrdinalIgnoreCase) Then
                Return text.Substring(0, text.Length - "검토결과".Length) & " Review Result"
            End If
            If text.EndsWith("검토", StringComparison.OrdinalIgnoreCase) Then
                Return text.Substring(0, text.Length - "검토".Length) & " Review"
            End If

            Return text
        End Function

        Private Function TranslateExportCellValue(value As Object,
                                                  originalHeader As String,
                                                  exportHeader As String,
                                                  exportKind As String,
                                                  locale As String) As Object
            If value Is Nothing OrElse value Is DBNull.Value Then Return value
            If Not TypeOf value Is String Then Return value

            Dim sourceText As String = DirectCast(value, String)
            If String.IsNullOrWhiteSpace(sourceText) Then Return sourceText

            Dim translatedText As String =
                If(String.Equals(locale, "en", StringComparison.OrdinalIgnoreCase),
                   TranslateExportCellText(sourceText, exportHeader, exportKind, locale),
                   sourceText)
            Dim resolvedOverride As String = Nothing
            If ReviewExportMessageOverrideService.TryResolve(exportKind, originalHeader, exportHeader, sourceText, translatedText, locale, resolvedOverride) Then
                Return resolvedOverride
            End If

            Return translatedText
        End Function

        Private Function TranslateExportCellText(text As String,
                                                 exportHeader As String,
                                                 exportKind As String,
                                                 locale As String) As String
            text = TranslateNoDataText(text, locale)
            Select Case NormalizeExportPolicyKey(exportKind, exportHeader, exportHeader)
                Case "tapalign"
                    Return TranslateTapAlignExportCell(text, exportHeader, locale)
                Case "connector"
                    Return TranslateConnectorExportCell(text, exportHeader, locale)
                Case "guid"
                    Return TranslateGuidExportCell(text, exportHeader, locale)
                Case "familylink"
                    Return TranslateFamilyLinkExportCell(text, exportHeader, locale)
                Case "floorinfo"
                    Return TranslateFloorInfoExportCell(text, exportHeader, locale)
                Case "familysuitability"
                    Return TranslateFamilySuitabilityExportCell(text, exportHeader, locale)
                Case "pms"
                    Return TranslateSegmentPmsExportCell(text, exportHeader, locale)
                Case "linkworkset"
                    Return TranslateLinkWorksetExportCell(text, exportHeader, locale)
                Case "worksetassignment"
                    Return TranslateWorksetAssignmentExportCell(text, exportHeader, locale)
                Case "parametermissing"
                    Return TranslateParameterMissingExportCell(text, exportHeader, locale)
                Case Else
                    If ShouldTranslateGenericExportText(exportHeader) Then
                        Return TranslateCommonLocalizedText(text, locale)
                    End If
                    Return text
            End Select
        End Function

        Private Function ShouldTranslateGenericExportText(exportHeader As String) As Boolean
            Dim text As String = If(exportHeader, String.Empty).Trim().ToLowerInvariant()
            If text = String.Empty Then Return False

            Return text.Contains("review") OrElse
                   text.Contains("note") OrElse
                   text.Contains("message") OrElse
                   text.Contains("issue") OrElse
                   text.Contains("result") OrElse
                   text.Contains("status") OrElse
                   text.Contains("level") OrElse
                   text.Contains("detail") OrElse
                   text.Contains("remark") OrElse
                   text.Contains("comment") OrElse
                   text.Contains("content") OrElse
                   text.Contains("solution")
        End Function

        Private Function TranslateNoDataText(text As String, locale As String) As String
            Dim trimmed As String = If(text, String.Empty).Trim()
            If trimmed = String.Empty Then Return text

            Select Case trimmed
                Case "오류가 없습니다.", "이상 없음."
                    Return If(locale = "en", "No issues.", "오류가 없습니다.")
                Case "No issues."
                    Return If(locale = "en", "No issues.", "오류가 없습니다.")
                Case "검토 결과가 없습니다."
                    Return If(locale = "en", "No review results.", "검토 결과가 없습니다.")
                Case "No review results."
                    Return If(locale = "en", "No review results.", "검토 결과가 없습니다.")
                Case "집계 가능한 객체가 없습니다."
                    Return If(locale = "en", "No rows to export.", "집계 가능한 객체가 없습니다.")
                Case "No rows to export."
                    Return If(locale = "en", "No rows to export.", "집계 가능한 객체가 없습니다.")
                Case Else
                    Return text
            End Select
        End Function

        Private Function TranslateTapAlignExportCell(text As String, exportHeader As String, locale As String) As String
            If String.Equals(exportHeader, "Domain", StringComparison.OrdinalIgnoreCase) Then
                If String.Equals(text, "배관", StringComparison.OrdinalIgnoreCase) Then Return "Pipe"
                If String.Equals(text, "덕트", StringComparison.OrdinalIgnoreCase) Then Return "Duct"
                Return text
            End If

            If Not String.Equals(exportHeader, "Message", StringComparison.OrdinalIgnoreCase) Then
                Return text
            End If

            If locale = "en" Then
                If String.Equals(text, "중심축에서 벗어났습니다.", StringComparison.OrdinalIgnoreCase) Then Return "Offset from center axis."
                Return text
            End If

            If String.Equals(text, "Offset from center axis.", StringComparison.OrdinalIgnoreCase) Then Return "중심축에서 벗어났습니다."
            Return text
        End Function

        Private Function TranslateConnectorExportCell(text As String, exportHeader As String, locale As String) As String
            If String.Equals(exportHeader, "Review Details", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateConnectorReviewDetailsText(text, locale)
            End If

            If String.Equals(exportHeader, "Comments", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateConnectorCommentText(text, locale)
            End If

            If String.Equals(exportHeader, "Connection Type", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateConnectorConnectionTypeValue(text, locale)
            End If

            If Not ShouldTranslateGenericExportText(exportHeader) Then Return text
            Return TranslateConnectorText(text, locale)
        End Function

        Private Function TranslateGuidExportCell(text As String, exportHeader As String, locale As String) As String
            If String.Equals(exportHeader, "Scope", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateGuidScopeValue(text)
            End If

            If String.Equals(exportHeader, "Result", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateGuidResultValue(text, locale)
            End If

            If String.Equals(exportHeader, "Notes", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateGuidNotes(text, locale)
            End If

            Return text
        End Function

        Private Function TranslateFamilyLinkExportCell(text As String, exportHeader As String, locale As String) As String
            If String.Equals(exportHeader, "Issue", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateFamilyLinkIssue(text, locale)
            End If

            If String.Equals(exportHeader, "Notes", StringComparison.OrdinalIgnoreCase) Then
                Return TranslateFamilyLinkNotes(text, locale)
            End If

            Return text
        End Function

        Private Function TranslateFloorInfoExportCell(text As String, exportHeader As String, locale As String) As String
            If Not ShouldTranslateGenericExportText(exportHeader) Then Return text
            Return TranslateFloorInfoText(text, locale)
        End Function

        Private Function TranslateFamilySuitabilityExportCell(text As String, exportHeader As String, locale As String) As String
            If Not String.Equals(exportHeader, "Review", StringComparison.OrdinalIgnoreCase) Then Return text
            Return TranslateCommonLocalizedText(text, locale)
        End Function

        Private Function TranslateSegmentPmsExportCell(text As String, exportHeader As String, locale As String) As String
            Dim header As String = If(exportHeader, String.Empty).Trim().ToLowerInvariant()
            If header.Contains("result") OrElse header.Contains("review") OrElse header.Contains("status") Then
                Dim translatedStatus As String = TranslateSegmentPmsStatusValue(text, locale)
                If Not String.Equals(translatedStatus, text, StringComparison.Ordinal) Then
                    Return translatedStatus
                End If
            End If

            If ShouldTranslateGenericExportText(exportHeader) Then
                Return TranslateCommonLocalizedText(text, locale)
            End If

            Return text
        End Function

        Private Function TranslateLinkWorksetExportCell(text As String, exportHeader As String, locale As String) As String
            If String.Equals(exportHeader, "Status", StringComparison.OrdinalIgnoreCase) Then
                Select Case If(text, String.Empty).Trim().ToLowerInvariant()
                    Case "ok", "정상"
                        Return If(locale = "en", "ok", "정상")
                    Case "changed", "변경됨"
                        Return If(locale = "en", "changed", "변경됨")
                    Case "warning", "경고"
                        Return If(locale = "en", "warning", "경고")
                    Case "skipped", "건너뜀"
                        Return If(locale = "en", "skipped", "건너뜀")
                End Select
            End If

            If ShouldTranslateGenericExportText(exportHeader) Then
                Return TranslateCommonLocalizedText(text, locale)
            End If

            Return text
        End Function

        Private Function TranslateWorksetAssignmentExportCell(text As String, exportHeader As String, locale As String) As String
            If Not ShouldTranslateGenericExportText(exportHeader) Then Return text
            Return TranslateCommonLocalizedText(text, locale)
        End Function

        Private Function TranslateParameterMissingExportCell(text As String, exportHeader As String, locale As String) As String
            Dim value As String = If(text, String.Empty).Trim()
            If value = String.Empty Then Return text

            If locale = "en" Then
                If String.Equals(exportHeader, "Item", StringComparison.OrdinalIgnoreCase) Then
                    Dim itemMatch = System.Text.RegularExpressions.Regex.Match(value, "^속성누락검토 오류\s*\((\d+)건\)$")
                    If itemMatch.Success Then
                        Return $"Parameter Value omission check({itemMatch.Groups(1).Value} cases)"
                    End If
                End If

                If String.Equals(exportHeader, "Content", StringComparison.OrdinalIgnoreCase) Then
                    Dim emptyValueMatch = System.Text.RegularExpressions.Regex.Match(value, "^\[Parameter\]:\s*\[(.+?)\]\s*값이 누락입니다\.$")
                    If emptyValueMatch.Success Then
                        Return $"[Parameter]: [{emptyValueMatch.Groups(1).Value}] value is empty."
                    End If

                    Dim notFoundMatch = System.Text.RegularExpressions.Regex.Match(value, "^\[Parameter\]\s*:\s*\[(.+?)\]\s*파라미터가 존재 하지 않습니다\.$")
                    If notFoundMatch.Success Then
                        Return $"[Parameter] : [{notFoundMatch.Groups(1).Value}] does not exist."
                    End If

                    Dim okMatch = System.Text.RegularExpressions.Regex.Match(value, "^(\d+)개 요소의 데이터가 입력 되어 있습니다 !!!$")
                    If okMatch.Success Then
                        Return $"parameter values for {okMatch.Groups(1).Value} elements have been entered."
                    End If

                    If String.Equals(value, "검사 대상 요소가 없습니다 !!!", StringComparison.OrdinalIgnoreCase) Then
                        Return "No target elements found."
                    End If
                End If

                If String.Equals(exportHeader, "Solutions", StringComparison.OrdinalIgnoreCase) Then
                    If String.Equals(value, "속성 값을 입력해주세요.", StringComparison.OrdinalIgnoreCase) Then
                        Return "Please enter the property values."
                    End If
                End If

                Return value
            End If

            If String.Equals(exportHeader, "Item", StringComparison.OrdinalIgnoreCase) Then
                Dim itemMatch = System.Text.RegularExpressions.Regex.Match(value, "^Parameter Value omission check\((\d+) cases\)$")
                If itemMatch.Success Then
                    Return $"속성누락검토 오류 ({itemMatch.Groups(1).Value}건)"
                End If
            End If

            If String.Equals(exportHeader, "Content", StringComparison.OrdinalIgnoreCase) Then
                Dim emptyValueMatch = System.Text.RegularExpressions.Regex.Match(value, "^\[Parameter\]:\s*\[(.+?)\]\s*value is empty\.$")
                If emptyValueMatch.Success Then
                    Return $"[Parameter]: [{emptyValueMatch.Groups(1).Value}] 값이 누락입니다."
                End If

                Dim notFoundMatch = System.Text.RegularExpressions.Regex.Match(value, "^\[Parameter\]\s*:\s*\[(.+?)\]\s*does not exist\.$")
                If notFoundMatch.Success Then
                    Return $"[Parameter] : [{notFoundMatch.Groups(1).Value}] 파라미터가 존재 하지 않습니다."
                End If

                Dim okMatch = System.Text.RegularExpressions.Regex.Match(value, "^parameter values for (\d+) elements have been entered\.$")
                If okMatch.Success Then
                    Return $"{okMatch.Groups(1).Value}개 요소의 데이터가 입력 되어 있습니다 !!!"
                End If

                If String.Equals(value, "No target elements found.", StringComparison.OrdinalIgnoreCase) Then
                    Return "검사 대상 요소가 없습니다 !!!"
                End If
            End If

            If String.Equals(exportHeader, "Solutions", StringComparison.OrdinalIgnoreCase) Then
                If String.Equals(value, "Please enter the property values.", StringComparison.OrdinalIgnoreCase) Then
                    Return "속성 값을 입력해주세요."
                End If
            End If

            Return value
        End Function

        Private Function TranslateSegmentPmsStatusValue(text As String, locale As String) As String
            Select Case If(text, String.Empty).Trim()
                Case "OK", "정상"
                    Return If(locale = "en", "OK", "정상")
                Case "Mismatch", "불일치"
                    Return If(locale = "en", "Mismatch", "불일치")
                Case "MismatchID", "ID 불일치"
                    Return If(locale = "en", "MismatchID", "ID 불일치")
                Case "MismatchOD", "OD 불일치"
                    Return If(locale = "en", "MismatchOD", "OD 불일치")
                Case "MissingPMS", "PMS 누락"
                    Return If(locale = "en", "MissingPMS", "PMS 누락")
                Case "MissingPmsRow", "PMS 행 누락"
                    Return If(locale = "en", "MissingPmsRow", "PMS 행 누락")
                Case "MissingMapping", "매핑 누락"
                    Return If(locale = "en", "MissingMapping", "매핑 누락")
                Case "MissingRevit", "Revit 누락"
                    Return If(locale = "en", "MissingRevit", "Revit 누락")
                Case "MissingRevitRow", "Revit 행 누락"
                    Return If(locale = "en", "MissingRevitRow", "Revit 행 누락")
                Case "N/A", "해당 없음"
                    Return If(locale = "en", "N/A", "해당 없음")
                Case Else
                    Return text
            End Select
        End Function

        Private Function TranslateCommonLocalizedText(text As String, locale As String) As String
            If locale = "en" Then
                Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("검토 필요", "Needs review"),
                    New KeyValuePair(Of String, String)("기준 미일치", "Does not match criteria"),
                    New KeyValuePair(Of String, String)("기준 일치", "Matches criteria"),
                    New KeyValuePair(Of String, String)("불일치", "Mismatch"),
                    New KeyValuePair(Of String, String)("누락", "Missing"),
                    New KeyValuePair(Of String, String)("오류", "Error"),
                    New KeyValuePair(Of String, String)("정상", "OK"),
                    New KeyValuePair(Of String, String)("경고", "Warning"),
                    New KeyValuePair(Of String, String)("성공", "Success"),
                    New KeyValuePair(Of String, String)("실패", "Fail"),
                    New KeyValuePair(Of String, String)("건너뜀", "Skipped"),
                    New KeyValuePair(Of String, String)("변경됨", "Changed"),
                    New KeyValuePair(Of String, String)("해당 없음", "N/A"),
                    New KeyValuePair(Of String, String)("미지원", "Unsupported"),
                    New KeyValuePair(Of String, String)("검사 대상 요소가 없습니다 !!!", "No target elements found."),
                    New KeyValuePair(Of String, String)("검토 대상 객체가 없습니다.", "No target elements found"),
                    New KeyValuePair(Of String, String)("중복된 객체가 없습니다.", "No duplicated elements found."),
                    New KeyValuePair(Of String, String)("중복 객체가 없습니다.", "No duplicated elements found."),
                    New KeyValuePair(Of String, String)("객체를 확인 후 중복 객체를 수정해주세요.", "Please review duplicated elements by referring to the report and delete unnecessary elements."),
                    New KeyValuePair(Of String, String)("리포트를 참고하여 중복 요소를 검토하고 불필요한 요소를 삭제해주세요.", "Please review duplicated elements by referring to the report and delete unnecessary elements.")
                })
            End If

            Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("Review required", "검토 필요"),
                New KeyValuePair(Of String, String)("Needs review", "검토 필요"),
                New KeyValuePair(Of String, String)("Does not match criteria", "기준 미일치"),
                New KeyValuePair(Of String, String)("Matches criteria", "기준 일치"),
                New KeyValuePair(Of String, String)("Mismatch", "불일치"),
                New KeyValuePair(Of String, String)("Missing", "누락"),
                New KeyValuePair(Of String, String)("Error", "오류"),
                New KeyValuePair(Of String, String)("Warning", "경고"),
                New KeyValuePair(Of String, String)("Success", "성공"),
                New KeyValuePair(Of String, String)("Failed", "실패"),
                New KeyValuePair(Of String, String)("Failure", "실패"),
                New KeyValuePair(Of String, String)("Fail", "실패"),
                New KeyValuePair(Of String, String)("Skipped", "건너뜀"),
                New KeyValuePair(Of String, String)("Changed", "변경됨"),
                New KeyValuePair(Of String, String)("Unsupported", "미지원"),
                New KeyValuePair(Of String, String)("N/A", "해당 없음"),
                New KeyValuePair(Of String, String)("OK", "정상"),
                New KeyValuePair(Of String, String)("Match", "일치"),
                New KeyValuePair(Of String, String)("No duplicated elements found.", "중복 객체가 없습니다."),
                New KeyValuePair(Of String, String)("No duplicated elements found", "중복 객체가 없습니다."),
                New KeyValuePair(Of String, String)("No target elements found.", "검토 대상 객체가 없습니다."),
                New KeyValuePair(Of String, String)("No target elements found", "검토 대상 객체가 없습니다."),
                New KeyValuePair(Of String, String)("Please review duplicated elements by referring to the report and delete unnecessary elements.", "객체를 확인 후 중복 객체를 수정해주세요.")
            })
        End Function

        Private Function TranslateConnectorText(text As String, locale As String) As String
            If locale = "en" Then
                Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("값이 서로 불일치. 확인이 필요합니다.", "Values do not match. Review required."),
                    New KeyValuePair(Of String, String)("파라미터 속성이 모두 누락되어있습니다.", "All parameter attributes are missing."),
                    New KeyValuePair(Of String, String)("Shared Parameter 등록 필요", "Shared Parameter registration required"),
                    New KeyValuePair(Of String, String)("연결 대상 객체 없음", "No connected target object"),
                    New KeyValuePair(Of String, String)("연결 대상 없음", "No connected target"),
                    New KeyValuePair(Of String, String)("검토 필요", "Review required")
                })
            End If

            Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("Values do not match. Review required.", "값이 서로 불일치. 확인이 필요합니다."),
                New KeyValuePair(Of String, String)("All parameter attributes are missing.", "파라미터 속성이 모두 누락되어있습니다."),
                New KeyValuePair(Of String, String)("Shared Parameter registration required", "Shared Parameter 등록 필요"),
                New KeyValuePair(Of String, String)("No connected target object", "연결 대상 객체 없음"),
                New KeyValuePair(Of String, String)("No connected target", "연결 대상 없음"),
                New KeyValuePair(Of String, String)("Review required", "검토 필요")
            })
        End Function

        Private Function TranslateConnectorReviewDetailsText(text As String, locale As String) As String
            Dim value As String = If(text, String.Empty).Trim()
            If value = String.Empty Then Return text

            If locale = "en" Then
                Dim withParam = System.Text.RegularExpressions.Regex.Match(value, "^\[(.+?)\]\s*값이 서로 불일치\.\s*확인이 필요합니다\.$")
                If withParam.Success Then Return $"[{withParam.Groups(1).Value}] Value mismatch. Please check."
                Dim withNoErrorParam = System.Text.RegularExpressions.Regex.Match(value, "^\[(.+?)\]\s*파라미터(?:에 대한)? 연속성 오류가 없습니다\.$")
                If withNoErrorParam.Success Then Return $"[{withNoErrorParam.Groups(1).Value}] No parameter continuity errors."
                If String.Equals(value, "값이 서로 불일치. 확인이 필요합니다.", StringComparison.OrdinalIgnoreCase) Then Return "Value mismatch. Please check."
                If String.Equals(value, "검토 대상 객체가 없습니다.", StringComparison.OrdinalIgnoreCase) Then Return "No target elements found."
                If value.EndsWith(" : Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then
                    Dim paramName As String = value.Substring(0, value.Length - " : Shared Parameter 등록 필요".Length).Trim()
                    If paramName <> String.Empty Then Return $"{paramName}: Shared Parameter registration required"
                End If
                Return ApplyTextReplacements(value, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("파라미터 속성이 모두 누락되어있습니다.", "All parameter attributes are missing."),
                    New KeyValuePair(Of String, String)("Shared Parameter 등록 필요", "Shared Parameter registration required"),
                    New KeyValuePair(Of String, String)("연결 대상 객체 없음", "No connected target object"),
                    New KeyValuePair(Of String, String)("파라미터 연속성 오류가 없습니다.", "No parameter continuity errors."),
                    New KeyValuePair(Of String, String)("파라미터에 대한 연속성 오류가 없습니다.", "No parameter continuity errors."),
                    New KeyValuePair(Of String, String)("연속성 오류가 없습니다.", "No continuity issue found."),
                    New KeyValuePair(Of String, String)("검토 대상 객체가 없습니다.", "No target elements found."),
                    New KeyValuePair(Of String, String)("검토 필요", "Please review.")
                })
            End If

            Dim withParamEn = System.Text.RegularExpressions.Regex.Match(value, "^\[(.+?)\]\s*Value mismatch\.\s*Please check\.$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If withParamEn.Success Then Return $"[{withParamEn.Groups(1).Value}] 값이 서로 불일치. 확인이 필요합니다."
            Dim withNoErrorParamEn = System.Text.RegularExpressions.Regex.Match(value, "^\[(.+?)\]\s*No parameter continuity errors\.$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If withNoErrorParamEn.Success Then Return $"[{withNoErrorParamEn.Groups(1).Value}] 파라미터 연속성 오류가 없습니다."
            If String.Equals(value, "Value mismatch. Please check.", StringComparison.OrdinalIgnoreCase) Then Return "값이 서로 불일치. 확인이 필요합니다."
            If String.Equals(value, "No target elements found.", StringComparison.OrdinalIgnoreCase) Then Return "검토 대상 객체가 없습니다."
            If value.EndsWith(": Shared Parameter registration required", StringComparison.OrdinalIgnoreCase) Then
                Dim paramName As String = value.Substring(0, value.Length - ": Shared Parameter registration required".Length).Trim()
                If paramName <> String.Empty Then Return $"{paramName} : Shared Parameter 등록 필요"
            End If
            Return ApplyTextReplacements(value, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("All parameter attributes are missing.", "파라미터 속성이 모두 누락되어있습니다."),
                New KeyValuePair(Of String, String)("Shared Parameter registration required", "Shared Parameter 등록 필요"),
                New KeyValuePair(Of String, String)("No connected target object", "연결 대상 객체 없음"),
                New KeyValuePair(Of String, String)("No parameter continuity errors.", "파라미터 연속성 오류가 없습니다."),
                New KeyValuePair(Of String, String)("No continuity issue found.", "연속성 오류가 없습니다."),
                New KeyValuePair(Of String, String)("No target elements found.", "검토 대상 객체가 없습니다."),
                New KeyValuePair(Of String, String)("Please review.", "검토 필요")
            })
        End Function

        Private Function TranslateConnectorCommentText(text As String, locale As String) As String
            Dim value As String = If(text, String.Empty).Trim()
            If value = String.Empty Then Return text

            Dim isMatchComment As Boolean =
                String.Equals(value, "Value1과 Value2가 일치합니다.", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(value, "Value1 and Value2 match.", StringComparison.OrdinalIgnoreCase)

            Dim isMismatchComment As Boolean =
                String.Equals(value, "Value1과 Value2를 일치시켜 주세요.", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(value, "Match Value1 and Value2", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(value, "Match Value1 and Value2.", StringComparison.OrdinalIgnoreCase)

            If locale = "en" Then
                If isMatchComment Then Return "Value1 and Value2 match."
                If isMismatchComment Then Return "Match Value1 and Value2"
                Return value
            End If

            If isMatchComment Then Return "Value1과 Value2가 일치합니다."
            If isMismatchComment Then Return "Value1과 Value2를 일치시켜 주세요."
            Return value
        End Function

        Private Function TranslateConnectorConnectionTypeValue(text As String, locale As String) As String
            Dim value As String = If(text, String.Empty).Trim()
            If value = String.Empty Then Return text

            Dim isDisconnected As Boolean =
                value.IndexOf("proximity", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("disconnected", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("미연결", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("연결 필요", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0

            Dim isConnected As Boolean =
                value.IndexOf("physical", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                (value.IndexOf("connected", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso Not isDisconnected) OrElse
                value.IndexOf("연결됨", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                value.IndexOf("연결 됨", StringComparison.OrdinalIgnoreCase) >= 0

            If isConnected Then Return If(locale = "en", "Connected", "연결됨")
            If isDisconnected Then Return If(locale = "en", "Disconnected", "미연결")

            If String.Equals(value, "Connected", StringComparison.OrdinalIgnoreCase) Then Return If(locale = "en", "Connected", "연결됨")
            If String.Equals(value, "Disconnected", StringComparison.OrdinalIgnoreCase) Then Return If(locale = "en", "Disconnected", "미연결")

            Return value
        End Function

        Private Function TranslateGuidScopeValue(text As String) As String
            Select Case If(text, String.Empty).Trim()
                Case "공유"
                    Return "Shared"
                Case "프로젝트"
                    Return "Project"
                Case "내장"
                    Return "BuiltIn"
                Case "패밀리"
                    Return "Family"
                Case Else
                    Return text
            End Select
        End Function

        Private Function TranslateGuidResultValue(text As String, locale As String) As String
            Dim value As String = If(text, String.Empty).Trim()
            Select Case value.ToUpperInvariant()
                Case "OK", "일치"
                    Return If(locale = "en", "Match", "일치")
                Case "OK(MULTI_IN_FILE)", "일치(기준파일 중복)"
                    Return If(locale = "en", "Match (duplicate GUIDs in source file)", "일치(기준파일 중복)")
                Case "MISMATCH", "불일치"
                    Return If(locale = "en", "Mismatch", "불일치")
                Case "NOT_FOUND_IN_FILE", "기준파일 없음"
                    Return If(locale = "en", "Not found in source file", "기준파일 없음")
                Case "PROJECT_PARAM", "프로젝트 파라미터"
                    Return If(locale = "en", "Project parameter", "프로젝트 파라미터")
                Case "FAMILY_PARAM", "패밀리 파라미터"
                    Return If(locale = "en", "Family parameter", "패밀리 파라미터")
                Case "BUILTIN", "내장 파라미터"
                    Return If(locale = "en", "Built-in parameter", "내장 파라미터")
                Case "GUID_FAIL", "GUID 추출 실패"
                    Return If(locale = "en", "GUID extraction failed", "GUID 추출 실패")
                Case "OPEN_FAIL", "열기 실패"
                    Return If(locale = "en", "Open failed", "열기 실패")
                Case "ERROR", "오류"
                    Return If(locale = "en", "Error", "오류")
                Case Else
                    Return text
            End Select
        End Function

        Private Function TranslateGuidNotes(text As String, locale As String) As String
            If locale = "en" Then
                Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("RVT의 GUID와 Shared Parameter 파일 GUID 불일치", "RVT GUID does not match the Shared Parameter file GUID"),
                    New KeyValuePair(Of String, String)("Shared Parameter 파일에서 동일 이름을 찾지 못함", "Could not find the same parameter name in the Shared Parameter file"),
                    New KeyValuePair(Of String, String)("파일 내 동일 이름 GUID 여러 개", "Multiple GUIDs with the same name in the source file"),
                    New KeyValuePair(Of String, String)("문서 내 동일 이름 GUID 여러 개", "Multiple GUIDs with the same name in the document")
                })
            End If

            Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("RVT GUID does not match the Shared Parameter file GUID", "RVT의 GUID와 Shared Parameter 파일 GUID 불일치"),
                New KeyValuePair(Of String, String)("Could not find the same parameter name in the Shared Parameter file", "Shared Parameter 파일에서 동일 이름을 찾지 못함"),
                New KeyValuePair(Of String, String)("Multiple GUIDs with the same name in the source file", "파일 내 동일 이름 GUID 여러 개"),
                New KeyValuePair(Of String, String)("Multiple GUIDs with the same name in the document", "문서 내 동일 이름 GUID 여러 개")
            })
        End Function

        Private Function TranslateFamilyLinkIssue(text As String, locale As String) As String
            Select Case If(text, String.Empty).Trim()
                Case "OK"
                    Return If(locale = "en", "OK", "정상")
                Case "GuidMismatch"
                    Return If(locale = "en", "GUID mismatch", "GUID 불일치")
                Case "HostParamNotShared"
                    Return If(locale = "en", "Host parameter not shared", "호스트 파라미터 Shared 아님")
                Case "MissingAssociation"
                    Return If(locale = "en", "Missing association", "연결 누락")
                Case "ParamNotFound"
                    Return If(locale = "en", "Parameter not found", "파라미터 없음")
                Case "DescendantNestedUnsupported"
                    Return If(locale = "en", "Unsupported descendant nesting", "하위 중첩 구조 미지원")
                Case "Error"
                    Return If(locale = "en", "Error", "오류")
                Case Else
                    Return text
            End Select
        End Function

        Private Function TranslateFamilyLinkNotes(text As String, locale As String) As String
            If locale = "en" Then
                Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("중첩 패밀리 파라미터에 호스트 연결(Associate)이 없습니다", "Nested family parameter is missing a host association"),
                    New KeyValuePair(Of String, String)("중첩 패밀리 파라미터 GUID 불일치", "Nested family parameter GUID mismatch"),
                    New KeyValuePair(Of String, String)("중첩 패밀리 파라미터 IsShared=True 이나 GUID 추출 실패(특이 케이스)", "Nested family parameter has IsShared=True but GUID extraction failed"),
                    New KeyValuePair(Of String, String)("중첩 패밀리 파라미터 IsShared=False (Shared 아님, 이름만 일치)", "Nested family parameter has IsShared=False (name matches only)"),
                    New KeyValuePair(Of String, String)("중첩 패밀리 파라미터 Shared 여부 확인 실패(이름만 일치)", "Could not determine whether the nested family parameter is shared (name matches only)"),
                    New KeyValuePair(Of String, String)("연결된 호스트 FamilyParameter가 Shared가 아님", "Associated host FamilyParameter is not shared"),
                    New KeyValuePair(Of String, String)("호스트 파라미터 GUID 불일치", "Host parameter GUID mismatch")
                })
            End If

            Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("Nested family parameter is missing a host association", "중첩 패밀리 파라미터에 호스트 연결(Associate)이 없습니다"),
                New KeyValuePair(Of String, String)("Nested family parameter GUID mismatch", "중첩 패밀리 파라미터 GUID 불일치"),
                New KeyValuePair(Of String, String)("Nested family parameter has IsShared=True but GUID extraction failed", "중첩 패밀리 파라미터 IsShared=True 이나 GUID 추출 실패(특이 케이스)"),
                New KeyValuePair(Of String, String)("Nested family parameter has IsShared=False (name matches only)", "중첩 패밀리 파라미터 IsShared=False (Shared 아님, 이름만 일치)"),
                New KeyValuePair(Of String, String)("Could not determine whether the nested family parameter is shared (name matches only)", "중첩 패밀리 파라미터 Shared 여부 확인 실패(이름만 일치)"),
                New KeyValuePair(Of String, String)("Associated host FamilyParameter is not shared", "연결된 호스트 FamilyParameter가 Shared가 아님"),
                New KeyValuePair(Of String, String)("Host parameter GUID mismatch", "호스트 파라미터 GUID 불일치")
            })
        End Function

        Private Function TranslateFloorInfoText(text As String, locale As String) As String
            If locale = "en" Then
                Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                    New KeyValuePair(Of String, String)("BoundingBox를 가져오지 못했습니다.", "Could not read BoundingBox."),
                    New KeyValuePair(Of String, String)("층정보 값이 비어 있습니다.", "Floor info value is empty."),
                    New KeyValuePair(Of String, String)("검토 필요", "Needs review")
                })
            End If

            Return ApplyTextReplacements(text, New KeyValuePair(Of String, String)() {
                New KeyValuePair(Of String, String)("Could not read BoundingBox.", "BoundingBox를 가져오지 못했습니다."),
                New KeyValuePair(Of String, String)("Floor info value is empty.", "층정보 값이 비어 있습니다."),
                New KeyValuePair(Of String, String)("Needs review", "검토 필요")
            })
        End Function

        Private Function ApplyTextReplacements(text As String,
                                               replacements As IEnumerable(Of KeyValuePair(Of String, String))) As String
            Dim result As String = If(text, String.Empty)
            If replacements Is Nothing OrElse result = String.Empty Then Return result

            For Each pair In replacements
                If String.IsNullOrEmpty(pair.Key) Then Continue For
                result = result.Replace(pair.Key, pair.Value)
            Next
            Return result
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

        Public Sub MarkColumnHidden(column As DataColumn)
            If column Is Nothing Then Return
            column.ExtendedProperties(HiddenColumnFlag) = True
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

                ' ---- 핵심: 저장하면서 행 상태를 판정해서 스타일 적용 ----
                Dim status = ExcelExportStyleRegistry.Resolve(If(sheetKey, sheetName), dr, table)
                If status <> ExcelStyleHelper.RowStatus.None Then
                    If isConnector Then
                        ' Connector: 행 전체가 아니라 "이슈 내용" 셀만 배경색/글자색 적용
                        ApplyConnectorIssueCellStyles(wb, row, table, dr, status)
                    Else
                        Dim style = ExcelStyleHelper.GetRowStyle(wb, status)
                        ExcelStyleHelper.ApplyStyleToRow(row, colCount, style)
                    End If
                End If

                If (r Mod 200) = 0 OrElse r = total - 1 Then
                    TryReportProgress(progressKey, r + 1, total, sheetName)
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

            ApplyHiddenColumns(sheet, table)
        End Sub

        Private Sub ApplyHiddenColumns(sheet As ISheet, table As DataTable)
            If sheet Is Nothing OrElse table Is Nothing Then Return

            For c As Integer = 0 To table.Columns.Count - 1
                Dim column As DataColumn = table.Columns(c)
                If column Is Nothing Then Continue For

                Dim shouldHide As Boolean = False
                Try
                    If column.ExtendedProperties.Contains(HiddenColumnFlag) Then
                        shouldHide = Convert.ToBoolean(column.ExtendedProperties(HiddenColumnFlag))
                    End If
                Catch
                    shouldHide = False
                End Try

                If Not shouldHide Then Continue For

                Try
                    sheet.SetColumnHidden(c, True)
                Catch
                End Try
            Next
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

        ' Connector: 행 전체가 아니라 "이슈 내용" 셀만 스타일 적용
        Private Sub ApplyConnectorIssueCellStyles(wb As IWorkbook, row As IRow, table As DataTable, dr As DataRow, status As ExcelStyleHelper.RowStatus)
            If wb Is Nothing OrElse row Is Nothing OrElse table Is Nothing OrElse dr Is Nothing Then Return

            Dim bg As ICellStyle = Nothing
            If status = ExcelStyleHelper.RowStatus.[Error] Then
                bg = ExcelStyleHelper.GetRowStyleNoWrap(wb, ExcelStyleHelper.RowStatus.[Error])
            Else
                bg = ExcelStyleHelper.GetRowStyleNoWrap(wb, ExcelStyleHelper.RowStatus.Warning)
            End If
            If bg Is Nothing Then Return

            Dim warnRed As ICellStyle = ExcelStyleHelper.GetWarningRedTextStyleNoWrap(wb)
            Dim errRed As ICellStyle = ExcelStyleHelper.GetErrorRedTextStyleNoWrap(wb)

            Dim idxParamCompare As Integer = FindFirstColumnIndex(table, "ParamCompare", "Param Compare")
            Dim idxConnType As Integer = FindFirstColumnIndex(table, "ConnectionType", "Connection Type")
            Dim idxDist As Integer = FindFirstColumnIndex(table, "Distance (inch)", "Distance (mm)", "Distance (inch or mm)", "Distance")
            Dim idxStatus As Integer = FindColumnIndex(table, "Status")
            Dim idxErrMsg As Integer = FindColumnIndex(table, "ErrorMessage")

            Dim pcText As String = GetColText(dr, table, idxParamCompare)
            Dim ctText As String = GetColText(dr, table, idxConnType)
            Dim stText As String = GetColText(dr, table, idxStatus)
            Dim emText As String = GetColText(dr, table, idxErrMsg)

            Dim isMismatch As Boolean = (Not String.IsNullOrWhiteSpace(pcText)) AndAlso
                                       (pcText.IndexOf("불일치", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        pcText.IndexOf("mismatch", StringComparison.OrdinalIgnoreCase) >= 0)

            Dim isProximity As Boolean = (Not String.IsNullOrWhiteSpace(ctText)) AndAlso
                                         (ctText.IndexOf("proximity", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                          ctText.IndexOf("disconnected", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                          ctText.IndexOf("미연결", StringComparison.OrdinalIgnoreCase) >= 0)

            Dim isConnError As Boolean = (Not String.IsNullOrWhiteSpace(ctText)) AndAlso
                                         (ctText.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)

            Dim isError As Boolean = (status = ExcelStyleHelper.RowStatus.[Error]) OrElse isConnError

            ' 1) ParamCompare: Mismatch 문구는 글자색 빨강 + (Warning/Error) 배경
            If idxParamCompare >= 0 AndAlso isMismatch Then
                ApplyStyleToCell(row, idxParamCompare, If(isError AndAlso errRed IsNot Nothing, errRed, warnRed))
            ElseIf idxParamCompare >= 0 AndAlso (Not String.IsNullOrWhiteSpace(pcText)) AndAlso status <> ExcelStyleHelper.RowStatus.None Then
                ' 기타 비교 관련 경고 문구가 있을 때만 배경색 적용
                ApplyStyleToCell(row, idxParamCompare, bg)
            End If

            ' 2) 연결 이슈(Proximity 등): ConnectionType / Distance 셀만 배경색 적용
            If isProximity Then
                If idxConnType >= 0 Then ApplyStyleToCell(row, idxConnType, bg)
                If idxDist >= 0 Then ApplyStyleToCell(row, idxDist, bg)
            End If

            ' 3) 명시적 에러: ConnectionType / ErrorMessage 셀만 배경색 적용
            If isError Then
                If idxConnType >= 0 AndAlso Not String.IsNullOrWhiteSpace(ctText) Then ApplyStyleToCell(row, idxConnType, bg)
                If idxErrMsg >= 0 AndAlso Not String.IsNullOrWhiteSpace(emText) Then ApplyStyleToCell(row, idxErrMsg, bg)
            End If

            ' 4) Status 셀이 존재하고 내용이 있으면(엑셀 스키마에 포함된 경우) 해당 셀만 배경색 적용
            If idxStatus >= 0 AndAlso Not String.IsNullOrWhiteSpace(stText) Then
                ApplyStyleToCell(row, idxStatus, bg)
            End If
        End Sub

        Private Function FindColumnIndex(table As DataTable, name As String) As Integer
            If table Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return -1
            For i As Integer = 0 To table.Columns.Count - 1
                If String.Equals(table.Columns(i).ColumnName, name, StringComparison.OrdinalIgnoreCase) Then Return i
            Next
            Return -1
        End Function

        Private Function FindFirstColumnIndex(table As DataTable, ParamArray names() As String) As Integer
            If table Is Nothing OrElse names Is Nothing Then Return -1
            For Each name As String In names
                Dim idx As Integer = FindColumnIndex(table, name)
                If idx >= 0 Then Return idx
            Next
            Return -1
        End Function

        Private Function GetColText(dr As DataRow, table As DataTable, idx As Integer) As String
            If dr Is Nothing OrElse table Is Nothing OrElse idx < 0 OrElse idx >= table.Columns.Count Then Return ""
            Try
                Dim v = dr(idx)
                If v Is Nothing OrElse v Is DBNull.Value Then Return ""
                Return v.ToString()
            Catch
                Return ""
            End Try
        End Function

        Private Sub ApplyStyleToCell(row As IRow, colIndex As Integer, style As ICellStyle)
            If row Is Nothing OrElse style Is Nothing OrElse colIndex < 0 Then Return
            Dim cell = row.GetCell(colIndex)
            If cell Is Nothing Then cell = row.CreateCell(colIndex)
            cell.CellStyle = style
        End Sub


        Private Function PickSavePath(filter As String, defaultFileName As String, title As String) As String
            Using dlg As New SaveFileDialog()
                dlg.Filter = filter
                dlg.Title = If(String.IsNullOrWhiteSpace(title), "저장", title)
                dlg.FileName = If(String.IsNullOrWhiteSpace(defaultFileName), "export.xlsx", defaultFileName)
                dlg.DefaultExt = "xlsx"
                dlg.AddExtension = True
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then Return ""
                Return EnsureExcelExtension(dlg.FileName)
            End Using
        End Function

        Private Function EnsureExcelExtension(filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return filePath

            Dim ext = Path.GetExtension(filePath)
            If String.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) Then Return filePath
            Return Path.ChangeExtension(filePath, ".xlsx")
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
                        Dim usedRange As Object = Nothing
                        Try
                            usedRange = ws.UsedRange
                            If usedRange IsNot Nothing Then
                                usedRange.EntireColumn.AutoFit()
                            Else
                                ws.Columns.AutoFit()
                            End If
                        Catch
                        Finally
                            ReleaseCom(usedRange)
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

                mi.Invoke(Nothing, New Object() {
                    progressKey,
                    "EXCEL_WRITE",
                    $"Writing {sheetName}",
                    Math.Max(0, current),
                    Math.Max(0, total),
                    Nothing,
                    False,
                    Nothing,
                    Nothing
                })
            Catch
            End Try
        End Sub

        Private Sub ReportExcelPhase(progressKey As String,
                                     phase As String,
                                     message As String,
                                     current As Integer,
                                     total As Integer,
                                     Optional percent As Double? = Nothing,
                                     Optional force As Boolean = False)
            If String.IsNullOrWhiteSpace(progressKey) Then Return
            Try
                Dim t = Type.GetType("KKY_Tool_Revit.UI.Hub.ExcelProgressReporter, " & GetType(ExcelCore).Assembly.FullName, throwOnError:=False)
                If t Is Nothing Then Return

                Dim mi = t.GetMethod("Report", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If mi Is Nothing Then Return

                mi.Invoke(Nothing, New Object() {
                    progressKey,
                    phase,
                    message,
                    Math.Max(0, current),
                    Math.Max(0, total),
                    percent,
                    force,
                    Nothing,
                    Nothing
                })
            Catch
            End Try
        End Sub

    End Module

End Namespace
