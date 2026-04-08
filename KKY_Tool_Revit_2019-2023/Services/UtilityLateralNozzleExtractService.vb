Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports KKY_Tool_Revit.Infrastructure
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel

Namespace Services

    Public NotInheritable Class UtilityLateralNozzleExtractService

        Private Const MaxHeaderColumnDistance As Integer = 20
        Private Shared ReadOnly _nozzleCodePattern As New Regex("_\d{3}$", RegexOptions.Compiled Or RegexOptions.CultureInvariant)

        Private Sub New()
        End Sub

        Public Class Settings
            Public Property ExcelPaths As List(Of String) = New List(Of String)()
            Public Property OutputFolder As String = String.Empty
        End Class

        Public Class RunSummary
            Public Property TotalFileCount As Integer
            Public Property SuccessCount As Integer
            Public Property FailCount As Integer
            Public Property NoDataCount As Integer
            Public Property ExtractedRowCount As Integer
            Public Property RemarkRowCount As Integer
        End Class

        Public Class FileRunResult
            Public Property FilePath As String = String.Empty
            Public Property FileName As String = String.Empty
            Public Property Status As String = String.Empty
            Public Property SheetCount As Integer
            Public Property ExtractedRowCount As Integer
            Public Property RemarkRowCount As Integer
            Public Property Message As String = String.Empty
        End Class

        Public Class ResultRow
            Public Property Utility As String = String.Empty
            Public Property LateralNo As String = String.Empty
            Public Property NozzleCode As String = String.Empty
            Public Property Remark As String = String.Empty
            Public Property SourceFile As String = String.Empty
            Public Property SourceSheet As String = String.Empty
        End Class

        Public Class LogEntry
            Public Property Level As String = String.Empty
            Public Property FilePath As String = String.Empty
            Public Property SheetName As String = String.Empty
            Public Property Message As String = String.Empty
        End Class

        Public Class RunResult
            Public Property Ok As Boolean
            Public Property Message As String = String.Empty
            Public Property OutputFolder As String = String.Empty
            Public Property ResultWorkbookPath As String = String.Empty
            Public Property Summary As RunSummary = New RunSummary()
            Public Property Files As List(Of FileRunResult) = New List(Of FileRunResult)()
            Public Property Rows As List(Of ResultRow) = New List(Of ResultRow)()
            Public Property Logs As List(Of LogEntry) = New List(Of LogEntry)()
        End Class

        Private Structure HeaderBlock
            Public Sub New(headerRow As Integer, utilityCol As Integer, lateralCol As Integer, nozzleCol As Integer)
                Me.HeaderRow = headerRow
                Me.UtilityCol = utilityCol
                Me.LateralCol = lateralCol
                Me.NozzleCol = nozzleCol
            End Sub

            Public Property HeaderRow As Integer
            Public Property UtilityCol As Integer
            Public Property LateralCol As Integer
            Public Property NozzleCol As Integer
        End Structure

        Public Shared Function Run(settings As Settings,
                                   progress As IProgress(Of Object)) As RunResult
            Dim result As New RunResult()
            Dim effectiveSettings = If(settings, New Settings())
            Dim excelPaths = effectiveSettings.ExcelPaths _
                .Where(Function(path) Not String.IsNullOrWhiteSpace(path)) _
                .Distinct(StringComparer.OrdinalIgnoreCase) _
                .ToList()

            If excelPaths.Count = 0 Then
                result.Message = "추출할 엑셀 파일을 1개 이상 선택해 주세요."
                Return result
            End If

            result.OutputFolder = ResolveOutputFolder(effectiveSettings)
            Directory.CreateDirectory(result.OutputFolder)

            ReportProgress(progress, 0, excelPaths.Count, "엑셀 파일 확인 중...", "")

            For i As Integer = 0 To excelPaths.Count - 1
                Dim path = excelPaths(i)
                ReportProgress(progress, i, excelPaths.Count, "엑셀 읽는 중...", IO.Path.GetFileName(path))
                result.Files.Add(ProcessWorkbook(path, result.Rows, result.Logs))
            Next

            result.Summary = BuildSummary(result.Files, result.Rows)
            result.Ok = result.Files.Any(Function(file) String.Equals(file.Status, "Success", StringComparison.OrdinalIgnoreCase) OrElse
                                                       String.Equals(file.Status, "NoData", StringComparison.OrdinalIgnoreCase))

            If result.Files.Count = 0 Then
                result.Message = "처리할 엑셀 파일이 없습니다."
            ElseIf result.Summary.SuccessCount = 0 AndAlso result.Summary.NoDataCount = 0 Then
                result.Message = "모든 파일 처리에 실패했습니다."
            ElseIf result.Summary.ExtractedRowCount = 0 Then
                result.Message = "추출 가능한 데이터가 없습니다."
            Else
                result.Message = $"추출 완료: {result.Summary.ExtractedRowCount}건"
            End If

            SaveArtifacts(result)
            ReportProgress(progress, excelPaths.Count, excelPaths.Count, "완료", result.Message)
            Return result
        End Function

        Private Shared Function ProcessWorkbook(path As String,
                                                rows As List(Of ResultRow),
                                                logs As List(Of LogEntry)) As FileRunResult
            Dim fileResult As New FileRunResult() With {
                .FilePath = If(path, String.Empty),
                .FileName = SafeFileName(path)
            }

            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                fileResult.Status = "Fail"
                fileResult.Message = "엑셀 파일을 찾을 수 없습니다."
                AddLog(logs, "FAIL", path, "", fileResult.Message)
                Return fileResult
            End If

            Try
                Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim workbook = OpenWorkbook(fs, IO.Path.GetExtension(path))
                    If workbook Is Nothing Then Throw New InvalidOperationException("엑셀 워크북을 열 수 없습니다.")

                    fileResult.SheetCount = workbook.NumberOfSheets
                    For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                        Dim sheet = workbook.GetSheetAt(sheetIndex)
                        If sheet Is Nothing Then Continue For
                        ProcessSheet(path, sheet, rows, logs, fileResult)
                    Next
                End Using

                If fileResult.ExtractedRowCount > 0 Then
                    fileResult.Status = "Success"
                    fileResult.Message = $"{fileResult.ExtractedRowCount}건 추출"
                Else
                    fileResult.Status = "NoData"
                    fileResult.Message = "추출 가능한 데이터가 없습니다."
                End If
            Catch ex As Exception
                fileResult.Status = "Fail"
                fileResult.Message = ex.Message
                AddLog(logs, "FAIL", path, "", $"{fileResult.FileName} 처리 실패: {ex.Message}")
            End Try

            Return fileResult
        End Function

        Private Shared Sub ProcessSheet(filePath As String,
                                        sheet As ISheet,
                                        rows As List(Of ResultRow),
                                        logs As List(Of LogEntry),
                                        fileResult As FileRunResult)
            If sheet Is Nothing Then Return

            Dim firstRow As Integer
            Dim firstCol As Integer
            Dim lastRow As Integer
            Dim lastCol As Integer
            If Not TryGetUsedBounds(sheet, firstRow, firstCol, lastRow, lastCol) Then Return

            Dim mergedChildren = BuildMergedChildLookup(sheet)
            Dim blocks As New List(Of HeaderBlock)()
            Dim headerRows As New HashSet(Of Integer)()
            FindHeaderBlocks(sheet, firstRow, firstCol, lastRow, lastCol, mergedChildren, blocks, headerRows)

            If blocks.Count = 0 Then
                AddLog(logs, "INFO", filePath, SafeSheetName(sheet), "헤더 블록을 찾지 못했습니다.")
                Return
            End If

            Dim headerRowArr = headerRows.OrderBy(Function(value) value).ToArray()
            Dim rowIndexMap As New Dictionary(Of Integer, Integer)()
            For i As Integer = 0 To headerRowArr.Length - 1
                rowIndexMap(headerRowArr(i)) = i
            Next

            For Each block In blocks
                Dim headerIndex = rowIndexMap(block.HeaderRow)
                Dim upperRow = If(headerIndex = 0, firstRow, headerRowArr(headerIndex - 1) + 1)
                Dim lowerRow = If(headerIndex = headerRowArr.Length - 1, lastRow, headerRowArr(headerIndex + 1) - 1)

                For rowIndex As Integer = upperRow To lowerRow
                    If rowIndex = block.HeaderRow Then Continue For

                    Dim utilityValue = GetDataCellValue(sheet, rowIndex, block.UtilityCol, mergedChildren)
                    Dim lateralValue = GetDataCellValue(sheet, rowIndex, block.LateralCol, mergedChildren)
                    Dim nozzleValue = GetDataCellValue(sheet, rowIndex, block.NozzleCol, mergedChildren)

                    If Not IsRealData(utilityValue, lateralValue, nozzleValue) Then Continue For

                    Dim remark = BuildRemark(utilityValue, lateralValue, nozzleValue)
                    rows.Add(New ResultRow() With {
                        .Utility = utilityValue,
                        .LateralNo = lateralValue,
                        .NozzleCode = nozzleValue,
                        .Remark = remark,
                        .SourceFile = SafeFileName(filePath),
                        .SourceSheet = SafeSheetName(sheet)
                    })

                    fileResult.ExtractedRowCount += 1
                    If Not String.IsNullOrWhiteSpace(remark) Then fileResult.RemarkRowCount += 1
                Next
            Next
        End Sub

        Private Shared Function OpenWorkbook(stream As Stream, extension As String) As IWorkbook
            If String.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) Then
                Return New HSSFWorkbook(stream)
            End If

            Return New XSSFWorkbook(stream)
        End Function

        Private Shared Function TryGetUsedBounds(sheet As ISheet,
                                                 ByRef firstRow As Integer,
                                                 ByRef firstCol As Integer,
                                                 ByRef lastRow As Integer,
                                                 ByRef lastCol As Integer) As Boolean
            firstRow = Integer.MaxValue
            firstCol = Integer.MaxValue
            lastRow = -1
            lastCol = -1

            If sheet Is Nothing Then Return False

            For rowIndex As Integer = sheet.FirstRowNum To sheet.LastRowNum
                Dim row = sheet.GetRow(rowIndex)
                If row Is Nothing Then Continue For
                If row.FirstCellNum < 0 Then Continue For

                For colIndex As Integer = row.FirstCellNum To row.LastCellNum - 1
                    If colIndex < 0 Then Continue For
                    Dim text = GetCellRawText(GetCell(sheet, rowIndex, colIndex))
                    If String.IsNullOrWhiteSpace(text) Then Continue For

                    firstRow = Math.Min(firstRow, rowIndex)
                    firstCol = Math.Min(firstCol, colIndex)
                    lastRow = Math.Max(lastRow, rowIndex)
                    lastCol = Math.Max(lastCol, colIndex)
                Next
            Next

            Return lastRow >= 0 AndAlso lastCol >= 0
        End Function

        Private Shared Sub FindHeaderBlocks(sheet As ISheet,
                                            firstRow As Integer,
                                            firstCol As Integer,
                                            lastRow As Integer,
                                            lastCol As Integer,
                                            mergedChildren As HashSet(Of String),
                                            blocks As List(Of HeaderBlock),
                                            headerRows As HashSet(Of Integer))
            For rowIndex As Integer = firstRow To lastRow
                Dim utilityCols As New List(Of Integer)()
                Dim lateralCols As New List(Of Integer)()
                Dim nozzleCols As New List(Of Integer)()

                For colIndex As Integer = firstCol To lastCol
                    Dim normalized = NormalizeText(GetCellDisplayText(sheet, rowIndex, colIndex, mergedChildren))
                    Select Case normalized
                        Case NormalizeText("UT명")
                            utilityCols.Add(colIndex)
                        Case NormalizeText("배관No")
                            lateralCols.Add(colIndex)
                        Case NormalizeText("연결호기")
                            nozzleCols.Add(colIndex)
                    End Select
                Next

                If utilityCols.Count = 0 OrElse lateralCols.Count = 0 OrElse nozzleCols.Count = 0 Then Continue For

                Dim usedLateral As New HashSet(Of Integer)()
                Dim usedNozzle As New HashSet(Of Integer)()
                For Each utilityCol In utilityCols
                    Dim lateralCol = FindNearestUnusedColumn(lateralCols, usedLateral, utilityCol, MaxHeaderColumnDistance)
                    Dim nozzleCol = FindNearestUnusedColumn(nozzleCols, usedNozzle, utilityCol, MaxHeaderColumnDistance)
                    If lateralCol < 0 OrElse nozzleCol < 0 Then Continue For

                    blocks.Add(New HeaderBlock(rowIndex, utilityCol, lateralCol, nozzleCol))
                    headerRows.Add(rowIndex)
                Next
            Next
        End Sub

        Private Shared Function FindNearestUnusedColumn(columns As IEnumerable(Of Integer),
                                                        usedSet As HashSet(Of Integer),
                                                        baseCol As Integer,
                                                        maxDistance As Integer) As Integer
            Dim bestCol As Integer = -1
            Dim bestDistance As Integer = Integer.MaxValue

            For Each col In columns
                If usedSet.Contains(col) Then Continue For

                Dim distance = Math.Abs(col - baseCol)
                If distance > maxDistance Then Continue For
                If distance >= bestDistance Then Continue For

                bestDistance = distance
                bestCol = col
            Next

            If bestCol >= 0 Then usedSet.Add(bestCol)
            Return bestCol
        End Function

        Private Shared Function GetDataCellValue(sheet As ISheet,
                                                 rowIndex As Integer,
                                                 colIndex As Integer,
                                                 mergedChildren As HashSet(Of String)) As String
            Dim text = CleanOutputText(GetCellDisplayText(sheet, rowIndex, colIndex, mergedChildren))
            If IsHeaderOrNoise(text) Then Return String.Empty
            Return text
        End Function

        Private Shared Function GetCellDisplayText(sheet As ISheet,
                                                   rowIndex As Integer,
                                                   colIndex As Integer,
                                                   mergedChildren As HashSet(Of String)) As String
            If mergedChildren IsNot Nothing AndAlso mergedChildren.Contains(MakeCellKey(rowIndex, colIndex)) Then
                Return String.Empty
            End If

            Return GetCellRawText(GetCell(sheet, rowIndex, colIndex))
        End Function

        Private Shared Function GetCellRawText(cell As ICell) As String
            If cell Is Nothing Then Return String.Empty

            Select Case cell.CellType
                Case CellType.String
                    Return If(cell.StringCellValue, String.Empty)
                Case CellType.Numeric
                    If DateUtil.IsCellDateFormatted(cell) Then
                        Return cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    End If
                    Return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture)
                Case CellType.Boolean
                    Return If(cell.BooleanCellValue, "TRUE", "FALSE")
                Case CellType.Formula
                    Return GetFormulaResultText(cell)
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function GetFormulaResultText(cell As ICell) As String
            If cell Is Nothing Then Return String.Empty

            Select Case cell.CachedFormulaResultType
                Case CellType.String
                    Return If(cell.StringCellValue, String.Empty)
                Case CellType.Numeric
                    If DateUtil.IsCellDateFormatted(cell) Then
                        Return cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    End If
                    Return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture)
                Case CellType.Boolean
                    Return If(cell.BooleanCellValue, "TRUE", "FALSE")
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function BuildMergedChildLookup(sheet As ISheet) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.Ordinal)
            If sheet Is Nothing Then Return result

            For i As Integer = 0 To sheet.NumMergedRegions - 1
                Dim region = sheet.GetMergedRegion(i)
                If region Is Nothing Then Continue For

                For rowIndex As Integer = region.FirstRow To region.LastRow
                    For colIndex As Integer = region.FirstColumn To region.LastColumn
                        If rowIndex = region.FirstRow AndAlso colIndex = region.FirstColumn Then Continue For
                        result.Add(MakeCellKey(rowIndex, colIndex))
                    Next
                Next
            Next

            Return result
        End Function

        Private Shared Function BuildRemark(utilityValue As String,
                                            lateralValue As String,
                                            nozzleValue As String) As String
            Dim remarks As New List(Of String)()

            If String.IsNullOrWhiteSpace(utilityValue) Then remarks.Add("UTILITY 누락")
            If String.IsNullOrWhiteSpace(lateralValue) Then remarks.Add("LATERAL NO 누락")
            If String.IsNullOrWhiteSpace(nozzleValue) Then remarks.Add("Nozzle Code 누락")

            Dim trimmedNozzle = If(nozzleValue, String.Empty).Trim()
            If trimmedNozzle.Length > 0 AndAlso Not _nozzleCodePattern.IsMatch(trimmedNozzle) Then
                remarks.Add("Nozzle Code 형식 불일치")
            End If

            Return String.Join(" / ", remarks)
        End Function

        Private Shared Function IsRealData(utilityValue As String,
                                           lateralValue As String,
                                           nozzleValue As String) As Boolean
            Return If(utilityValue, String.Empty).Trim().Length > 0 OrElse
                   If(lateralValue, String.Empty).Trim().Length > 0 OrElse
                   If(nozzleValue, String.Empty).Trim().Length > 0
        End Function

        Private Shared Function IsHeaderOrNoise(value As String) As Boolean
            Dim normalized = NormalizeText(value)
            If normalized = String.Empty Then Return True
            If normalized = NormalizeText("UT명") Then Return True
            If normalized = NormalizeText("배관No") Then Return True
            If normalized = NormalizeText("연결호기") Then Return True
            If normalized.Contains("MAINSIZE") Then Return True
            Return False
        End Function

        Private Shared Function NormalizeText(value As String) As String
            Dim text = If(value, String.Empty)
            text = text.Replace(vbCr, String.Empty)
            text = text.Replace(vbLf, String.Empty)
            text = text.Replace(vbTab, String.Empty)
            text = text.Replace(ChrW(160), String.Empty)
            text = text.Replace("　", String.Empty)
            text = text.Replace(" ", String.Empty)
            Return text.Trim().ToUpperInvariant()
        End Function

        Private Shared Function CleanOutputText(value As String) As String
            Dim text = If(value, String.Empty)
            text = text.Replace(vbCrLf, " ")
            text = text.Replace(vbCr, " ")
            text = text.Replace(vbLf, " ")
            text = text.Trim()

            Do While text.Contains("  ")
                text = text.Replace("  ", " ")
            Loop

            Return text
        End Function

        Private Shared Function ResolveOutputFolder(settings As Settings) As String
            If settings IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(settings.OutputFolder) Then
                Return settings.OutputFolder.Trim()
            End If

            Return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End Function

        Private Shared Sub SaveArtifacts(result As RunResult)
            If result Is Nothing Then Return
            If String.IsNullOrWhiteSpace(result.OutputFolder) Then
                result.OutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            End If

            Directory.CreateDirectory(result.OutputFolder)
            Dim timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            Dim workbookPath = IO.Path.Combine(result.OutputFolder, $"노즐코드_KTA_단일화_{timeStamp}.xlsx")

            Dim sheets As New List(Of KeyValuePair(Of String, DataTable)) From {
                New KeyValuePair(Of String, DataTable)("추출결과", BuildResultTable(result.Rows)),
                New KeyValuePair(Of String, DataTable)("요약", BuildSummaryTable(result.Files, result.Summary)),
                New KeyValuePair(Of String, DataTable)("Logs", BuildLogTable(result.Logs))
            }

            ExcelCore.SaveXlsxMulti(workbookPath, sheets, autoFit:=True)
            result.ResultWorkbookPath = workbookPath
        End Sub

        Private Shared Function BuildResultTable(rows As IEnumerable(Of ResultRow)) As DataTable
            Dim table As New DataTable("추출결과")
            table.Columns.Add("UTILITY", GetType(String))
            table.Columns.Add("LATERAL NO", GetType(String))
            table.Columns.Add("Nozzle Code", GetType(String))
            table.Columns.Add("비고", GetType(String))

            For Each rowItem In If(rows, Enumerable.Empty(Of ResultRow)())
                table.Rows.Add(rowItem.Utility, rowItem.LateralNo, rowItem.NozzleCode, rowItem.Remark)
            Next

            Return table
        End Function

        Private Shared Function BuildSummaryTable(files As IEnumerable(Of FileRunResult),
                                                  summary As RunSummary) As DataTable
            Dim table As New DataTable("요약")
            table.Columns.Add("FileName", GetType(String))
            table.Columns.Add("Status", GetType(String))
            table.Columns.Add("SheetCount", GetType(Integer))
            table.Columns.Add("ExtractedRows", GetType(Integer))
            table.Columns.Add("RemarkRows", GetType(Integer))
            table.Columns.Add("Message", GetType(String))

            For Each fileResult In If(files, Enumerable.Empty(Of FileRunResult)())
                table.Rows.Add(fileResult.FileName, fileResult.Status, fileResult.SheetCount, fileResult.ExtractedRowCount, fileResult.RemarkRowCount, fileResult.Message)
            Next

            table.Rows.Add("TOTAL", "Summary", summary.TotalFileCount, summary.ExtractedRowCount, summary.RemarkRowCount,
                           $"성공 {summary.SuccessCount} / 실패 {summary.FailCount} / 추출없음 {summary.NoDataCount}")

            Return table
        End Function

        Private Shared Function BuildLogTable(logs As IEnumerable(Of LogEntry)) As DataTable
            Dim table As New DataTable("Logs")
            table.Columns.Add("Level", GetType(String))
            table.Columns.Add("FilePath", GetType(String))
            table.Columns.Add("SheetName", GetType(String))
            table.Columns.Add("Message", GetType(String))

            For Each logEntry In If(logs, Enumerable.Empty(Of LogEntry)())
                table.Rows.Add(logEntry.Level, logEntry.FilePath, logEntry.SheetName, logEntry.Message)
            Next

            Return table
        End Function

        Private Shared Function BuildSummary(files As IEnumerable(Of FileRunResult),
                                             rows As IEnumerable(Of ResultRow)) As RunSummary
            Dim fileList = If(files, Enumerable.Empty(Of FileRunResult)()).ToList()
            Dim rowList = If(rows, Enumerable.Empty(Of ResultRow)()).ToList()

            Return New RunSummary() With {
                .TotalFileCount = fileList.Count,
                .SuccessCount = fileList.Where(Function(file) String.Equals(file.Status, "Success", StringComparison.OrdinalIgnoreCase)).Count(),
                .FailCount = fileList.Where(Function(file) String.Equals(file.Status, "Fail", StringComparison.OrdinalIgnoreCase)).Count(),
                .NoDataCount = fileList.Where(Function(file) String.Equals(file.Status, "NoData", StringComparison.OrdinalIgnoreCase)).Count(),
                .ExtractedRowCount = rowList.Count,
                .RemarkRowCount = rowList.Where(Function(row) Not String.IsNullOrWhiteSpace(row.Remark)).Count()
            }
        End Function

        Private Shared Sub AddLog(logs As List(Of LogEntry),
                                  level As String,
                                  filePath As String,
                                  sheetName As String,
                                  message As String)
            If logs Is Nothing OrElse String.IsNullOrWhiteSpace(message) Then Return

            logs.Add(New LogEntry() With {
                .Level = If(level, "INFO"),
                .FilePath = If(filePath, String.Empty),
                .SheetName = If(sheetName, String.Empty),
                .Message = message
            })
        End Sub

        Private Shared Sub ReportProgress(progress As IProgress(Of Object),
                                          current As Integer,
                                          total As Integer,
                                          message As String,
                                          detail As String)
            If progress Is Nothing Then Return

            Dim percent As Double = 0.0R
            If total > 0 Then
                percent = Math.Max(0.0R, Math.Min(100.0R, (CDbl(current) / CDbl(total)) * 100.0R))
            End If

            progress.Report(New With {
                .title = "노즐코드 KTA 단일화",
                .message = If(message, String.Empty),
                .detail = If(detail, String.Empty),
                .current = current,
                .total = total,
                .percent = percent
            })
        End Sub

        Private Shared Function SafeFileName(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then Return String.Empty
            Try
                Return IO.Path.GetFileName(path)
            Catch
                Return path
            End Try
        End Function

        Private Shared Function SafeSheetName(sheet As ISheet) As String
            If sheet Is Nothing Then Return String.Empty
            Try
                Return If(sheet.SheetName, String.Empty)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function GetCell(sheet As ISheet, rowIndex As Integer, colIndex As Integer) As ICell
            If sheet Is Nothing OrElse rowIndex < 0 OrElse colIndex < 0 Then Return Nothing

            Dim row = sheet.GetRow(rowIndex)
            If row Is Nothing Then Return Nothing
            Return row.GetCell(colIndex)
        End Function

        Private Shared Function MakeCellKey(rowIndex As Integer, colIndex As Integer) As String
            Return rowIndex.ToString(CultureInfo.InvariantCulture) & ":" & colIndex.ToString(CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
