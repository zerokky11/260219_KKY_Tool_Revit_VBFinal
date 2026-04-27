Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Friend NotInheritable Class ReviewExportMessageOverrideService

        Private Const OverrideWorkbookFileName As String = "review_export_message_overrides.xlsx"
        Private Const OverrideJsonFileName As String = "review_export_message_overrides.json"
        Private Const MatrixWorkbookFileName As String = "review_export_feature_sheet_matrix.xlsx"
        Private Const MatrixIndexSheetName As String = "Index"
        Private Shared ReadOnly SyncRoot As New Object()
        Private Shared ReadOnly EmptyEntries As New Dictionary(Of String, List(Of OverrideEntry))(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly EmptyHeaderEntries As New Dictionary(Of String, HeaderOverrideEntry)(StringComparer.OrdinalIgnoreCase)
        Private Const SourceSnapshotTtlTicks As Long = 10000000L

        Private Shared _cachedEntriesSignature As String = String.Empty
        Private Shared _cachedEntries As Dictionary(Of String, List(Of OverrideEntry)) = EmptyEntries
        Private Shared _cachedHeaderSignature As String = String.Empty
        Private Shared _cachedHeaderEntries As Dictionary(Of String, HeaderOverrideEntry) = EmptyHeaderEntries
        Private Shared _cachedSourceSnapshot As OverrideSourceSnapshot = Nothing

        Private Sub New()
        End Sub

        Friend NotInheritable Class OverrideDocument
            Public Property Entries As List(Of OverrideEntry)
        End Class

        Friend NotInheritable Class OverrideEntry
            Public Property ExportKey As String
            Public Property OriginalHeader As String
            Public Property ExportHeader As String
            Public Property MatchTexts As List(Of String)
            Public Property KoText As String
            Public Property EnText As String
        End Class

        Private NotInheritable Class HeaderOverrideEntry
            Public Property KoHeader As String
            Public Property EnHeader As String
            Public Property KoChanged As Boolean
            Public Property EnChanged As Boolean
        End Class

        Private NotInheritable Class MatrixSheetBinding
            Public Property SheetName As String
            Public Property Feature As String
            Public Property Section As String
            Public Property Locale As String
        End Class

        Private NotInheritable Class WorkbookSourceRow
            Public Property RowIndex As Integer
            Public Property Section As String
            Public Property OriginalHeader As String
            Public Property Condition As String
        End Class

        Private NotInheritable Class OverrideSourceSnapshot
            Public Property WorkbookPath As String
            Public Property MatrixPath As String
            Public Property JsonPath As String
            Public Property EntriesSignature As String
            Public Property HeaderSignature As String
            Public Property ExpiresAtUtcTicks As Long
        End Class

        Public Shared Function TryResolve(exportKey As String,
                                          originalHeader As String,
                                          exportHeader As String,
                                          sourceText As String,
                                          translatedText As String,
                                          locale As String,
                                          ByRef resolvedText As String) As Boolean
            resolvedText = Nothing

            Dim entriesByKey = GetEntries()
            If entriesByKey Is Nothing OrElse entriesByKey.Count = 0 Then Return False

            Dim candidates As New List(Of String)()
            AddCandidate(candidates, sourceText)
            AddCandidate(candidates, translatedText)
            If candidates.Count = 0 Then Return False

            Dim normalizedExportKey As String = NormalizeExportKey(exportKey)
            Dim exactKey As String = BuildKey(normalizedExportKey, originalHeader, exportHeader)
            Dim fallbackKey As String = BuildKey(normalizedExportKey, "", exportHeader)

            If TryResolveFromBucket(entriesByKey, exactKey, candidates, locale, resolvedText) Then Return True
            If String.Equals(exactKey, fallbackKey, StringComparison.OrdinalIgnoreCase) Then Return False

            Return TryResolveFromBucket(entriesByKey, fallbackKey, candidates, locale, resolvedText)
        End Function

        Public Shared Function ResolveHeader(exportKey As String,
                                             originalHeader As String,
                                             defaultExportHeader As String,
                                             locale As String) As String
            Dim entriesByKey = GetHeaderEntries()
            If entriesByKey Is Nothing OrElse entriesByKey.Count = 0 Then Return defaultExportHeader

            Dim normalizedExportKey As String = NormalizeExportKey(exportKey)
            Dim exactKey As String = BuildHeaderKey(normalizedExportKey, originalHeader, defaultExportHeader)
            Dim fallbackKey As String = BuildHeaderKey(normalizedExportKey, originalHeader, "")

            Dim entry As HeaderOverrideEntry = Nothing
            If Not entriesByKey.TryGetValue(exactKey, entry) Then
                entriesByKey.TryGetValue(fallbackKey, entry)
            End If

            If entry Is Nothing Then Return defaultExportHeader

            Dim localizedHeader As String = If(NormalizeLocale(locale) = "en", entry.EnHeader, entry.KoHeader)
            If String.IsNullOrWhiteSpace(localizedHeader) Then Return defaultExportHeader

            Return localizedHeader
        End Function

        Private Shared Function TryResolveFromBucket(entriesByKey As Dictionary(Of String, List(Of OverrideEntry)),
                                                     bucketKey As String,
                                                     candidates As IList(Of String),
                                                     locale As String,
                                                     ByRef resolvedText As String) As Boolean
            Dim bucket As List(Of OverrideEntry) = Nothing
            If Not entriesByKey.TryGetValue(bucketKey, bucket) OrElse bucket Is Nothing OrElse bucket.Count = 0 Then
                Return False
            End If

            For Each entry In bucket
                If entry Is Nothing Then Continue For
                If Not MatchesAny(entry.MatchTexts, candidates) Then Continue For

                Dim localizedText As String = If(NormalizeLocale(locale) = "en", entry.EnText, entry.KoText)
                If localizedText Is Nothing Then Continue For

                resolvedText = ApplyTemplateCaptures(localizedText, entry.MatchTexts, candidates)
                Return True
            Next

            Return False
        End Function

        Private Shared Function GetEntries() As Dictionary(Of String, List(Of OverrideEntry))
            Dim sourceSnapshot As OverrideSourceSnapshot = GetSourceSnapshot()
            Dim workbookPath As String = sourceSnapshot.WorkbookPath
            Dim matrixPath As String = sourceSnapshot.MatrixPath
            Dim jsonPath As String = sourceSnapshot.JsonPath
            Dim cacheSignature As String = sourceSnapshot.EntriesSignature

            If cacheSignature = String.Empty Then Return EmptyEntries

            SyncLock SyncRoot
                If String.Equals(cacheSignature, _cachedEntriesSignature, StringComparison.OrdinalIgnoreCase) Then
                    Return _cachedEntries
                End If
            End SyncLock

            Dim loadedEntries As Dictionary(Of String, List(Of OverrideEntry)) = EmptyEntries
            If Not String.IsNullOrWhiteSpace(workbookPath) AndAlso Not String.IsNullOrWhiteSpace(matrixPath) Then
                loadedEntries = LoadEntriesFromMatrixWorkbook(workbookPath, matrixPath)
            ElseIf Not String.IsNullOrWhiteSpace(workbookPath) Then
                loadedEntries = LoadEntriesFromWorkbook(workbookPath)
            ElseIf Not String.IsNullOrWhiteSpace(jsonPath) Then
                loadedEntries = LoadEntriesFromJson(jsonPath)
            End If

            SyncLock SyncRoot
                _cachedEntriesSignature = cacheSignature
                _cachedEntries = loadedEntries
                Return _cachedEntries
            End SyncLock
        End Function

        Private Shared Function GetHeaderEntries() As Dictionary(Of String, HeaderOverrideEntry)
            Dim sourceSnapshot As OverrideSourceSnapshot = GetSourceSnapshot()
            Dim workbookPath As String = sourceSnapshot.WorkbookPath
            If String.IsNullOrWhiteSpace(workbookPath) Then Return EmptyHeaderEntries

            Dim matrixPath As String = sourceSnapshot.MatrixPath
            Dim cacheSignature As String = sourceSnapshot.HeaderSignature

            SyncLock SyncRoot
                If String.Equals(cacheSignature, _cachedHeaderSignature, StringComparison.OrdinalIgnoreCase) Then
                    Return _cachedHeaderEntries
                End If
            End SyncLock

            Dim loadedEntries As Dictionary(Of String, HeaderOverrideEntry) =
                If(String.IsNullOrWhiteSpace(matrixPath),
                   LoadHeadersFromWorkbook(workbookPath),
                   LoadHeadersFromMatrixWorkbook(workbookPath, matrixPath))

            SyncLock SyncRoot
                _cachedHeaderSignature = cacheSignature
                _cachedHeaderEntries = loadedEntries
                Return _cachedHeaderEntries
            End SyncLock
        End Function

        Private Shared Function GetSourceSnapshot() As OverrideSourceSnapshot
            Dim nowTicks As Long = DateTime.UtcNow.Ticks

            SyncLock SyncRoot
                If _cachedSourceSnapshot IsNot Nothing AndAlso nowTicks <= _cachedSourceSnapshot.ExpiresAtUtcTicks Then
                    Return _cachedSourceSnapshot
                End If
            End SyncLock

            Dim loadedSnapshot As OverrideSourceSnapshot = CreateSourceSnapshot(nowTicks + SourceSnapshotTtlTicks)

            SyncLock SyncRoot
                If _cachedSourceSnapshot IsNot Nothing AndAlso nowTicks <= _cachedSourceSnapshot.ExpiresAtUtcTicks Then
                    Return _cachedSourceSnapshot
                End If

                _cachedSourceSnapshot = loadedSnapshot
                Return _cachedSourceSnapshot
            End SyncLock
        End Function

        Private Shared Function CreateSourceSnapshot(expiresAtUtcTicks As Long) As OverrideSourceSnapshot
            Dim workbookPath As String = ResolveOverrideWorkbookPath()
            Dim matrixPath As String = ResolveMatrixWorkbookPath()
            Dim jsonPath As String = ResolveOverrideJsonPath()

            Dim entriesSignature As String = String.Empty
            Dim headerSignature As String = String.Empty

            If Not String.IsNullOrWhiteSpace(workbookPath) AndAlso Not String.IsNullOrWhiteSpace(matrixPath) Then
                entriesSignature = BuildCacheSignature("matrix", workbookPath, matrixPath)
                headerSignature = entriesSignature
            ElseIf Not String.IsNullOrWhiteSpace(workbookPath) Then
                entriesSignature = BuildCacheSignature("workbook", workbookPath)
                headerSignature = entriesSignature
            ElseIf Not String.IsNullOrWhiteSpace(jsonPath) Then
                entriesSignature = BuildCacheSignature("json", jsonPath)
            End If

            Return New OverrideSourceSnapshot With {
                .WorkbookPath = workbookPath,
                .MatrixPath = matrixPath,
                .JsonPath = jsonPath,
                .EntriesSignature = entriesSignature,
                .HeaderSignature = headerSignature,
                .ExpiresAtUtcTicks = expiresAtUtcTicks
            }
        End Function

        Private Shared Function LoadEntriesFromJson(path As String) As Dictionary(Of String, List(Of OverrideEntry))
            Dim result As New Dictionary(Of String, List(Of OverrideEntry))(StringComparer.OrdinalIgnoreCase)

            Try
                Dim rawJson As String = File.ReadAllText(path, Encoding.UTF8)
                If String.IsNullOrWhiteSpace(rawJson) Then Return result

                Dim serializer As New JavaScriptSerializer()
                serializer.MaxJsonLength = Integer.MaxValue

                Dim doc = serializer.Deserialize(Of OverrideDocument)(rawJson)
                If doc Is Nothing OrElse doc.Entries Is Nothing Then Return result

                For Each entry In doc.Entries
                    If entry Is Nothing Then Continue For

                    Dim exportKey As String = NormalizeExportKey(entry.ExportKey)
                    Dim originalHeader As String = NormalizeToken(entry.OriginalHeader)
                    Dim exportHeader As String = NormalizeToken(entry.ExportHeader)
                    Dim matchTexts As List(Of String) = NormalizeTexts(entry.MatchTexts)

                    If exportKey = String.Empty OrElse exportHeader = String.Empty OrElse matchTexts.Count = 0 Then
                        Continue For
                    End If

                    Dim normalizedKoText As String = NormalizeOverrideOutput(entry.KoText)
                    Dim normalizedEnText As String = NormalizeOverrideOutput(entry.EnText)
                    If normalizedKoText Is Nothing AndAlso normalizedEnText Is Nothing Then
                        Continue For
                    End If

                    Dim normalizedEntry As New OverrideEntry With {
                        .ExportKey = exportKey,
                        .OriginalHeader = originalHeader,
                        .ExportHeader = exportHeader,
                        .MatchTexts = matchTexts,
                        .KoText = normalizedKoText,
                        .EnText = normalizedEnText
                    }

                    AddEntry(result, BuildKey(exportKey, originalHeader, exportHeader), normalizedEntry)
                    If originalHeader <> String.Empty Then
                        AddEntry(result, BuildKey(exportKey, "", exportHeader), normalizedEntry)
                    End If
                Next
            Catch
                Return EmptyEntries
            End Try

            Return result
        End Function

        Private Shared Function LoadEntriesFromMatrixWorkbook(workbookPath As String,
                                                             matrixPath As String) As Dictionary(Of String, List(Of OverrideEntry))
            Try
                Using fs As New FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim workbook As IWorkbook = New XSSFWorkbook(fs)
                    ApplyMatrixWorkbookOverrides(workbook, matrixPath)
                    Return LoadEntriesFromWorkbook(workbook)
                End Using
            Catch
                Return LoadEntriesFromWorkbook(workbookPath)
            End Try
        End Function

        Private Shared Function LoadEntriesFromWorkbook(path As String) As Dictionary(Of String, List(Of OverrideEntry))
            Try
                Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim workbook As IWorkbook = New XSSFWorkbook(fs)
                    Return LoadEntriesFromWorkbook(workbook)
                End Using
            Catch
                Return EmptyEntries
            End Try
        End Function

        Private Shared Function LoadEntriesFromWorkbook(workbook As IWorkbook) As Dictionary(Of String, List(Of OverrideEntry))
            Dim result As New Dictionary(Of String, List(Of OverrideEntry))(StringComparer.OrdinalIgnoreCase)
            If workbook Is Nothing Then Return result

            Try
                Dim formatter As New DataFormatter()

                For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                    Dim sheet As ISheet = workbook.GetSheetAt(sheetIndex)
                    If sheet Is Nothing Then Continue For
                    If ShouldSkipOverrideSheet(sheet.SheetName) Then Continue For

                    Dim headerRow As IRow = sheet.GetRow(sheet.FirstRowNum)
                    If headerRow Is Nothing Then Continue For

                    Dim headerMap As Dictionary(Of String, Integer) = BuildSheetHeaderMap(headerRow, formatter)
                    If Not HasWorkbookOverrideHeaders(headerMap) Then Continue For

                    Dim featureFromSheet As String = NormalizeFeatureName(sheet.SheetName)

                    For rowIndex As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                        Dim row As IRow = sheet.GetRow(rowIndex)
                        If row Is Nothing Then Continue For

                        Dim koOverride As String = ReadCellText(row, headerMap, formatter, "한글 override", "한글 적용문구", "한국어 override", "한국어 적용문구")
                        Dim enOverride As String = ReadCellText(row, headerMap, formatter, "영문 override", "영문 적용문구", "English override", "English Applied Text")
                        koOverride = NormalizeOverrideOutput(koOverride)
                        enOverride = NormalizeOverrideOutput(enOverride)
                        If koOverride Is Nothing AndAlso enOverride Is Nothing Then Continue For

                        Dim feature As String = ReadCellText(row, headerMap, formatter, "기능")
                        If feature = String.Empty Then feature = featureFromSheet

                        Dim exportKey As String = NormalizeExportKey(feature)
                        Dim originalHeader As String = ReadCellText(row, headerMap, formatter, "원본/내부 열명")
                        Dim exportHeader As String = ReadCellText(row, headerMap, formatter, "영문 export 열명", "한글 export 열명")
                        If exportHeader = String.Empty Then exportHeader = originalHeader

                        Dim rawText As String = ReadCellText(row, headerMap, formatter, "원본(raw)")
                        Dim koText As String = ReadCellText(row, headerMap, formatter, "한글 출력")
                        Dim enText As String = ReadCellText(row, headerMap, formatter, "영문 출력")
                        Dim matchTexts As List(Of String) = NormalizeTexts(New String() {rawText, koText, enText})

                        If exportKey = String.Empty OrElse exportHeader = String.Empty OrElse matchTexts.Count = 0 Then
                            Continue For
                        End If

                        Dim entry As New OverrideEntry With {
                            .ExportKey = exportKey,
                            .OriginalHeader = originalHeader,
                            .ExportHeader = exportHeader,
                            .MatchTexts = matchTexts,
                            .KoText = koOverride,
                            .EnText = enOverride
                        }

                        AddEntry(result, BuildKey(exportKey, originalHeader, exportHeader), entry)
                        If originalHeader <> String.Empty Then
                            AddEntry(result, BuildKey(exportKey, "", exportHeader), entry)
                        End If
                    Next
                Next
            Catch
                Return EmptyEntries
            End Try

            Return result
        End Function

        Private Shared Function LoadHeadersFromMatrixWorkbook(workbookPath As String,
                                                              matrixPath As String) As Dictionary(Of String, HeaderOverrideEntry)
            Try
                Using fs As New FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim workbook As IWorkbook = New XSSFWorkbook(fs)
                    ApplyMatrixWorkbookOverrides(workbook, matrixPath)
                    Return LoadHeadersFromWorkbook(workbook)
                End Using
            Catch
                Return LoadHeadersFromWorkbook(workbookPath)
            End Try
        End Function

        Private Shared Function LoadHeadersFromWorkbook(path As String) As Dictionary(Of String, HeaderOverrideEntry)
            Try
                Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim workbook As IWorkbook = New XSSFWorkbook(fs)
                    Return LoadHeadersFromWorkbook(workbook)
                End Using
            Catch
                Return EmptyHeaderEntries
            End Try
        End Function

        Private Shared Function LoadHeadersFromWorkbook(workbook As IWorkbook) As Dictionary(Of String, HeaderOverrideEntry)
            Dim result As New Dictionary(Of String, HeaderOverrideEntry)(StringComparer.OrdinalIgnoreCase)
            If workbook Is Nothing Then Return result

            Try
                Dim formatter As New DataFormatter()

                For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                    Dim sheet As ISheet = workbook.GetSheetAt(sheetIndex)
                    If sheet Is Nothing Then Continue For
                    If ShouldSkipOverrideSheet(sheet.SheetName) Then Continue For

                    Dim headerRow As IRow = sheet.GetRow(sheet.FirstRowNum)
                    If headerRow Is Nothing Then Continue For

                    Dim headerMap As Dictionary(Of String, Integer) = BuildSheetHeaderMap(headerRow, formatter)
                    If Not HasWorkbookHeaderOverrideHeaders(headerMap) Then Continue For

                    Dim featureFromSheet As String = NormalizeFeatureName(sheet.SheetName)

                    For rowIndex As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                        Dim row As IRow = sheet.GetRow(rowIndex)
                        If row Is Nothing Then Continue For

                        Dim feature As String = ReadCellText(row, headerMap, formatter, "기능")
                        If feature = String.Empty Then feature = featureFromSheet

                        Dim exportKey As String = NormalizeExportKey(feature)
                        Dim originalHeader As String = ReadCellText(row, headerMap, formatter, "원본/내부 열명")
                        Dim koDefaultHeader As String = ReadCellText(row, headerMap, formatter, "한글 export 열명")
                        If koDefaultHeader = String.Empty Then koDefaultHeader = originalHeader

                        Dim enDefaultHeader As String = ReadCellText(row, headerMap, formatter, "영문 export 열명")
                        Dim defaultExportHeader As String = If(enDefaultHeader <> String.Empty, enDefaultHeader, koDefaultHeader)
                        If defaultExportHeader = String.Empty Then defaultExportHeader = originalHeader

                        If exportKey = String.Empty OrElse originalHeader = String.Empty OrElse defaultExportHeader = String.Empty Then
                            Continue For
                        End If

                        Dim koOverrideHeader As String = ReadCellText(row, headerMap, formatter, "한글 헤더 override")
                        koOverrideHeader = NormalizeHeaderOverride(koOverrideHeader)
                        If koOverrideHeader = String.Empty Then koOverrideHeader = koDefaultHeader

                        Dim enOverrideHeader As String = ReadCellText(row, headerMap, formatter, "영문 헤더 override")
                        enOverrideHeader = NormalizeHeaderOverride(enOverrideHeader)
                        If enOverrideHeader = String.Empty Then enOverrideHeader = enDefaultHeader

                        Dim exactKey As String = BuildHeaderKey(exportKey, originalHeader, defaultExportHeader)
                        Dim fallbackKey As String = BuildHeaderKey(exportKey, originalHeader, "")

                        MergeHeaderEntry(result, exactKey, koDefaultHeader, enDefaultHeader, koOverrideHeader, enOverrideHeader)
                        MergeHeaderEntry(result, fallbackKey, koDefaultHeader, enDefaultHeader, koOverrideHeader, enOverrideHeader)
                    Next
                Next
            Catch
                Return EmptyHeaderEntries
            End Try

            Return result
        End Function

        Private Shared Sub ApplyMatrixWorkbookOverrides(overrideWorkbook As IWorkbook,
                                                        matrixPath As String)
            If overrideWorkbook Is Nothing OrElse String.IsNullOrWhiteSpace(matrixPath) Then Return

            Using fs As New FileStream(matrixPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim matrixWorkbook As IWorkbook = New XSSFWorkbook(fs)
                Dim formatter As New DataFormatter()
                Dim bindings As List(Of MatrixSheetBinding) = ReadMatrixBindings(matrixWorkbook, formatter)
                If bindings.Count = 0 Then Return

                Dim sourceRowCache As New Dictionary(Of String, List(Of WorkbookSourceRow))(StringComparer.OrdinalIgnoreCase)

                For Each binding In bindings
                    If binding Is Nothing Then Continue For

                    Dim matrixSheet As ISheet = matrixWorkbook.GetSheet(binding.SheetName)
                    If matrixSheet Is Nothing Then Continue For

                    Dim sourceSheet As ISheet = FindFeatureSheet(overrideWorkbook, binding.Feature)
                    If sourceSheet Is Nothing Then Continue For

                    Dim sourceRows As List(Of WorkbookSourceRow) = Nothing
                    If Not sourceRowCache.TryGetValue(sourceSheet.SheetName, sourceRows) Then
                        sourceRows = ReadWorkbookSourceRows(sourceSheet, formatter)
                        sourceRowCache(sourceSheet.SheetName) = sourceRows
                    End If

                    ApplyMatrixSheet(binding, matrixSheet, sourceSheet, sourceRows, formatter)
                Next
            End Using
        End Sub

        Private Shared Function ReadMatrixBindings(matrixWorkbook As IWorkbook,
                                                   formatter As DataFormatter) As List(Of MatrixSheetBinding)
            Dim result As New List(Of MatrixSheetBinding)()
            If matrixWorkbook Is Nothing OrElse formatter Is Nothing Then Return result

            Dim indexSheet As ISheet = matrixWorkbook.GetSheet(MatrixIndexSheetName)
            If indexSheet Is Nothing Then Return result

            For rowIndex As Integer = indexSheet.FirstRowNum + 1 To indexSheet.LastRowNum
                Dim row As IRow = indexSheet.GetRow(rowIndex)
                If row Is Nothing Then Continue For

                Dim sheetName As String = ReadCellText(row, 0, formatter)
                Dim feature As String = ReadCellText(row, 1, formatter)
                Dim section As String = ReadCellText(row, 2, formatter)
                If sheetName = String.Empty OrElse feature = String.Empty OrElse section = String.Empty Then Continue For

                result.Add(New MatrixSheetBinding With {
                    .SheetName = sheetName,
                    .Feature = feature,
                    .Section = section,
                    .Locale = NormalizeMatrixLocale(ReadCellText(row, 3, formatter), sheetName)
                })
            Next

            Return result
        End Function

        Private Shared Function FindFeatureSheet(workbook As IWorkbook,
                                                 feature As String) As ISheet
            If workbook Is Nothing Then Return Nothing

            Dim normalizedFeature As String = NormalizeToken(feature)
            If normalizedFeature = String.Empty Then Return Nothing

            Dim exactSheet As ISheet = workbook.GetSheet(normalizedFeature)
            If exactSheet IsNot Nothing Then Return exactSheet

            Dim normalizedExportKey As String = NormalizeExportKey(normalizedFeature)

            For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                Dim sheet As ISheet = workbook.GetSheetAt(sheetIndex)
                If sheet Is Nothing Then Continue For

                Dim sheetFeature As String = NormalizeFeatureName(sheet.SheetName)
                If String.Equals(sheetFeature, normalizedFeature, StringComparison.OrdinalIgnoreCase) Then
                    Return sheet
                End If

                If String.Equals(NormalizeExportKey(sheetFeature), normalizedExportKey, StringComparison.OrdinalIgnoreCase) Then
                    Return sheet
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function ReadWorkbookSourceRows(sheet As ISheet,
                                                       formatter As DataFormatter) As List(Of WorkbookSourceRow)
            Dim result As New List(Of WorkbookSourceRow)()
            If sheet Is Nothing OrElse formatter Is Nothing Then Return result

            For rowIndex As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                Dim row As IRow = sheet.GetRow(rowIndex)
                If row Is Nothing Then Continue For

                Dim originalHeader As String = ReadCellText(row, 1, formatter)
                If originalHeader = String.Empty Then Continue For

                result.Add(New WorkbookSourceRow With {
                    .RowIndex = rowIndex,
                    .Section = ReadCellText(row, 0, formatter),
                    .OriginalHeader = originalHeader,
                    .Condition = ReadCellText(row, 6, formatter)
                })
            Next

            Return result
        End Function

        Private Shared Sub ApplyMatrixSheet(binding As MatrixSheetBinding,
                                            matrixSheet As ISheet,
                                            overrideSheet As ISheet,
                                            sourceRows As IEnumerable(Of WorkbookSourceRow),
                                            formatter As DataFormatter)
            If binding Is Nothing OrElse matrixSheet Is Nothing OrElse overrideSheet Is Nothing OrElse sourceRows Is Nothing OrElse formatter Is Nothing Then
                Return
            End If

            Dim rowsForSection As New List(Of WorkbookSourceRow)()
            Dim rowIndexByCondition As New Dictionary(Of String, List(Of Integer))(StringComparer.OrdinalIgnoreCase)
            Dim orderedHeaders As New List(Of String)()
            Dim seenHeaders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each sourceRow In sourceRows
                If sourceRow Is Nothing Then Continue For
                If Not String.Equals(NormalizeToken(sourceRow.Section), NormalizeToken(binding.Section), StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                rowsForSection.Add(sourceRow)

                If seenHeaders.Add(sourceRow.OriginalHeader) Then
                    orderedHeaders.Add(sourceRow.OriginalHeader)
                End If

                Dim rowKey As String = BuildMatrixRowKey(sourceRow.Condition, sourceRow.OriginalHeader)
                Dim targetRows As List(Of Integer) = Nothing
                If Not rowIndexByCondition.TryGetValue(rowKey, targetRows) Then
                    targetRows = New List(Of Integer)()
                    rowIndexByCondition(rowKey) = targetRows
                End If

                targetRows.Add(sourceRow.RowIndex)
            Next

            If rowsForSection.Count = 0 OrElse orderedHeaders.Count = 0 Then Return

            Dim headerColumnIndex As Integer = If(binding.Locale = "en", 5, 4)
            Dim valueColumnIndex As Integer = If(binding.Locale = "en", 10, 9)
            Dim matrixHeaderRow As IRow = matrixSheet.GetRow(matrixSheet.FirstRowNum)
            If matrixHeaderRow Is Nothing Then Return

            For headerOffset As Integer = 0 To orderedHeaders.Count - 1
                Dim headerValue As String = ReadCellText(matrixHeaderRow, headerOffset + 1, formatter)
                Dim originalHeader As String = orderedHeaders(headerOffset)

                For Each sourceRow In rowsForSection
                    If sourceRow Is Nothing Then Continue For
                    If Not String.Equals(sourceRow.OriginalHeader, originalHeader, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    WriteCellText(overrideSheet, sourceRow.RowIndex, headerColumnIndex, headerValue, formatter)
                Next
            Next

            For rowIndex As Integer = matrixSheet.FirstRowNum + 1 To matrixSheet.LastRowNum
                Dim matrixRow As IRow = matrixSheet.GetRow(rowIndex)
                If matrixRow Is Nothing Then Continue For

                Dim condition As String = ReadCellText(matrixRow, 0, formatter)
                If condition = String.Empty Then Continue For

                For headerOffset As Integer = 0 To orderedHeaders.Count - 1
                    Dim originalHeader As String = orderedHeaders(headerOffset)
                    Dim rowKey As String = BuildMatrixRowKey(condition, originalHeader)
                    Dim targetRows As List(Of Integer) = Nothing
                    If Not rowIndexByCondition.TryGetValue(rowKey, targetRows) OrElse targetRows Is Nothing Then
                        Continue For
                    End If

                    Dim value As String = ReadCellText(matrixRow, headerOffset + 1, formatter)
                    For Each targetRowIndex In targetRows
                        WriteCellText(overrideSheet, targetRowIndex, valueColumnIndex, value, formatter)
                    Next
                Next
            Next
        End Sub

        Private Shared Function BuildMatrixRowKey(condition As String,
                                                  originalHeader As String) As String
            Return NormalizeToken(condition) & "|" & NormalizeToken(originalHeader)
        End Function

        Private Shared Function NormalizeMatrixLocale(value As String,
                                                      sheetName As String) As String
            Dim normalizedValue As String = NormalizeToken(value).ToLowerInvariant()
            If normalizedValue = "en" OrElse normalizedValue = "eng" OrElse normalizedValue = "english" Then
                Return "en"
            End If

            Dim normalizedSheetName As String = NormalizeToken(sheetName).ToUpperInvariant()
            If normalizedSheetName.EndsWith("_EN", StringComparison.Ordinal) Then Return "en"

            Return "ko"
        End Function

        Private Shared Function ReadCellText(row As IRow,
                                             cellIndex As Integer,
                                             formatter As DataFormatter) As String
            If row Is Nothing OrElse formatter Is Nothing OrElse cellIndex < 0 Then Return String.Empty

            Try
                Return NormalizeToken(formatter.FormatCellValue(row.GetCell(cellIndex)))
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Sub WriteCellText(sheet As ISheet,
                                         rowIndex As Integer,
                                         cellIndex As Integer,
                                         value As String,
                                         formatter As DataFormatter)
            If sheet Is Nothing OrElse formatter Is Nothing OrElse rowIndex < 0 OrElse cellIndex < 0 Then Return

            Dim row As IRow = sheet.GetRow(rowIndex)
            If row Is Nothing Then Return

            Dim newValue As String = NormalizeToken(value)
            Dim currentValue As String = ReadCellText(row, cellIndex, formatter)
            If String.Equals(currentValue, newValue, StringComparison.Ordinal) Then Return

            Dim cell As ICell = row.GetCell(cellIndex)
            If cell Is Nothing Then
                cell = row.CreateCell(cellIndex)
            End If

            cell.SetCellValue(newValue)
        End Sub

        Private Shared Sub AddEntry(entriesByKey As Dictionary(Of String, List(Of OverrideEntry)),
                                    bucketKey As String,
                                    entry As OverrideEntry)
            If String.IsNullOrWhiteSpace(bucketKey) OrElse entry Is Nothing Then Return

            Dim bucket As List(Of OverrideEntry) = Nothing
            If Not entriesByKey.TryGetValue(bucketKey, bucket) Then
                bucket = New List(Of OverrideEntry)()
                entriesByKey(bucketKey) = bucket
            End If

            bucket.Add(entry)
        End Sub

        Private Shared Function ResolveOverrideWorkbookPath() As String
            Return ResolveOverrideFilePath(OverrideWorkbookFileName)
        End Function

        Private Shared Function ResolveOverrideJsonPath() As String
            Return ResolveOverrideFilePath(OverrideJsonFileName)
        End Function

        Private Shared Function ResolveMatrixWorkbookPath() As String
            Dim candidates As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each root In GetSearchRoots()
                Dim dir As DirectoryInfo = Nothing
                Try
                    dir = New DirectoryInfo(root)
                Catch
                    dir = Nothing
                End Try

                For depth As Integer = 0 To 6
                    If dir Is Nothing Then Exit For
                    AddCandidate(candidates, seen, Path.Combine(dir.FullName, "docs", MatrixWorkbookFileName))
                    AddCandidate(candidates, seen, Path.Combine(dir.FullName, "KKY_Tool_Revit", "docs", MatrixWorkbookFileName))
                    dir = dir.Parent
                Next
            Next

            For Each candidate In candidates
                If File.Exists(candidate) Then Return candidate
            Next

            Return String.Empty
        End Function

        Private Shared Function BuildCacheSignature(prefix As String, ParamArray paths() As String) As String
            Dim parts As New List(Of String)()
            parts.Add(NormalizeToken(prefix))

            If paths IsNot Nothing Then
                For Each path In paths
                    Dim normalizedPath As String = NormalizeToken(path)
                    If normalizedPath = String.Empty Then Continue For

                    Dim ticks As Long = 0
                    Try
                        ticks = File.GetLastWriteTimeUtc(normalizedPath).Ticks
                    Catch
                    End Try

                    parts.Add(normalizedPath)
                    parts.Add(ticks.ToString())
                Next
            End If

            Return String.Join("|", parts.ToArray())
        End Function

        Private Shared Function ResolveOverrideFilePath(fileName As String) As String
            Dim candidates As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each root In GetSearchRoots()
                Dim dir As DirectoryInfo = Nothing
                Try
                    dir = New DirectoryInfo(root)
                Catch
                    dir = Nothing
                End Try

                For depth As Integer = 0 To 6
                    If dir Is Nothing Then Exit For
                    AddCandidate(candidates, seen, Path.Combine(dir.FullName, "Resources", "Overrides", fileName))
                    AddCandidate(candidates, seen, Path.Combine(dir.FullName, "KKY_Tool_Revit_2019-2023", "Resources", "Overrides", fileName))
                    dir = dir.Parent
                Next
            Next

            For Each candidate In candidates
                If File.Exists(candidate) Then Return candidate
            Next

            Return String.Empty
        End Function

        Private Shared Function GetSearchRoots() As List(Of String)
            Dim roots As New List(Of String)()

            AddDirectory(roots, AppDomain.CurrentDomain.BaseDirectory)

            Try
                AddDirectory(roots, Directory.GetCurrentDirectory())
            Catch
            End Try

            Try
                Dim assemblyPath As String = GetType(ReviewExportMessageOverrideService).Assembly.Location
                If Not String.IsNullOrWhiteSpace(assemblyPath) Then
                    AddDirectory(roots, Path.GetDirectoryName(assemblyPath))
                End If
            Catch
            End Try

            Return roots
        End Function

        Private Shared Sub AddDirectory(target As ICollection(Of String), path As String)
            If target Is Nothing OrElse String.IsNullOrWhiteSpace(path) Then Return
            If target.Contains(path) Then Return
            target.Add(path)
        End Sub

        Private Shared Sub AddCandidate(target As ICollection(Of String),
                                        seen As ISet(Of String),
                                        path As String)
            If target Is Nothing OrElse seen Is Nothing OrElse String.IsNullOrWhiteSpace(path) Then Return
            If seen.Add(path) Then target.Add(path)
        End Sub

        Private Shared Function BuildKey(exportKey As String,
                                         originalHeader As String,
                                         exportHeader As String) As String
            Return String.Join("|", New String() {
                NormalizeToken(exportKey),
                NormalizeToken(originalHeader),
                NormalizeToken(exportHeader)
            })
        End Function

        Private Shared Function BuildHeaderKey(exportKey As String,
                                               originalHeader As String,
                                               exportHeader As String) As String
            Return BuildKey(exportKey, originalHeader, exportHeader)
        End Function

        Private Shared Function NormalizeToken(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Shared Function NormalizeLocale(locale As String) As String
            Dim normalized As String = If(locale, String.Empty).Trim().ToLowerInvariant()
            If normalized = "en" OrElse normalized = "eng" OrElse normalized = "english" Then Return "en"
            Return "ko"
        End Function

        Private Shared Function NormalizeExportKey(value As String) As String
            Dim text As String = NormalizeToken(value).ToLowerInvariant()
            If text = String.Empty Then Return String.Empty

            If text.Contains("duplicate") OrElse text.Contains("clash") OrElse text.Contains("dupclash") Then Return "dupclash"
            If text.Contains("linkworkset") OrElse text.Contains("link workset") Then Return "linkworkset"
            If text.Contains("guid") Then Return "guid"
            If text.Contains("familylink") OrElse text.Contains("family link") Then Return "familylink"
            If text.Contains("floorinfo") Then Return "floorinfo"
            If text.Contains("familysuitability") Then Return "familysuitability"
            If text.Contains("worksetassignment") Then Return "worksetassignment"
            If text.Contains("tapalign") Then Return "tapalign"
            If text.Contains("sharedparambatch") Then Return "sharedparambatch"
            If text.Contains("projectparameterduplication") OrElse text = "parameterduplication" OrElse text.Contains("param") Then Return "paramprop"
            If text.Contains("pms") OrElse text.Contains("segment") Then Return "pms"
            If text.Contains("connector") Then Return "connector"

            Return text
        End Function

        Private Shared Sub AddCandidate(target As ICollection(Of String), text As String)
            If target Is Nothing Then Return

            Dim value As String = NormalizeComparableText(text)
            If value = String.Empty Then Return
            If target.Contains(value) Then Return

            target.Add(value)
        End Sub

        Private Shared Function NormalizeTexts(values As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            If values Is Nothing Then Return result

            For Each value In values
                AddCandidate(result, value)
            Next

            Return result
        End Function

        Private Shared Function MatchesAny(matchTexts As IEnumerable(Of String),
                                           candidates As IEnumerable(Of String)) As Boolean
            If matchTexts Is Nothing OrElse candidates Is Nothing Then Return False

            Dim candidateList As New List(Of String)()
            Dim candidateSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each candidate In candidates
                Dim candidateText As String = NormalizeComparableText(candidate)
                If candidateText = String.Empty Then Continue For
                If candidateSet.Add(candidateText) Then candidateList.Add(candidateText)
            Next

            If candidateList.Count = 0 Then Return False

            For Each value In matchTexts
                Dim matchText As String = NormalizeComparableText(value)
                If matchText = String.Empty Then Continue For
                If candidateSet.Contains(matchText) Then Return True
                If TemplateMatchesAny(matchText, candidateList) Then Return True
            Next

            Return False
        End Function

        Private Shared Function NormalizeOverrideOutput(value As String) As String
            Dim text As String = NormalizeToken(value)
            If text = String.Empty Then Return Nothing
            If String.Equals(text, "N/A", StringComparison.OrdinalIgnoreCase) Then Return Nothing
            Return value
        End Function

        Private Shared Function NormalizeHeaderOverride(value As String) As String
            Dim text As String = NormalizeToken(value)
            If text = String.Empty Then Return String.Empty
            If String.Equals(text, "N/A", StringComparison.OrdinalIgnoreCase) Then Return String.Empty
            Return value
        End Function

        Private Shared Function NormalizeComparableText(value As String) As String
            Dim text As String = NormalizeToken(value)
            If text = String.Empty Then Return String.Empty

            text = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            text = Regex.Replace(text, "\s+", " ")
            Return text.Trim()
        End Function

        Private Shared Function TemplateMatchesAny(template As String,
                                                   candidates As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(template) OrElse candidates Is Nothing Then Return False
            If template.IndexOf("{"c) < 0 OrElse template.IndexOf("}"c) < 0 Then Return False

            Const wildcardToken As String = "__CODex_PLACEHOLDER_WILDCARD__"
            Dim hasPlaceholder As Boolean = False
            Dim tokenizedTemplate As String =
                Regex.Replace(template,
                              "\{[^{}]+\}",
                              Function(match)
                                  hasPlaceholder = True
                                  Return wildcardToken
                              End Function)
            If Not hasPlaceholder Then Return False

            Dim wildcardPattern As String = Regex.Escape(tokenizedTemplate).Replace(Regex.Escape(wildcardToken), ".*?")

            Dim matcher As New Regex("^" & wildcardPattern & "$", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            For Each candidate In candidates
                Dim candidateText As String = NormalizeComparableText(candidate)
                If candidateText = String.Empty Then Continue For
                If matcher.IsMatch(candidateText) Then Return True
            Next

            Return False
        End Function

        Private Shared Function ApplyTemplateCaptures(localizedText As String,
                                                      matchTexts As IEnumerable(Of String),
                                                      candidates As IEnumerable(Of String)) As String
            If String.IsNullOrWhiteSpace(localizedText) OrElse matchTexts Is Nothing OrElse candidates Is Nothing Then
                Return localizedText
            End If

            For Each template In matchTexts
                For Each candidate In candidates
                    Dim captures As Dictionary(Of String, String) = Nothing
                    If Not TryGetTemplateCaptures(template, candidate, captures) Then Continue For
                    Return ApplyCapturedValues(localizedText, captures)
                Next
            Next

            Return localizedText
        End Function

        Private Shared Function TryGetTemplateCaptures(template As String,
                                                       candidate As String,
                                                       ByRef captures As Dictionary(Of String, String)) As Boolean
            captures = Nothing

            Dim normalizedTemplate As String = NormalizeComparableText(template)
            Dim normalizedCandidate As String = NormalizeComparableText(candidate)
            If normalizedTemplate = String.Empty OrElse normalizedCandidate = String.Empty Then Return False
            If normalizedTemplate.IndexOf("{"c) < 0 OrElse normalizedTemplate.IndexOf("}"c) < 0 Then Return False

            Dim tokenNames As New List(Of String)()
            Dim pattern As New StringBuilder()
            Dim cursor As Integer = 0

            While cursor < normalizedTemplate.Length
                Dim openIndex As Integer = normalizedTemplate.IndexOf("{"c, cursor)
                If openIndex < 0 Then
                    pattern.Append(Regex.Escape(normalizedTemplate.Substring(cursor)))
                    Exit While
                End If

                Dim closeIndex As Integer = normalizedTemplate.IndexOf("}"c, openIndex + 1)
                If closeIndex < 0 Then Return False

                pattern.Append(Regex.Escape(normalizedTemplate.Substring(cursor, openIndex - cursor)))

                Dim tokenName As String = normalizedTemplate.Substring(openIndex + 1, closeIndex - openIndex - 1).Trim()
                If tokenName = String.Empty Then Return False

                tokenNames.Add(tokenName)
                pattern.Append("(.*?)")
                cursor = closeIndex + 1
            End While

            If tokenNames.Count = 0 Then Return False

            Dim matcher As New Regex("^" & pattern.ToString() & "$", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            Dim match As Match = matcher.Match(normalizedCandidate)
            If Not match.Success Then Return False

            captures = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For index As Integer = 0 To tokenNames.Count - 1
                Dim value As String = match.Groups(index + 1).Value
                If Not captures.ContainsKey(tokenNames(index)) Then
                    captures(tokenNames(index)) = value
                End If
            Next

            Return captures.Count > 0
        End Function

        Private Shared Function ApplyCapturedValues(localizedText As String,
                                                    captures As Dictionary(Of String, String)) As String
            If String.IsNullOrWhiteSpace(localizedText) OrElse captures Is Nothing OrElse captures.Count = 0 Then
                Return localizedText
            End If

            Dim result As String = localizedText
            For Each pair In captures
                result = result.Replace("{" & pair.Key & "}", pair.Value)
            Next

            Dim actualWorksetName As String = Nothing
            If captures.TryGetValue("실제웍셋", actualWorksetName) AndAlso Not String.IsNullOrWhiteSpace(actualWorksetName) Then
                result = result.Replace("Workset Name", actualWorksetName)
            End If

            Return result
        End Function

        Private Shared Function BuildSheetHeaderMap(headerRow As IRow,
                                                    formatter As DataFormatter) As Dictionary(Of String, Integer)
            Dim result As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            If headerRow Is Nothing Then Return result

            Dim lastCell As Integer = Math.Max(0, CInt(headerRow.LastCellNum) - 1)
            For cellIndex As Integer = 0 To lastCell
                Dim text As String = String.Empty
                Try
                    text = NormalizeToken(formatter.FormatCellValue(headerRow.GetCell(cellIndex)))
                Catch
                    text = String.Empty
                End Try

                If text = String.Empty OrElse result.ContainsKey(text) Then Continue For
                result(text) = cellIndex
            Next

            Return result
        End Function

        Private Shared Function HasWorkbookOverrideHeaders(headerMap As Dictionary(Of String, Integer)) As Boolean
            If headerMap Is Nothing OrElse headerMap.Count = 0 Then Return False

            If Not (headerMap.ContainsKey("원본/내부 열명") AndAlso
                    headerMap.ContainsKey("한글 출력") AndAlso
                    headerMap.ContainsKey("영문 출력") AndAlso
                    headerMap.ContainsKey("원본(raw)")) Then
                Return False
            End If

            Return headerMap.ContainsKey("한글 override") OrElse
                   headerMap.ContainsKey("한글 적용문구") OrElse
                   headerMap.ContainsKey("한국어 override") OrElse
                   headerMap.ContainsKey("한국어 적용문구") OrElse
                   headerMap.ContainsKey("영문 override") OrElse
                   headerMap.ContainsKey("영문 적용문구") OrElse
                   headerMap.ContainsKey("English override") OrElse
                   headerMap.ContainsKey("English Applied Text")
        End Function

        Private Shared Function HasWorkbookHeaderOverrideHeaders(headerMap As Dictionary(Of String, Integer)) As Boolean
            If headerMap Is Nothing OrElse headerMap.Count = 0 Then Return False
            If Not headerMap.ContainsKey("원본/내부 열명") Then Return False
            If Not (headerMap.ContainsKey("한글 export 열명") OrElse headerMap.ContainsKey("영문 export 열명")) Then Return False

            Return headerMap.ContainsKey("한글 헤더 override") OrElse
                   headerMap.ContainsKey("영문 헤더 override")
        End Function

        Private Shared Function ReadCellText(row As IRow,
                                             headerMap As Dictionary(Of String, Integer),
                                             formatter As DataFormatter,
                                             ParamArray names() As String) As String
            If row Is Nothing OrElse headerMap Is Nothing OrElse formatter Is Nothing OrElse names Is Nothing Then Return String.Empty

            For Each name In names
                Dim index As Integer = -1
                If Not headerMap.TryGetValue(name, index) Then Continue For

                Try
                    Return NormalizeToken(formatter.FormatCellValue(row.GetCell(index)))
                Catch
                End Try
            Next

            Return String.Empty
        End Function

        Private Shared Function NormalizeFeatureName(sheetName As String) As String
            Dim text As String = NormalizeToken(sheetName)
            If text = "Duplicate+Clash" Then Return "Duplicate/Clash"
            If text = "요약" OrElse text = "메모" OrElse text = "전체정리" Then Return String.Empty
            Return text
        End Function

        Private Shared Function ShouldSkipOverrideSheet(sheetName As String) As Boolean
            Return NormalizeFeatureName(sheetName) = String.Empty
        End Function

        Private Shared Sub MergeHeaderEntry(entriesByKey As Dictionary(Of String, HeaderOverrideEntry),
                                            key As String,
                                            koDefaultHeader As String,
                                            enDefaultHeader As String,
                                            koOverrideHeader As String,
                                            enOverrideHeader As String)
            If entriesByKey Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return

            Dim entry As HeaderOverrideEntry = Nothing
            If Not entriesByKey.TryGetValue(key, entry) OrElse entry Is Nothing Then
                entry = New HeaderOverrideEntry()
                entriesByKey(key) = entry
            End If

            MergeLocalizedHeader(entry, koDefaultHeader, koOverrideHeader, isEnglish:=False)
            MergeLocalizedHeader(entry, enDefaultHeader, enOverrideHeader, isEnglish:=True)
        End Sub

        Private Shared Sub MergeLocalizedHeader(entry As HeaderOverrideEntry,
                                                defaultHeader As String,
                                                overrideHeader As String,
                                                isEnglish As Boolean)
            If entry Is Nothing Then Return

            Dim normalizedDefault As String = NormalizeToken(defaultHeader)
            Dim normalizedOverride As String = NormalizeToken(overrideHeader)
            Dim hasChangedOverride As Boolean =
                normalizedOverride <> String.Empty AndAlso
                Not String.Equals(normalizedOverride, normalizedDefault, StringComparison.Ordinal)

            If isEnglish Then
                If hasChangedOverride Then
                    entry.EnHeader = normalizedOverride
                    entry.EnChanged = True
                ElseIf Not entry.EnChanged AndAlso String.IsNullOrWhiteSpace(entry.EnHeader) Then
                    entry.EnHeader = If(normalizedOverride <> String.Empty, normalizedOverride, normalizedDefault)
                End If
            Else
                If hasChangedOverride Then
                    entry.KoHeader = normalizedOverride
                    entry.KoChanged = True
                ElseIf Not entry.KoChanged AndAlso String.IsNullOrWhiteSpace(entry.KoHeader) Then
                    entry.KoHeader = If(normalizedOverride <> String.Empty, normalizedOverride, normalizedDefault)
                End If
            End If
        End Sub

    End Class

End Namespace
