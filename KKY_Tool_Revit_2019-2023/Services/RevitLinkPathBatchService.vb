Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application
Imports Autodesk.Revit.DB
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Services

    Public Class RevitLinkPathRow
        Public Property HostFileName As String = ""
        Public Property HostFilePath As String = ""
        Public Property ReferenceElementId As String = ""
        Public Property LinkName As String = ""
        Public Property LinkFileName As String = ""
        Public Property TypeWorksetNames As String = ""
        Public Property InstanceWorksetNames As String = ""
        Public Property ApplyTypeWorksetNames As String = ""
        Public Property ApplyInstanceWorksetNames As String = ""
        Public Property CurrentLinkPath As String = ""
        Public Property StoredLinkPath As String = ""
        Public Property CurrentPathType As String = ""
        Public Property TargetLinkPath As String = ""
        Public Property TargetPathType As String = ""
        Public Property ApplyStatus As String = ""
        Public Property ApplyMessage As String = ""
    End Class

    Public Class RevitLinkPathApplyOptions
        Public Property NewLinkPlacement As String = "Origin"
        Public Property HostPathHints As IList(Of String) = Nothing
    End Class

    Public NotInheritable Class RevitLinkPathBatchService

        Private Sub New()
        End Sub

        Public Shared Function Extract(app As RevitApp,
                                       rvtPaths As IList(Of String),
                                       progress As Action(Of Integer, String)) As List(Of RevitLinkPathRow)

            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim rows As New List(Of RevitLinkPathRow)()
            Dim initialOpenDocs As HashSet(Of Integer) = CaptureOpenDocumentHandles(app)
            Dim cleanPaths As List(Of String) = NormalizePaths(rvtPaths)
            Dim total As Integer = cleanPaths.Count
            If total = 0 Then Return rows

            Try
                For i As Integer = 0 To total - 1
                    Dim hostPath As String = cleanPaths(i)
                    Dim hostName As String = SafeFileName(hostPath)
                    ReportWeightedProgress(progress, total, i, 0.05R, $"[{i + 1}/{total}] RVT 확인 중... {hostName}")

                    Try
                        If String.IsNullOrWhiteSpace(hostPath) Then
                            Throw New ArgumentException("RVT 경로가 비어 있습니다.")
                        End If
                        If Not File.Exists(hostPath) AndAlso Not IsServerPath(hostPath) Then
                            Throw New FileNotFoundException("RVT 파일을 찾을 수 없습니다.", hostPath)
                        End If

                        ReportWeightedProgress(progress, total, i, 0.18R, $"[{i + 1}/{total}] 링크 데이터 읽는 중... {hostName}")

                        Dim hostRows As List(Of RevitLinkPathRow) = ExtractRowsFromTransmission(hostPath, hostName, progress, total, i)
                        If hostRows.Count = 0 Then
                            ReportWeightedProgress(progress, total, i, 0.52R, $"[{i + 1}/{total}] 실제 문서 열어 링크 확인 중... {hostName}")
                            hostRows = ExtractRowsFromDocument(app, hostPath, hostName, progress, total, i)
                        Else
                            ReportWeightedProgress(progress, total, i, 0.52R, $"[{i + 1}/{total}] 호스트 웍셋/클라우드 링크 확인 중... {hostName}")
                            MergeDocumentRowsIntoExtractedRows(app, hostPath, hostName, hostRows, progress, total, i)
                        End If

                        If hostRows.Count = 0 Then
                            rows.Add(New RevitLinkPathRow With {
                                .HostFileName = hostName,
                                .HostFilePath = hostPath,
                                .ApplyStatus = "info",
                                .ApplyMessage = "Revit 링크가 없습니다."
                            })
                        Else
                            rows.AddRange(hostRows)
                        End If

                        ReportWeightedProgress(progress, total, i, 0.98R, $"[{i + 1}/{total}] 링크 추출 완료 ({hostRows.Count}개)... {hostName}")

                    Catch ex As Exception
                        rows.Add(New RevitLinkPathRow With {
                            .HostFileName = hostName,
                            .HostFilePath = hostPath,
                            .ApplyStatus = "error",
                            .ApplyMessage = "링크 추출 실패: " & ex.Message
                        })
                        ReportWeightedProgress(progress, total, i, 0.98R, $"[{i + 1}/{total}] 추출 실패... {hostName}")
                    End Try
                Next

                ReportProgress(progress, total, total, "링크 추출 완료")
                Return rows
            Finally
                CloseDocumentsOpenedDuring(app, initialOpenDocs)
            End Try
        End Function

        Private Shared Function ExtractRowsFromTransmission(hostPath As String,
                                                            hostName As String,
                                                            progress As Action(Of Integer, String),
                                                            total As Integer,
                                                            itemIndex As Integer) As List(Of RevitLinkPathRow)

            Dim rows As New List(Of RevitLinkPathRow)()
            Dim modelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(hostPath)
            Dim transData As TransmissionData = TransmissionData.ReadTransmissionData(modelPath)
            If transData Is Nothing Then Return rows

            Dim refIds As List(Of ElementId) = transData.GetAllExternalFileReferenceIds().ToList()
            ReportWeightedProgress(progress, total, itemIndex, 0.32R, $"[{itemIndex + 1}/{total}] 외부 참조 확인 중 ({refIds.Count}개)... {hostName}")

            For refIndex As Integer = 0 To refIds.Count - 1
                Dim refId As ElementId = refIds(refIndex)
                Dim scanFraction As Double = 0.32R + (0.16R * ((CDbl(refIndex) + 1.0R) / CDbl(Math.Max(1, refIds.Count))))
                ReportWeightedProgress(progress, total, itemIndex, scanFraction, $"[{itemIndex + 1}/{total}] 링크 스캔 중 ({refIndex + 1}/{refIds.Count})... {hostName}")

                Dim extRef As ExternalFileReference = Nothing
                Try
                    extRef = transData.GetLastSavedReferenceData(refId)
                Catch
                    extRef = Nothing
                End Try
                If extRef Is Nothing Then Continue For

                Dim refType As ExternalFileReferenceType
                Try
                    refType = extRef.ExternalFileReferenceType
                Catch
                    Continue For
                End Try
                If refType <> ExternalFileReferenceType.RevitLink Then Continue For

                Dim storedPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(extRef))
                Dim absolutePath As String = SafeModelPathToUserVisiblePath(TryGetAbsoluteModelPath(extRef))
                Dim visiblePath As String = If(Not String.IsNullOrWhiteSpace(absolutePath), absolutePath, storedPath)
                Dim linkFileName As String = SafeFileName(If(Not String.IsNullOrWhiteSpace(visiblePath), visiblePath, storedPath))
                Dim linkName As String = If(Not String.IsNullOrWhiteSpace(linkFileName),
                                            Path.GetFileNameWithoutExtension(linkFileName),
                                            $"Link #{SafeElementIdText(refId)}")

                rows.Add(New RevitLinkPathRow With {
                    .HostFileName = hostName,
                    .HostFilePath = hostPath,
                    .ReferenceElementId = SafeElementIdText(refId),
                    .LinkName = linkName,
                    .LinkFileName = linkFileName,
                    .TypeWorksetNames = "",
                    .InstanceWorksetNames = "",
                    .ApplyTypeWorksetNames = "",
                    .ApplyInstanceWorksetNames = "",
                    .CurrentLinkPath = visiblePath,
                    .StoredLinkPath = storedPath,
                    .CurrentPathType = SafePathTypeText(extRef),
                    .TargetLinkPath = "",
                    .TargetPathType = "",
                    .ApplyStatus = "",
                    .ApplyMessage = ""
                })
            Next

            Return rows
        End Function

        Private Shared Function ExtractRowsFromDocument(app As RevitApp,
                                                        hostPath As String,
                                                        hostName As String,
                                                        progress As Action(Of Integer, String),
                                                        total As Integer,
                                                        itemIndex As Integer) As List(Of RevitLinkPathRow)

            Dim rows As New List(Of RevitLinkPathRow)()
            If app Is Nothing Then Return rows

            Dim doc As Document = TryFindOpenDocument(app, hostPath)
            Dim openPath As String = hostPath
            Dim createdLocal As Boolean = False
            Dim openedHere As Boolean = False

            Try
                If doc Is Nothing Then
                    Dim basicInfo As BasicFileInfo = TryExtractBasicFileInfo(hostPath)
                    If basicInfo IsNot Nothing AndAlso basicInfo.IsCentral Then
                        ReportWeightedProgress(progress, total, itemIndex, 0.60R, $"[{itemIndex + 1}/{total}] 센트럴 로컬 파일 생성 중... {hostName}")
                        openPath = CreateNewLocalPath(hostPath)
                        createdLocal = True
                    End If

                    ReportWeightedProgress(progress, total, itemIndex, 0.68R, $"[{itemIndex + 1}/{total}] RVT 열기 중 (웍셋 닫기)... {hostName}")
                    doc = OpenProjectDocument(app, openPath, closeAllWorksets:=True)
                    openedHere = (doc IsNot Nothing)
                Else
                    ReportWeightedProgress(progress, total, itemIndex, 0.68R, $"[{itemIndex + 1}/{total}] 이미 열린 문서에서 링크 확인 중... {hostName}")
                End If

                If doc Is Nothing Then Return rows
                Return ExtractRowsFromOpenDocument(doc, hostPath, hostName, progress, total, itemIndex)
            Finally
                If openedHere Then
                    SafeClose(doc)
                End If
                If createdLocal Then
                    TryDeleteFile(openPath)
                End If
            End Try
        End Function

        Private Shared Function ExtractRowsFromOpenDocument(doc As Document,
                                                            hostPath As String,
                                                            hostName As String,
                                                            progress As Action(Of Integer, String),
                                                            total As Integer,
                                                            itemIndex As Integer) As List(Of RevitLinkPathRow)
            Dim rows As New List(Of RevitLinkPathRow)()
            If doc Is Nothing Then Return rows

            Dim linkTypes As List(Of RevitLinkType) =
                New FilteredElementCollector(doc).
                    OfClass(GetType(RevitLinkType)).
                    Cast(Of RevitLinkType)().
                    Where(Function(x) x IsNot Nothing AndAlso Not x.IsNestedLink).
                    OrderBy(Function(x) SafeStr(x.Name), StringComparer.OrdinalIgnoreCase).
                    ToList()

            If linkTypes.Count = 0 Then Return rows

            For linkIndex As Integer = 0 To linkTypes.Count - 1
                Dim linkType As RevitLinkType = linkTypes(linkIndex)
                Dim linkName As String = SafeStr(linkType.Name)
                Dim scanFraction As Double = 0.74R + (0.20R * ((CDbl(linkIndex) + 1.0R) / CDbl(Math.Max(1, linkTypes.Count))))
                ReportWeightedProgress(progress, total, itemIndex, scanFraction, $"[{itemIndex + 1}/{total}] 열린 문서 링크 확인 중 ({linkIndex + 1}/{linkTypes.Count})... {linkName}")

                Dim resourceRef As ExternalResourceReference = TryGetExternalResourceReference(linkType)
                Dim currentPath As String = GetCurrentVisibleLinkPath(linkType, Nothing)
                Dim storedPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
                If String.IsNullOrWhiteSpace(storedPath) Then
                    storedPath = SerializeExternalResourceReference(resourceRef)
                End If
                Dim linkFileName As String = SafeFileName(If(Not String.IsNullOrWhiteSpace(currentPath), currentPath, storedPath))
                If String.IsNullOrWhiteSpace(linkFileName) Then
                    linkFileName = GetExternalResourceShortName(resourceRef)
                End If

                rows.Add(New RevitLinkPathRow With {
                    .HostFileName = hostName,
                    .HostFilePath = hostPath,
                    .ReferenceElementId = SafeElementIdText(linkType.Id),
                    .LinkName = If(Not String.IsNullOrWhiteSpace(linkName), linkName, $"Link #{SafeElementIdText(linkType.Id)}"),
                    .LinkFileName = linkFileName,
                    .TypeWorksetNames = GetLinkTypeWorksetNamesText(doc, linkType.Id),
                    .InstanceWorksetNames = GetLinkInstanceWorksetNamesText(doc, linkType.Id),
                    .ApplyTypeWorksetNames = GetLinkTypeWorksetNamesText(doc, linkType.Id),
                    .ApplyInstanceWorksetNames = GetLinkInstanceWorksetNamesText(doc, linkType.Id),
                    .CurrentLinkPath = currentPath,
                    .StoredLinkPath = storedPath,
                    .CurrentPathType = SafePathTypeText(linkType),
                    .TargetLinkPath = "",
                    .TargetPathType = "",
                    .ApplyStatus = "",
                    .ApplyMessage = ""
                })
            Next

            Return rows
        End Function

        Public Shared Function ImportWorkbook(xlsxPath As String) As List(Of RevitLinkPathRow)
            Dim rows As New List(Of RevitLinkPathRow)()
            Dim normalizedPath As String = NormalizeUserVisiblePath(xlsxPath)
            If String.IsNullOrWhiteSpace(normalizedPath) Then
                Throw New ArgumentException("엑셀 경로가 비어 있습니다.", NameOf(xlsxPath))
            End If
            If Not File.Exists(normalizedPath) Then
                Throw New FileNotFoundException("엑셀 파일을 찾을 수 없습니다.", normalizedPath)
            End If

            Dim formatter As New DataFormatter(CultureInfo.InvariantCulture)
            Using stream As New FileStream(normalizedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using workbook As IWorkbook = New XSSFWorkbook(stream)
                    For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                        Dim sheet As ISheet = workbook.GetSheetAt(sheetIndex)
                        If sheet Is Nothing Then Continue For

                        Dim headerRow As IRow = sheet.GetRow(sheet.FirstRowNum)
                        If headerRow Is Nothing Then Continue For

                        Dim headers As New Dictionary(Of Integer, String)()
                        For col As Integer = 0 To CInt(headerRow.LastCellNum) - 1
                            headers(col) = GetCellText(headerRow.GetCell(col), formatter)
                        Next

                        If FindHeaderIndex(headers, "HostFilePath") < 0 OrElse FindHeaderIndex(headers, "ReferenceElementId") < 0 Then
                            Continue For
                        End If

                        For rowIndex As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                            Dim row As IRow = sheet.GetRow(rowIndex)
                            If row Is Nothing Then Continue For

                            Dim item As New RevitLinkPathRow With {
                                .HostFileName = GetCellTextByHeader(headers, row, formatter, "HostFileName"),
                                .HostFilePath = NormalizeUserVisiblePath(GetCellTextByHeader(headers, row, formatter, "HostFilePath")),
                                .ReferenceElementId = GetCellTextByHeader(headers, row, formatter, "ReferenceElementId"),
                                .LinkName = GetCellTextByHeader(headers, row, formatter, "LinkName"),
                                .LinkFileName = GetCellTextByHeader(headers, row, formatter, "LinkFileName"),
                                .TypeWorksetNames = GetCellTextByHeader(headers, row, formatter, "TypeWorksetNames"),
                                .InstanceWorksetNames = GetCellTextByHeader(headers, row, formatter, "InstanceWorksetNames"),
                                .ApplyTypeWorksetNames = GetCellTextByHeader(headers, row, formatter, "ApplyTypeWorksetNames"),
                                .ApplyInstanceWorksetNames = GetCellTextByHeader(headers, row, formatter, "ApplyInstanceWorksetNames"),
                                .CurrentLinkPath = NormalizeUserVisiblePath(GetCellTextByHeader(headers, row, formatter, "CurrentLinkPath")),
                                .StoredLinkPath = NormalizeUserVisiblePath(GetCellTextByHeader(headers, row, formatter, "StoredLinkPath")),
                                .CurrentPathType = GetCellTextByHeader(headers, row, formatter, "CurrentPathType"),
                                .TargetLinkPath = NormalizeUserVisiblePath(GetCellTextByHeader(headers, row, formatter, "TargetLinkPath")),
                                .TargetPathType = GetCellTextByHeader(headers, row, formatter, "TargetPathType"),
                                .ApplyStatus = GetCellTextByHeader(headers, row, formatter, "ApplyStatus"),
                                .ApplyMessage = GetCellTextByHeader(headers, row, formatter, "ApplyMessage")
                            }

                            If String.IsNullOrWhiteSpace(item.TypeWorksetNames) Then
                                item.TypeWorksetNames = GetCellTextByHeader(headers, row, formatter, "HostWorksetNames")
                            End If
                            If String.IsNullOrWhiteSpace(item.InstanceWorksetNames) Then
                                item.InstanceWorksetNames = GetCellTextByHeader(headers, row, formatter, "HostWorksetNames")
                            End If
                            If String.IsNullOrWhiteSpace(item.ApplyTypeWorksetNames) Then
                                item.ApplyTypeWorksetNames = GetCellTextByHeader(headers, row, formatter, "ApplyWorksetNames")
                            End If
                            If String.IsNullOrWhiteSpace(item.ApplyInstanceWorksetNames) Then
                                item.ApplyInstanceWorksetNames = GetCellTextByHeader(headers, row, formatter, "ApplyWorksetNames")
                            End If

                            If String.IsNullOrWhiteSpace(item.HostFileName) AndAlso
                               String.IsNullOrWhiteSpace(item.HostFilePath) AndAlso
                               String.IsNullOrWhiteSpace(item.ReferenceElementId) AndAlso
                               String.IsNullOrWhiteSpace(item.LinkName) AndAlso
                               String.IsNullOrWhiteSpace(item.CurrentLinkPath) AndAlso
                               String.IsNullOrWhiteSpace(item.TargetLinkPath) Then
                                Continue For
                            End If

                            rows.Add(item)
                        Next
                    Next
                End Using
            End Using

            Return rows
        End Function

        Public Shared Function Apply(app As RevitApp,
                                     rows As IList(Of RevitLinkPathRow),
                                     progress As Action(Of Integer, String),
                                     Optional options As RevitLinkPathApplyOptions = Nothing) As List(Of RevitLinkPathRow)

            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim clonedRows As List(Of RevitLinkPathRow) = CloneRows(rows)
            Dim initialOpenDocs As HashSet(Of Integer) = CaptureOpenDocumentHandles(app)
            If clonedRows.Count = 0 Then Return clonedRows
            Dim applyOptions As RevitLinkPathApplyOptions = If(options, New RevitLinkPathApplyOptions())
            ResolveMissingHostFilePaths(clonedRows, applyOptions.HostPathHints)

            For Each row In clonedRows
                If row Is Nothing Then Continue For
                row.TargetLinkPath = NormalizeUserVisiblePath(row.TargetLinkPath)
                row.HostFilePath = NormalizeUserVisiblePath(row.HostFilePath)
                row.CurrentLinkPath = NormalizeUserVisiblePath(row.CurrentLinkPath)
                row.StoredLinkPath = NormalizeUserVisiblePath(row.StoredLinkPath)
                row.TargetPathType = ""

                If (IsNewLinkRow(row) OrElse IsDeleteLinkRow(row)) AndAlso String.IsNullOrWhiteSpace(row.HostFilePath) Then
                    If String.IsNullOrWhiteSpace(row.ApplyStatus) Then row.ApplyStatus = "error"
                    If String.IsNullOrWhiteSpace(row.ApplyMessage) Then row.ApplyMessage = "HostFileName만으로 호스트 RVT 경로를 찾지 못했습니다. RVT 목록을 등록하거나 HostFilePath를 입력해 주세요."
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.ReferenceElementId) AndAlso Not IsNewLinkRow(row) Then
                    If String.IsNullOrWhiteSpace(row.ApplyStatus) Then row.ApplyStatus = "skip"
                    If String.IsNullOrWhiteSpace(row.ApplyMessage) Then row.ApplyMessage = "링크 식별자가 없어 건너뜁니다."
                    Continue For
                End If

                If IsDeleteLinkRow(row) Then
                    row.ApplyStatus = ""
                    row.ApplyMessage = ""
                ElseIf String.IsNullOrWhiteSpace(row.TargetLinkPath) Then
                    row.ApplyStatus = "skip"
                    row.ApplyMessage = "대상 경로가 비어 있어 건너뜁니다."
                Else
                    row.ApplyStatus = ""
                    row.ApplyMessage = ""
                End If
            Next

            Dim groups = clonedRows.
                Where(Function(x) x IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(x.HostFilePath) AndAlso
                                  (IsNewLinkRow(x) OrElse
                                   IsDeleteLinkRow(x) OrElse
                                   (Not String.IsNullOrWhiteSpace(x.ReferenceElementId) AndAlso
                                    Not String.IsNullOrWhiteSpace(x.TargetLinkPath)))).
                GroupBy(Function(x) x.HostFilePath, StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim total As Integer = groups.Count
            If total = 0 Then
                ReportProgress(progress, 1, 1, "적용할 대상 경로가 없습니다.")
                Return clonedRows
            End If

            Try
                For gi As Integer = 0 To groups.Count - 1
                    Dim hostPath As String = groups(gi).Key
                    Dim hostName As String = SafeFileName(hostPath)
                    ReportWeightedProgress(progress, total, gi, 0.03R, $"[{gi + 1}/{total}] 적용 준비 중... {hostName}")

                    Try
                        ApplyHostGroup(app, hostPath, groups(gi).ToList(), progress, gi, total, applyOptions)
                    Catch ex As Exception
                        MarkRows(groups(gi), "error", "링크 경로 적용 실패: " & ex.Message, overwriteChanged:=True)
                        ReportWeightedProgress(progress, total, gi, 0.98R, $"[{gi + 1}/{total}] 적용 실패... {hostName}")
                    End Try
                Next

                ReportProgress(progress, total, total, "링크 경로 적용 완료")
                Return clonedRows
            Finally
                CloseDocumentsOpenedDuring(app, initialOpenDocs)
            End Try
        End Function

        Private Shared Sub ApplyHostGroup(app As RevitApp,
                                          hostPath As String,
                                          rows As IList(Of RevitLinkPathRow),
                                          progress As Action(Of Integer, String),
                                          hostIndex As Integer,
                                          hostTotal As Integer,
                                          options As RevitLinkPathApplyOptions)
            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
            If rows Is Nothing OrElse rows.Count = 0 Then Return

            Dim normalizedHostPath As String = NormalizeUserVisiblePath(hostPath)
            Dim hostName As String = SafeFileName(normalizedHostPath)
            If String.IsNullOrWhiteSpace(normalizedHostPath) Then
                MarkRows(rows, "error", "호스트 RVT 경로가 비어 있습니다.")
                Return
            End If

            If Not IsServerPath(normalizedHostPath) AndAlso Not File.Exists(normalizedHostPath) Then
                MarkRows(rows, "error", "호스트 RVT 파일을 찾을 수 없습니다.")
                Return
            End If

            If IsAlreadyOpen(app, normalizedHostPath) Then
                MarkRows(rows, "error", "이미 열려 있는 문서라서 자동 적용할 수 없습니다.")
                Return
            End If

            ReportWeightedProgress(progress, hostTotal, hostIndex, 0.08R, $"[{hostIndex + 1}/{hostTotal}] 호스트 확인 중... {hostName}")

            Dim basicInfo As BasicFileInfo = TryExtractBasicFileInfo(normalizedHostPath)
            Dim wasCentralFile As Boolean = (basicInfo IsNot Nothing AndAlso basicInfo.IsCentral)
            Dim openPath As String = normalizedHostPath
            Dim doc As Document = Nothing
            Dim changedRows As New List(Of RevitLinkPathRow)()
            Dim currentStage As String = "호스트 준비"
            ' Central hosts are now opened directly. Keep TransmissionData out of that path and apply changes only after opening the document.
            Dim allowTransmissionApply As Boolean = Not wasCentralFile

            Try
                Dim transmissionRows = rows.
                    Where(Function(x) x IsNot Nothing AndAlso CanApplyViaTransmission(x)).
                    ToList()
                If allowTransmissionApply AndAlso transmissionRows.Count > 0 Then
                    currentStage = "닫힌 문서 링크 경로 반영"
                    ReportWeightedProgress(progress, hostTotal, hostIndex, 0.20R, $"[{hostIndex + 1}/{hostTotal}] 닫힌 문서 링크 경로 반영 중... {hostName}")
                    ApplyRowsViaTransmission(openPath, transmissionRows)
                End If

                Dim preferredHostWorksetNames = CollectApplyWorksetNames(rows)
                currentStage = If(wasCentralFile, "센트럴 파일 열기", "호스트 문서 열기")
                ReportWeightedProgress(progress, hostTotal, hostIndex, 0.24R, $"[{hostIndex + 1}/{hostTotal}] RVT 열기 중 (호스트 웍셋 반영)... {hostName}")
                doc = OpenProjectDocumentForApply(app, openPath, preferredHostWorksetNames, includeSavedOpenWorksets:=Not wasCentralFile)
                If doc Is Nothing Then
                    MarkRows(rows, "error", "호스트 문서를 열지 못했습니다.")
                    Return
                End If

                ReportWeightedProgress(progress, hostTotal, hostIndex, 0.34R, $"[{hostIndex + 1}/{hostTotal}] 링크 재로드 준비 중... {hostName}")

                For rowIndex As Integer = 0 To rows.Count - 1
                    Dim row = rows(rowIndex)
                    Dim rowFraction As Double = 0.34R + (0.46R * ((CDbl(rowIndex) + 1.0R) / CDbl(Math.Max(1, rows.Count))))
                    Dim linkName As String = If(String.IsNullOrWhiteSpace(SafeStr(row.LinkName)),
                                                SafeElementIdText(ParseElementIdOrInvalid(row.ReferenceElementId)),
                                                row.LinkName)
                    Dim operationName As String = If(IsDeleteLinkRow(row),
                                                     "링크 삭제 중",
                                                     If(IsNewLinkRow(row), "신규 링크 생성 중", "링크 재로드 중"))
                    currentStage = operationName
                    ReportWeightedProgress(progress, hostTotal, hostIndex, rowFraction, $"[{hostIndex + 1}/{hostTotal}] {operationName} ({rowIndex + 1}/{rows.Count})... {linkName}")

                    Dim rowChanged As Boolean
                    If IsDeleteLinkRow(row) Then
                        rowChanged = DeleteLinkRow(doc, row)
                    ElseIf IsNewLinkRow(row) Then
                        rowChanged = CreateNewLinkRow(doc, row, options)
                    ElseIf allowTransmissionApply AndAlso CanApplyViaTransmission(row) Then
                        rowChanged = RefreshRowAfterTransmission(doc, row)
                    Else
                        rowChanged = ApplySingleRow(doc, row)
                    End If

                    If rowChanged Then
                        changedRows.Add(row)
                    End If
                Next

                Dim retryRows = rows.
                    Where(Function(x) NeedsHostWorksetRetry(x)).
                    ToList()
                If retryRows.Count > 0 Then
                    Dim retryWorksetNames = CollectRetryHostWorksetNames(doc, retryRows)
                    If retryWorksetNames.Count > 0 Then
                        currentStage = "필요한 host workset 열기 후 재시도"
                        ReportWeightedProgress(progress, hostTotal, hostIndex, 0.82R, $"[{hostIndex + 1}/{hostTotal}] 필요한 host workset 열기 후 재시도 중... {hostName}")
                        SafeClose(doc)
                        doc = OpenProjectDocumentForApply(app, openPath, retryWorksetNames, includeSavedOpenWorksets:=Not wasCentralFile)
                        If doc Is Nothing Then
                            MarkRows(retryRows, "error", "재시도용 호스트 문서를 열지 못했습니다.")
                            Return
                        End If

                        For retryIndex As Integer = 0 To retryRows.Count - 1
                            Dim row = retryRows(retryIndex)
                            Dim rowFraction As Double = 0.82R + (0.08R * ((CDbl(retryIndex) + 1.0R) / CDbl(Math.Max(1, retryRows.Count))))
                            Dim linkName As String = If(String.IsNullOrWhiteSpace(SafeStr(row.LinkName)),
                                                        SafeElementIdText(ParseElementIdOrInvalid(row.ReferenceElementId)),
                                                        row.LinkName)
                            Dim operationName As String = If(IsDeleteLinkRow(row),
                                                             "링크 삭제 재시도 중",
                                                             If(IsNewLinkRow(row), "신규 링크 생성 재시도 중", "링크 재시도 중"))
                            currentStage = operationName
                            ReportWeightedProgress(progress, hostTotal, hostIndex, rowFraction, $"[{hostIndex + 1}/{hostTotal}] {operationName} ({retryIndex + 1}/{retryRows.Count})... {linkName}")

                            Dim rowChanged As Boolean
                            If IsDeleteLinkRow(row) Then
                                rowChanged = DeleteLinkRow(doc, row)
                            ElseIf IsNewLinkRow(row) Then
                                rowChanged = CreateNewLinkRow(doc, row, options)
                            ElseIf allowTransmissionApply AndAlso CanApplyViaTransmission(row) Then
                                rowChanged = RefreshRowAfterTransmission(doc, row)
                            Else
                                rowChanged = ApplySingleRow(doc, row)
                            End If

                            If rowChanged AndAlso Not changedRows.Contains(row) Then
                                changedRows.Add(row)
                            End If
                        Next
                    End If
                End If

                If changedRows.Count = 0 Then
                    ReportWeightedProgress(progress, hostTotal, hostIndex, 0.98R, $"[{hostIndex + 1}/{hostTotal}] 변경할 링크가 없습니다... {hostName}")
                    Return
                End If

                If wasCentralFile Then
                    Dim syncUnavailableReason As String = ""
                    If Not CanSynchronizeWithCentral(doc, syncUnavailableReason) Then
                        MarkRows(changedRows, "error", "동기화 실패: " & syncUnavailableReason, overwriteChanged:=True)
                        Return
                    End If

                    currentStage = "센트럴 동기화"
                    ReportWeightedProgress(progress, hostTotal, hostIndex, 0.88R, $"[{hostIndex + 1}/{hostTotal}] 센트럴 동기화 중... {hostName}")
                    Dim syncError As String = ""
                    If SyncWithCentral(doc, "KKY Tools - Revit 링크 경로 재지정", syncError) Then
                        For Each row In changedRows
                            row.ApplyMessage = AppendMessage(row.ApplyMessage, "동기화 완료")
                        Next
                    Else
                        MarkRows(changedRows, "error", "동기화 실패: " & syncError, overwriteChanged:=True)
                    End If
                Else
                    currentStage = "RVT 저장"
                    ReportWeightedProgress(progress, hostTotal, hostIndex, 0.88R, $"[{hostIndex + 1}/{hostTotal}] RVT 저장 중... {hostName}")
                    doc.Save()
                    For Each row In changedRows
                        row.ApplyMessage = AppendMessage(row.ApplyMessage, "저장 완료")
                    Next
                End If

                currentStage = "마무리"
                ReportWeightedProgress(progress, hostTotal, hostIndex, 0.98R, $"[{hostIndex + 1}/{hostTotal}] 링크 적용 완료... {hostName}")
            Catch ex As Exception
                MarkRows(rows, "error", "호스트 적용 실패 (" & currentStage & "): " & ex.Message, overwriteChanged:=True)
                ReportWeightedProgress(progress, hostTotal, hostIndex, 0.98R, $"[{hostIndex + 1}/{hostTotal}] 적용 실패... {hostName}")
            Finally
                SafeClose(doc)
            End Try
        End Sub

        Private Shared Function ApplySingleRow(doc As Document,
                                               row As RevitLinkPathRow) As Boolean
            If doc Is Nothing OrElse row Is Nothing Then Return False

            Dim refInt As Integer
            If Not Integer.TryParse(NormalizeElementIdText(row.ReferenceElementId), refInt) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "링크 식별자를 해석하지 못했습니다."
                Return False
            End If

            Dim linkType As RevitLinkType = TryCast(doc.GetElement(New ElementId(refInt)), RevitLinkType)
            If linkType Is Nothing Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "대상 링크를 찾지 못했습니다."
                Return False
            End If

            Try
                If linkType.IsNestedLink Then
                    row.ApplyStatus = "skip"
                    row.ApplyMessage = "중첩 링크는 자동 변경 대상에서 제외합니다."
                    Return False
                End If
            Catch
            End Try

            Dim targetPath As String = NormalizeUserVisiblePath(row.TargetLinkPath)
            If String.IsNullOrWhiteSpace(targetPath) Then
                row.ApplyStatus = "skip"
                row.ApplyMessage = "대상 경로가 비어 있습니다."
                Return False
            End If

            If Not IsServerPath(targetPath) AndAlso Not File.Exists(targetPath) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "대상 링크 파일을 찾을 수 없습니다."
                Return False
            End If

            Dim currentVisiblePath As String = GetCurrentVisibleLinkPath(linkType, row)
            Dim currentStoredPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            If Not ShouldBypassOriginalLinkReferenceCheck(row, linkType, currentVisiblePath, currentStoredPath) AndAlso
               Not IsOriginalLinkReferenceStillCurrent(row, currentVisiblePath, currentStoredPath, linkType.Name) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "엑셀의 원본 링크 경로와 현재 RVT의 링크 경로가 달라 건너뜁니다. 최신 상태로 다시 추출해 주세요."
                Return False
            End If

            If String.Equals(NormalizeComparePath(currentVisiblePath), NormalizeComparePath(targetPath), StringComparison.OrdinalIgnoreCase) Then
                row.TargetPathType = ResolvePathType(targetPath).ToString()
                row.ApplyStatus = "skip"
                row.ApplyMessage = "현재 링크 경로와 같아서 건너뜁니다."
                Return False
            End If

            Dim checkoutError As String = ""
            If Not EnsureEditableHostWorksets(doc, linkType.Id, checkoutError) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = checkoutError
                Return False
            End If

            Dim targetModelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(targetPath)
            Dim targetPathType As PathType = ResolvePathType(targetPath)
            Dim config As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            Dim loadResultText As String = ""
            Dim loadErrorText As String = ""
            Dim retryApplied As Boolean = False
            Dim currentWasCloud As Boolean = IsCloudLinkType(linkType, currentVisiblePath, currentStoredPath) OrElse
                                             String.Equals(SafeStr(row.CurrentPathType), "Cloud", StringComparison.OrdinalIgnoreCase)
            Dim loadSucceeded As Boolean = TryReloadLinkFromTarget(doc, linkType, targetModelPath, config, loadResultText, loadErrorText, retryApplied)
            If Not loadSucceeded AndAlso currentWasCloud Then
                loadSucceeded = TryConvertCloudLinkToExternalFile(doc, linkType, row, targetModelPath, targetPathType, config, loadResultText, loadErrorText)
            End If

            If Not loadSucceeded Then
                row.ApplyStatus = "error"
                Dim failText As String = If(Not String.IsNullOrWhiteSpace(loadErrorText), loadErrorText, loadResultText)
                row.ApplyMessage = If(String.IsNullOrWhiteSpace(failText),
                                      "Reload From 처리에 실패했습니다.",
                                      "Reload From 처리에 실패했습니다. (" & failText & ")")
                Return False
            End If

            row.TargetPathType = targetPathType.ToString()
            Try
                Dim refreshedTypeId As ElementId = ParseElementIdOrInvalid(row.ReferenceElementId)
                If refreshedTypeId Is Nothing OrElse refreshedTypeId = ElementId.InvalidElementId Then
                    refreshedTypeId = linkType.Id
                End If
                Dim refreshedLinkType As RevitLinkType = TryCast(doc.GetElement(refreshedTypeId), RevitLinkType)
                If refreshedLinkType IsNot Nothing Then
                    linkType = refreshedLinkType
                End If
            Catch
            End Try
            row.CurrentLinkPath = GetCurrentVisibleLinkPath(linkType, row)
            row.StoredLinkPath = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            row.CurrentPathType = SafePathTypeText(linkType)
            If Not IsResolvedToTargetPath(row.CurrentLinkPath, row.StoredLinkPath, targetPath, targetPathType) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "Reload From 후 실제 링크 경로가 대상 경로로 바뀌지 않았습니다."
                Return False
            End If
            row.LinkFileName = SafeFileName(If(Not String.IsNullOrWhiteSpace(row.CurrentLinkPath), row.CurrentLinkPath, targetPath))
            row.TypeWorksetNames = GetLinkTypeWorksetNamesText(doc, linkType.Id)
            row.InstanceWorksetNames = GetLinkInstanceWorksetNamesText(doc, linkType.Id)
            If String.IsNullOrWhiteSpace(SafeStr(row.ApplyTypeWorksetNames)) Then
                row.ApplyTypeWorksetNames = row.TypeWorksetNames
            End If
            If String.IsNullOrWhiteSpace(SafeStr(row.ApplyInstanceWorksetNames)) Then
                row.ApplyInstanceWorksetNames = row.InstanceWorksetNames
            End If
            row.ApplyStatus = "changed"
            If currentWasCloud AndAlso Not String.Equals(row.CurrentPathType, "Cloud", StringComparison.OrdinalIgnoreCase) Then
                row.ApplyMessage = "Cloud 링크를 일반 파일 링크로 전환 완료"
            Else
                row.ApplyMessage = If(retryApplied, "닫힌 웍셋 링크를 재활성화한 뒤 Reload From 완료", "Reload From 완료")
            End If
            Return True
        End Function

        Private Shared Function TryConvertCloudLinkToExternalFile(doc As Document,
                                                                  oldLinkType As RevitLinkType,
                                                                  row As RevitLinkPathRow,
                                                                  targetModelPath As ModelPath,
                                                                  targetPathType As PathType,
                                                                  config As WorksetConfiguration,
                                                                  ByRef loadResultText As String,
                                                                  ByRef loadErrorText As String) As Boolean
            loadResultText = ""
            If doc Is Nothing OrElse oldLinkType Is Nothing OrElse row Is Nothing OrElse targetModelPath Is Nothing Then Return False

            Dim oldTypeId As ElementId = oldLinkType.Id
            Dim oldTypeWorksetName As String = ResolveWorksetName(doc, oldLinkType.WorksetId)
            Dim instances As List(Of RevitLinkInstance) = CollectLinkInstances(doc, oldTypeId)
            Dim tx As Transaction = Nothing

            Try
                tx = New Transaction(doc, "KKY Tools - Cloud 링크를 파일 링크로 전환")
                If tx.Start() <> TransactionStatus.Started Then
                    loadErrorText = "Cloud 링크 전환 트랜잭션을 시작하지 못했습니다."
                    Return False
                End If

                Dim isRelative As Boolean = (targetPathType = PathType.Relative)
                Dim linkOptions As New RevitLinkOptions(isRelative, config)
                Dim createResult As LinkLoadResult = RevitLinkType.Create(doc, targetModelPath, linkOptions)
                If Not EvaluateLinkLoadResult(createResult, loadResultText) Then
                    RollBackTransaction(tx)
                    Return False
                End If

                Dim newTypeId As ElementId = If(createResult Is Nothing, ElementId.InvalidElementId, createResult.ElementId)
                Dim newLinkType As RevitLinkType = TryCast(doc.GetElement(newTypeId), RevitLinkType)
                If newLinkType Is Nothing Then
                    loadErrorText = "전환용 새 링크 타입을 찾지 못했습니다."
                    RollBackTransaction(tx)
                    Return False
                End If

                If Not String.IsNullOrWhiteSpace(oldTypeWorksetName) Then
                    SetElementWorksetByName(newLinkType, oldTypeWorksetName)
                End If

                For Each inst In instances
                    If inst Is Nothing Then Continue For
                    inst.ChangeTypeId(newLinkType.Id)
                Next

                doc.Delete(oldTypeId)
                tx.Commit()

                row.ReferenceElementId = SafeElementIdText(newLinkType.Id)
                loadResultText = "Cloud 링크를 일반 파일 링크로 전환 완료"
                loadErrorText = ""
                Return True
            Catch ex As Exception
                loadErrorText = ex.Message
                RollBackTransaction(tx)
                Return False
            Finally
                If tx IsNot Nothing Then
                    Try
                        tx.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Shared Function DeleteLinkRow(doc As Document,
                                              row As RevitLinkPathRow) As Boolean
            If doc Is Nothing OrElse row Is Nothing Then Return False

            Dim linkType As RevitLinkType = ResolveExistingLinkType(doc, row)
            If linkType Is Nothing Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "삭제할 링크를 찾지 못했습니다."
                Return False
            End If

            Try
                If linkType.IsNestedLink Then
                    row.ApplyStatus = "skip"
                    row.ApplyMessage = "중첩 링크는 자동 삭제 대상에서 제외합니다."
                    Return False
                End If
            Catch
            End Try

            Dim currentVisiblePath As String = GetCurrentVisibleLinkPath(linkType, row)
            Dim storedPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            If Not ShouldBypassOriginalLinkReferenceCheck(row, linkType, currentVisiblePath, storedPath) AndAlso
               Not IsOriginalLinkReferenceStillCurrent(row, currentVisiblePath, storedPath, linkType.Name) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "엑셀의 원본 링크 경로와 현재 RVT의 링크 경로가 달라 삭제를 건너뜁니다. 최신 상태로 다시 추출해 주세요."
                Return False
            End If

            Dim preferredWorksetNames = CollectApplyWorksetNames(row, doc)
            Dim namedCheckoutError As String = ""
            If preferredWorksetNames IsNot Nothing AndAlso preferredWorksetNames.Any() AndAlso
               Not EnsureEditableNamedWorksets(doc, preferredWorksetNames, namedCheckoutError) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = namedCheckoutError
                Return False
            End If

            Dim checkoutError As String = ""
            If Not EnsureEditableHostWorksets(doc, linkType.Id, checkoutError) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = checkoutError
                Return False
            End If

            Dim deletePreparationNote As String = ""
            TryPrepareLinkTypeForDeletion(doc, linkType, deletePreparationNote)

            Dim tx As Transaction = Nothing
            Try
                tx = New Transaction(doc, "KKY Tools - Revit 링크 삭제")
                If tx.Start() <> TransactionStatus.Started Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "링크 삭제 트랜잭션을 시작하지 못했습니다."
                    Return False
                End If

                Dim linkTypeId As ElementId = linkType.Id
                Dim instanceIds As List(Of ElementId) =
                    CollectLinkInstances(doc, linkTypeId).
                        Where(Function(x) x IsNot Nothing).
                        Select(Function(x) x.Id).
                        ToList()

                If instanceIds.Count > 0 Then
                    doc.Delete(instanceIds)
                    Try
                        doc.Regenerate()
                    Catch
                    End Try
                End If

                If doc.GetElement(linkTypeId) IsNot Nothing Then
                    doc.Delete(linkTypeId)
                    Try
                        doc.Regenerate()
                    Catch
                    End Try
                End If

                If doc.GetElement(linkTypeId) IsNot Nothing OrElse CollectLinkInstances(doc, linkTypeId).Count > 0 Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "링크 삭제 후에도 링크 타입 또는 인스턴스가 남아 있습니다."
                    RollBackTransaction(tx)
                    Return False
                End If

                tx.Commit()

                row.ReferenceElementId = ""
                row.CurrentLinkPath = ""
                row.StoredLinkPath = ""
                row.CurrentPathType = ""
                row.TargetPathType = ""
                row.TypeWorksetNames = ""
                row.InstanceWorksetNames = ""
                row.ApplyStatus = "changed"
                row.ApplyMessage = "TargetLinkPath가 비어 있어 기존 Revit 링크 삭제 완료"
                If Not String.IsNullOrWhiteSpace(deletePreparationNote) Then
                    row.ApplyMessage = AppendMessage(row.ApplyMessage, deletePreparationNote)
                End If
                Return True
            Catch ex As Exception
                RollBackTransaction(tx)
                row.ApplyStatus = "error"
                row.ApplyMessage = "링크 삭제 실패: " & ex.Message
                Return False
            Finally
                If tx IsNot Nothing Then
                    Try
                        tx.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Shared Function ResolveExistingLinkType(doc As Document,
                                                        row As RevitLinkPathRow) As RevitLinkType
            If doc Is Nothing OrElse row Is Nothing Then Return Nothing

            Dim refId As ElementId = ParseElementIdOrInvalid(row.ReferenceElementId)
            If refId IsNot Nothing AndAlso refId <> ElementId.InvalidElementId Then
                Dim directType As RevitLinkType = TryCast(doc.GetElement(refId), RevitLinkType)
                If directType IsNot Nothing Then Return directType

                Dim linkInstance As RevitLinkInstance = TryCast(doc.GetElement(refId), RevitLinkInstance)
                If linkInstance IsNot Nothing Then
                    Dim instanceType As RevitLinkType = TryCast(doc.GetElement(linkInstance.GetTypeId()), RevitLinkType)
                    If instanceType IsNot Nothing Then Return instanceType
                End If
            End If

            Dim linkTypes As List(Of RevitLinkType) =
                New FilteredElementCollector(doc).
                    OfClass(GetType(RevitLinkType)).
                    Cast(Of RevitLinkType)().
                    Where(Function(x) x IsNot Nothing AndAlso Not x.IsNestedLink).
                    ToList()
            If linkTypes.Count = 0 Then Return Nothing

            Dim expectedCurrent As String = NormalizeUserVisiblePath(row.CurrentLinkPath)
            Dim expectedStored As String = NormalizeUserVisiblePath(row.StoredLinkPath)
            Dim expectedName As String = SafeStr(row.LinkName).Trim()

            If Not String.IsNullOrWhiteSpace(expectedCurrent) OrElse Not String.IsNullOrWhiteSpace(expectedStored) Then
                Dim matchesByPath = linkTypes.
                    Where(Function(x)
                              Dim actualCurrent As String = GetCurrentVisibleLinkPath(x, row)
                              Dim actualStored As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(x))
                              Return PathsMatch(expectedCurrent, actualCurrent) OrElse
                                     PathsMatch(expectedCurrent, actualStored) OrElse
                                     PathsMatch(expectedStored, actualCurrent) OrElse
                                     PathsMatch(expectedStored, actualStored)
                          End Function).
                    ToList()

                If matchesByPath.Count > 1 AndAlso Not String.IsNullOrWhiteSpace(expectedName) Then
                    matchesByPath =
                        matchesByPath.
                            Where(Function(x) String.Equals(SafeStr(x.Name).Trim(),
                                                            expectedName,
                                                            StringComparison.OrdinalIgnoreCase)).
                            ToList()
                End If

                If matchesByPath.Count = 1 Then
                    Return matchesByPath(0)
                End If
            End If

            If Not String.IsNullOrWhiteSpace(expectedName) Then
                Dim matchesByName = linkTypes.
                    Where(Function(x) String.Equals(SafeStr(x.Name).Trim(),
                                                    expectedName,
                                                    StringComparison.OrdinalIgnoreCase)).
                    ToList()
                If matchesByName.Count = 1 Then
                    Return matchesByName(0)
                End If
            End If

            Return Nothing
        End Function

        Private Shared Function CreateNewLinkRow(doc As Document,
                                                 row As RevitLinkPathRow,
                                                 options As RevitLinkPathApplyOptions) As Boolean
            If doc Is Nothing OrElse row Is Nothing Then Return False

            Dim targetPath As String = NormalizeUserVisiblePath(row.TargetLinkPath)
            If String.IsNullOrWhiteSpace(targetPath) Then
                row.ApplyStatus = "skip"
                row.ApplyMessage = "신규 링크 대상 경로가 비어 있습니다."
                Return False
            End If

            If IsLikelyCloudPath(targetPath) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "ACC/BIM 360 클라우드 경로는 신규 링크 자동 생성에서 제외합니다."
                Return False
            End If

            If Not IsServerPath(targetPath) AndAlso Not File.Exists(targetPath) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "신규 링크 파일을 찾을 수 없습니다."
                Return False
            End If

            Dim typeWorksetName As String = ResolveFirstWorksetName(If(Not String.IsNullOrWhiteSpace(SafeStr(row.ApplyTypeWorksetNames)),
                                                                        row.ApplyTypeWorksetNames,
                                                                        row.TypeWorksetNames))
            Dim instanceWorksetName As String = ResolveFirstWorksetName(If(Not String.IsNullOrWhiteSpace(SafeStr(row.ApplyInstanceWorksetNames)),
                                                                            row.ApplyInstanceWorksetNames,
                                                                            row.InstanceWorksetNames))
            If String.IsNullOrWhiteSpace(instanceWorksetName) Then instanceWorksetName = typeWorksetName

            Dim checkoutError As String = ""
            If Not EnsureEditableNamedWorksets(doc, New String() {typeWorksetName, instanceWorksetName}, checkoutError) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = checkoutError
                Return False
            End If

            Dim targetPathType As PathType = ResolvePathType(targetPath)
            Dim targetModelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(targetPath)
            Dim safeOptions As RevitLinkPathApplyOptions = If(options, New RevitLinkPathApplyOptions())
            Dim requestedPlacement As ImportPlacement = ResolveNewLinkPlacement(safeOptions.NewLinkPlacement)
            Dim loadResultText As String = ""

            Dim tx As Transaction = Nothing
            Try
                tx = New Transaction(doc, "KKY Tools - 신규 Revit 링크 추가")
                If tx.Start() <> TransactionStatus.Started Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "신규 링크 생성 트랜잭션을 시작하지 못했습니다."
                    Return False
                End If

                Dim linkOptions As New RevitLinkOptions(targetPathType = PathType.Relative, New WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets))
                Dim loadResult As LinkLoadResult = RevitLinkType.Create(doc, targetModelPath, linkOptions)
                If Not EvaluateLinkLoadResult(loadResult, loadResultText) Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = If(String.IsNullOrWhiteSpace(loadResultText),
                                          "신규 링크 타입을 만들지 못했습니다.",
                                          "신규 링크 타입을 만들지 못했습니다. (" & loadResultText & ")")
                    RollBackTransaction(tx)
                    Return False
                End If

                Dim linkTypeId As ElementId = If(loadResult Is Nothing, ElementId.InvalidElementId, loadResult.ElementId)
                Dim linkType As RevitLinkType = TryCast(doc.GetElement(linkTypeId), RevitLinkType)
                If linkType Is Nothing Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "신규 링크 타입을 다시 찾지 못했습니다."
                    RollBackTransaction(tx)
                    Return False
                End If

                Dim instance As RevitLinkInstance = CreateLinkInstanceForPlacement(doc, linkType.Id, requestedPlacement)
                If instance Is Nothing Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "신규 링크 인스턴스를 만들지 못했습니다."
                    RollBackTransaction(tx)
                    Return False
                End If

                SetElementWorksetByName(linkType, typeWorksetName)
                SetElementWorksetByName(instance, instanceWorksetName)

                Try
                    doc.Regenerate()
                Catch
                End Try

                tx.Commit()

                Dim instances As List(Of RevitLinkInstance) = CollectLinkInstances(doc, linkType.Id)
                Dim placementNote As String = ""
                Dim usedProjectBasePointFallback As Boolean = False
                Dim placementHealthy As Boolean = NormalizeExistingLinkPlacement(doc, linkType.Id, requestedPlacement, placementNote)
                If Not HasDrawableLinkInstance(doc, linkType.Id) OrElse Not placementHealthy Then
                    Dim fallbackError As String = ""
                    If Not RecreateLinkInstanceAtProjectBasePoint(doc, linkType.Id, instanceWorksetName, fallbackError) Then
                        row.ApplyStatus = "error"
                        row.ApplyMessage = If(String.IsNullOrWhiteSpace(fallbackError),
                                              "신규 링크 인스턴스가 실제 문서에 생성되지 않았습니다.",
                                              fallbackError)
                        Return False
                    End If

                    usedProjectBasePointFallback = True
                    instances = CollectLinkInstances(doc, linkType.Id)
                    Dim fallbackPlacementNote As String = ""
                    NormalizeExistingLinkPlacement(doc, linkType.Id, ImportPlacement.Site, fallbackPlacementNote)
                    placementNote = AppendMessage(placementNote, fallbackPlacementNote)
                End If

                row.ReferenceElementId = SafeElementIdText(linkType.Id)
                row.LinkName = SafeStr(linkType.Name)
                row.LinkFileName = SafeFileName(targetPath)
                row.CurrentLinkPath = GetCurrentVisibleLinkPath(linkType, row)
                If String.IsNullOrWhiteSpace(row.CurrentLinkPath) Then row.CurrentLinkPath = targetPath
                row.StoredLinkPath = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
                row.CurrentPathType = SafePathTypeText(linkType)
                row.TargetPathType = targetPathType.ToString()
                row.TypeWorksetNames = GetLinkTypeWorksetNamesText(doc, linkType.Id)
                row.InstanceWorksetNames = GetLinkInstanceWorksetNamesText(doc, linkType.Id)
                If String.IsNullOrWhiteSpace(SafeStr(row.ApplyTypeWorksetNames)) Then row.ApplyTypeWorksetNames = row.TypeWorksetNames
                If String.IsNullOrWhiteSpace(SafeStr(row.ApplyInstanceWorksetNames)) Then row.ApplyInstanceWorksetNames = row.InstanceWorksetNames
                row.ApplyStatus = "changed"
                row.ApplyMessage = "신규 Revit 링크 생성 완료 (" & requestedPlacement.ToString() & ")"
                If usedProjectBasePointFallback Then
                    row.ApplyMessage = AppendMessage(row.ApplyMessage, "인스턴스가 확인되지 않아 프로젝트 기준점으로 다시 배치")
                End If
                If Not String.IsNullOrWhiteSpace(placementNote) Then
                    row.ApplyMessage = AppendMessage(row.ApplyMessage, placementNote)
                End If
                Return True
            Catch ex As Exception
                RollBackTransaction(tx)
                row.ApplyStatus = "error"
                row.ApplyMessage = "신규 링크 생성 실패: " & ex.Message
                Return False
            Finally
                If tx IsNot Nothing Then
                    Try
                        tx.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Shared Function CreateLinkInstanceForPlacement(doc As Document,
                                                               linkTypeId As ElementId,
                                                               requestedPlacement As ImportPlacement) As RevitLinkInstance
            If doc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then Return Nothing

            Dim instance As RevitLinkInstance = Nothing
            Select Case requestedPlacement
                Case ImportPlacement.Shared
                    instance = RevitLinkInstance.Create(doc, linkTypeId, ImportPlacement.Shared)
                Case ImportPlacement.Site
                    instance = RevitLinkInstance.Create(doc, linkTypeId)
                Case ImportPlacement.Centered
                    instance = RevitLinkInstance.Create(doc, linkTypeId, ImportPlacement.Centered)
                Case Else
                    instance = RevitLinkInstance.Create(doc, linkTypeId)
            End Select

            If instance Is Nothing Then Return Nothing

            Try
                Dim linkType As RevitLinkType = TryCast(doc.GetElement(linkTypeId), RevitLinkType)
                If linkType IsNot Nothing Then
                    Dim loadError As String = ""
                    PrimeLinkTypeForReload(doc, linkType, loadError)
                End If
            Catch
            End Try

            Try
                doc.Regenerate()
            Catch
            End Try

            Try
                If instance.Pinned Then
                    instance.Pinned = False
                End If
            Catch
            End Try

            Try
                Select Case requestedPlacement
                    Case ImportPlacement.Site
                        instance.MoveBasePointToHostBasePoint(False)
                    Case ImportPlacement.Origin
                        instance.MoveOriginToHostOrigin(False)
                End Select
            Catch
            End Try

            Try
                doc.Regenerate()
            Catch
            End Try

            Return instance
        End Function

        Private Shared Function RecreateLinkInstanceAtProjectBasePoint(doc As Document,
                                                                       linkTypeId As ElementId,
                                                                       instanceWorksetName As String,
                                                                       ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If doc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then
                errorMessage = "프로젝트 기준점 재배치 대상 링크를 찾지 못했습니다."
                Return False
            End If

            Dim tx As Transaction = Nothing
            Try
                tx = New Transaction(doc, "KKY Tools - 프로젝트 기준점 재배치")
                If tx.Start() <> TransactionStatus.Started Then
                    errorMessage = "프로젝트 기준점 재배치 트랜잭션을 시작하지 못했습니다."
                    Return False
                End If

                Dim existingIds As List(Of ElementId) =
                    CollectLinkInstances(doc, linkTypeId).
                        Where(Function(x) x IsNot Nothing).
                        Select(Function(x) x.Id).
                        ToList()
                If existingIds.Count > 0 Then
                    doc.Delete(existingIds)
                End If

                Dim repairedInstance As RevitLinkInstance = CreateLinkInstanceForPlacement(doc, linkTypeId, ImportPlacement.Site)
                If repairedInstance Is Nothing Then
                    errorMessage = "프로젝트 기준점으로 신규 링크 인스턴스를 만들지 못했습니다."
                    RollBackTransaction(tx)
                    Return False
                End If

                SetElementWorksetByName(repairedInstance, instanceWorksetName)

                Try
                    doc.Regenerate()
                Catch
                End Try

                tx.Commit()

                If Not HasDrawableLinkInstance(doc, linkTypeId) Then
                    errorMessage = "프로젝트 기준점으로 다시 배치해도 링크 인스턴스가 확인되지 않습니다."
                    Return False
                End If

                Return True
            Catch ex As Exception
                RollBackTransaction(tx)
                errorMessage = "프로젝트 기준점 재배치 실패: " & ex.Message
                Return False
            Finally
                If tx IsNot Nothing Then
                    Try
                        tx.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Shared Function HasDrawableLinkInstance(doc As Document,
                                                        linkTypeId As ElementId) As Boolean
            If doc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then Return False

            For Each instance In CollectLinkInstances(doc, linkTypeId)
                If instance Is Nothing Then Continue For

                Try
                    If Not instance.IsValidObject Then Continue For
                Catch
                End Try

                Try
                    Dim bbox As BoundingBoxXYZ = instance.BoundingBox(Nothing)
                    If bbox IsNot Nothing Then Return True
                Catch
                End Try

                Try
                    If instance.GetLinkDocument() IsNot Nothing Then Return True
                Catch
                End Try
            Next

            Return False
        End Function

        Private Shared Function NormalizeExistingLinkPlacement(doc As Document,
                                                               linkTypeId As ElementId,
                                                               requestedPlacement As ImportPlacement,
                                                               ByRef note As String) As Boolean
            note = ""
            If doc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then Return False

            If requestedPlacement <> ImportPlacement.Origin AndAlso requestedPlacement <> ImportPlacement.Site Then
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            Dim expectedOrigin As XYZ = Nothing
            If Not TryGetExpectedLinkOrigin(doc, requestedPlacement, expectedOrigin) Then
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            Dim instances As List(Of RevitLinkInstance) = CollectLinkInstances(doc, linkTypeId)
            If instances.Count = 0 Then Return False

            Dim currentOrigin As XYZ = Nothing
            If Not TryGetLinkPlacementPoint(instances(0), requestedPlacement, currentOrigin) Then
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            Dim tolerance As Double = 1.0R
            Dim initialDistance As Double = currentOrigin.DistanceTo(expectedOrigin)
            If initialDistance <= tolerance Then
                note = "현재 배치 원점 " & FormatXyz(currentOrigin)
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            Dim tx As Transaction = Nothing
            Try
                tx = New Transaction(doc, "KKY Tools - 신규 링크 위치 보정")
                If tx.Start() <> TransactionStatus.Started Then
                    note = "위치 보정 트랜잭션을 시작하지 못했습니다."
                    Return False
                End If

                For Each instance In instances
                    If instance Is Nothing Then Continue For
                    EnsureInstanceUnpinned(instance)
                    Try
                        Select Case requestedPlacement
                            Case ImportPlacement.Site
                                instance.MoveBasePointToHostBasePoint(False)
                            Case ImportPlacement.Origin
                                instance.MoveOriginToHostOrigin(False)
                        End Select
                    Catch
                    End Try
                Next

                Try
                    doc.Regenerate()
                Catch
                End Try

                instances = CollectLinkInstances(doc, linkTypeId)
                If instances.Count > 0 AndAlso TryGetLinkPlacementPoint(instances(0), requestedPlacement, currentOrigin) Then
                    Dim adjustedDistance As Double = currentOrigin.DistanceTo(expectedOrigin)
                    If adjustedDistance > tolerance Then
                        Dim delta As New XYZ(expectedOrigin.X - currentOrigin.X,
                                             expectedOrigin.Y - currentOrigin.Y,
                                             expectedOrigin.Z - currentOrigin.Z)
                        For Each instance In instances
                            If instance Is Nothing Then Continue For
                            ElementTransformUtils.MoveElement(doc, instance.Id, delta)
                        Next
                    End If
                End If

                Try
                    doc.Regenerate()
                Catch
                End Try

                tx.Commit()
            Catch
                RollBackTransaction(tx)
                Return False
            Finally
                If tx IsNot Nothing Then
                    Try
                        tx.Dispose()
                    Catch
                    End Try
                End If
            End Try

            instances = CollectLinkInstances(doc, linkTypeId)
            If instances.Count = 0 Then Return False
            If Not TryGetLinkPlacementPoint(instances(0), requestedPlacement, currentOrigin) Then
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            Dim finalDistance As Double = currentOrigin.DistanceTo(expectedOrigin)
            If finalDistance <= tolerance Then
                note = "위치 보정 적용 " & FormatXyz(currentOrigin)
                Return HasDrawableLinkInstance(doc, linkTypeId)
            End If

            note = "위치 보정 후 원점 " & FormatXyz(currentOrigin)
            Return False
        End Function

        Private Shared Function TryPrepareLinkTypeForDeletion(doc As Document,
                                                              linkType As RevitLinkType,
                                                              ByRef note As String) As Boolean
            note = ""
            If doc Is Nothing OrElse linkType Is Nothing Then Return False

            Dim isLoaded As Boolean = False
            Try
                isLoaded = RevitLinkType.IsLoaded(doc, linkType.Id)
            Catch
                Return True
            End Try

            If Not isLoaded Then Return True

            Try
                If Not linkType.IsNotLoadedIntoMultipleOpenDocuments() Then
                    note = "삭제 전 unload는 건너뜀(여러 열린 문서에서 사용 중)"
                    Return True
                End If
            Catch
            End Try

            Try
                linkType.Unload(Nothing)
                note = "삭제 전 링크 unload 적용"
                Return True
            Catch ex As Exception
                note = "삭제 전 unload 실패: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Sub EnsureInstanceUnpinned(instance As RevitLinkInstance)
            If instance Is Nothing Then Return
            Try
                If instance.Pinned Then
                    instance.Pinned = False
                End If
            Catch
            End Try
        End Sub

        Private Shared Function TryGetLinkInstanceOrigin(instance As RevitLinkInstance,
                                                         ByRef origin As XYZ) As Boolean
            origin = Nothing
            If instance Is Nothing Then Return False

            Try
                Dim totalTransform As Transform = instance.GetTotalTransform()
                If totalTransform IsNot Nothing AndAlso totalTransform.Origin IsNot Nothing Then
                    origin = totalTransform.Origin
                    Return True
                End If
            Catch
            End Try

            Try
                Dim transform As Transform = instance.GetTransform()
                If transform IsNot Nothing AndAlso transform.Origin IsNot Nothing Then
                    origin = transform.Origin
                    Return True
                End If
            Catch
            End Try

            Try
                Dim locPoint As LocationPoint = TryCast(instance.Location, LocationPoint)
                If locPoint IsNot Nothing AndAlso locPoint.Point IsNot Nothing Then
                    origin = locPoint.Point
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        Private Shared Function TryGetLinkPlacementPoint(instance As RevitLinkInstance,
                                                         requestedPlacement As ImportPlacement,
                                                         ByRef point As XYZ) As Boolean
            point = Nothing
            If instance Is Nothing Then Return False

            Select Case requestedPlacement
                Case ImportPlacement.Site
                    Dim linkProjectBasePoint As XYZ = Nothing
                    Try
                        Dim linkDoc As Document = instance.GetLinkDocument()
                        If linkDoc IsNot Nothing AndAlso TryGetHostProjectBasePointPosition(linkDoc, linkProjectBasePoint) Then
                            Dim totalTransform As Transform = Nothing
                            Try
                                totalTransform = instance.GetTotalTransform()
                            Catch
                                totalTransform = Nothing
                            End Try
                            If totalTransform Is Nothing Then
                                Try
                                    totalTransform = instance.GetTransform()
                                Catch
                                    totalTransform = Nothing
                                End Try
                            End If
                            If totalTransform IsNot Nothing Then
                                point = totalTransform.OfPoint(linkProjectBasePoint)
                                Return True
                            End If
                        End If
                    Catch
                    End Try
            End Select

            Return TryGetLinkInstanceOrigin(instance, point)
        End Function

        Private Shared Function TryGetExpectedLinkOrigin(doc As Document,
                                                         requestedPlacement As ImportPlacement,
                                                         ByRef origin As XYZ) As Boolean
            origin = Nothing
            If doc Is Nothing Then Return False

            Select Case requestedPlacement
                Case ImportPlacement.Origin
                    origin = New XYZ(0, 0, 0)
                    Return True
                Case ImportPlacement.Site
                    Return TryGetHostProjectBasePointPosition(doc, origin)
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function TryGetHostProjectBasePointPosition(doc As Document,
                                                                   ByRef origin As XYZ) As Boolean
            origin = Nothing
            If doc Is Nothing Then Return False

            Try
                Dim basePointType As Type = GetType(Element).Assembly.GetType("Autodesk.Revit.DB.BasePoint")
                If basePointType Is Nothing Then Return False

                Dim getter = basePointType.GetMethod("GetProjectBasePoint", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If getter IsNot Nothing Then
                    Dim basePointObj As Object = getter.Invoke(Nothing, New Object() {doc})
                    origin = ReadBasePointPosition(basePointObj)
                    If origin IsNot Nothing Then Return True
                End If

                Dim candidates = New FilteredElementCollector(doc).OfClass(basePointType).ToElements()
                For Each candidate In candidates
                    Dim isShared As Boolean? = ReadNullableBooleanProperty(candidate, "IsShared")
                    If isShared.HasValue AndAlso isShared.Value Then Continue For

                    origin = ReadBasePointPosition(candidate)
                    If origin IsNot Nothing Then Return True
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function ReadBasePointPosition(basePointObj As Object) As XYZ
            If basePointObj Is Nothing Then Return Nothing

            Try
                Dim prop = basePointObj.GetType().GetProperty("Position", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                If prop IsNot Nothing Then
                    Return TryCast(prop.GetValue(basePointObj, Nothing), XYZ)
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function ReadNullableBooleanProperty(target As Object,
                                                            propertyName As String) As Boolean?
            If target Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return Nothing

            Try
                Dim prop = target.GetType().GetProperty(propertyName, System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                If prop Is Nothing Then Return Nothing
                Dim valueObj As Object = prop.GetValue(target, Nothing)
                If valueObj Is Nothing Then Return Nothing
                Return Convert.ToBoolean(valueObj, CultureInfo.InvariantCulture)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function FormatXyz(point As XYZ) As String
            If point Is Nothing Then Return "(n/a)"
            Return String.Format(CultureInfo.InvariantCulture,
                                 "({0:0.###}, {1:0.###}, {2:0.###})",
                                 point.X,
                                 point.Y,
                                 point.Z)
        End Function

        Private Shared Sub RollBackTransaction(tx As Transaction)
            If tx Is Nothing Then Return
            Try
                If tx.GetStatus() = TransactionStatus.Started Then
                    tx.RollBack()
                End If
            Catch
            End Try
        End Sub

        Private Shared Function OpenProjectDocument(app As RevitApp,
                                                    userVisiblePath As String,
                                                    closeAllWorksets As Boolean) As Document
            Dim modelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)
            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If closeAllWorksets Then
                openOpts.SetOpenWorksetsConfiguration(New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets))
            End If

            Return app.OpenDocumentFile(modelPath, openOpts)
        End Function

        Private Shared Function OpenProjectDocumentForApply(app As RevitApp,
                                                            userVisiblePath As String,
                                                            Optional additionalOpenWorksetNames As IEnumerable(Of String) = Nothing,
                                                            Optional includeSavedOpenWorksets As Boolean = True) As Document
            Dim modelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)
            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
            openOpts.SetOpenWorksetsConfiguration(BuildSavedWorksetOpenConfiguration(modelPath, additionalOpenWorksetNames, includeSavedOpenWorksets))
            Return app.OpenDocumentFile(modelPath, openOpts)
        End Function

        Private Shared Function SyncWithCentral(doc As Document,
                                                comment As String,
                                                ByRef err As String) As Boolean
            err = ""
            If doc Is Nothing Then
                err = "문서를 찾을 수 없습니다."
                Return False
            End If

            Try
                Dim twc As New TransactWithCentralOptions()
                Dim swc As New SynchronizeWithCentralOptions()
                swc.Comment = SafeStr(comment)
                Try
                    swc.SetRelinquishOptions(New RelinquishOptions(True))
                Catch
                End Try
                doc.SynchronizeWithCentral(twc, swc)
                Return True
            Catch ex As Exception
                err = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function CanSynchronizeWithCentral(doc As Document,
                                                          ByRef reason As String) As Boolean
            reason = ""
            If doc Is Nothing Then
                reason = "문서를 찾을 수 없습니다."
                Return False
            End If

            Try
                If Not doc.IsWorkshared Then
                    reason = "worksharing이 활성화된 문서가 아닙니다."
                    Return False
                End If
            Catch ex As Exception
                reason = "worksharing 상태를 확인하지 못했습니다: " & ex.Message
                Return False
            End Try

            Try
                Dim centralModelPath As ModelPath = doc.GetWorksharingCentralModelPath()
                Dim centralPath As String = SafeModelPathToUserVisiblePath(centralModelPath)
                If String.IsNullOrWhiteSpace(centralPath) Then
                    reason = "이 문서는 central location이 없어 SyncWithCentral을 사용할 수 없습니다."
                    Return False
                End If
            Catch ex As Exception
                reason = ex.Message
                Return False
            End Try

            Return True
        End Function

        Private Shared Function CreateNewLocalPath(centralPath As String) As String
            Dim localRoot As String = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "LinkPath", DateTime.Now.ToString("yyyyMMdd"))
            Directory.CreateDirectory(localRoot)

            Dim fileName As String = Path.GetFileNameWithoutExtension(centralPath) & "_" & Environment.UserName & "_" & DateTime.Now.ToString("HHmmssfff") & ".rvt"
            Dim localPath As String = Path.Combine(localRoot, fileName)

            Dim sourcePath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(centralPath)
            Dim targetPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath)
            WorksharingUtils.CreateNewLocal(sourcePath, targetPath)

            Try
                If File.Exists(localPath) Then
                    Dim attrs As FileAttributes = File.GetAttributes(localPath)
                    If (attrs And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                        File.SetAttributes(localPath, attrs And Not FileAttributes.ReadOnly)
                    End If
                End If
            Catch
            End Try

            Return localPath
        End Function

        Private Shared Function BuildSavedWorksetOpenConfiguration(projectPath As ModelPath,
                                                                   Optional additionalOpenWorksetNames As IEnumerable(Of String) = Nothing,
                                                                   Optional includeSavedOpenWorksets As Boolean = True) As WorksetConfiguration
            Dim config As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
            If projectPath Is Nothing Then Return config

            Try
                Dim previews = WorksharingUtils.GetUserWorksetInfo(projectPath)
                If previews Is Nothing OrElse previews.Count = 0 Then Return config

                Dim extraNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                If additionalOpenWorksetNames IsNot Nothing Then
                    For Each name In additionalOpenWorksetNames
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        extraNames.Add(name.Trim())
                    Next
                End If

                Dim openIds As New List(Of WorksetId)()
                For Each preview In previews
                    If preview Is Nothing Then Continue For

                    Dim shouldOpen As Boolean = includeSavedOpenWorksets AndAlso IsPreviewMarkedOpen(preview)
                    If Not shouldOpen AndAlso extraNames.Count > 0 Then
                        Try
                            Dim previewName = Convert.ToString(preview.Name)
                            shouldOpen = extraNames.Contains(If(previewName, String.Empty))
                        Catch
                            shouldOpen = False
                        End Try
                    End If

                    If shouldOpen Then
                        openIds.Add(preview.Id)
                    End If
                Next

                If openIds.Count > 0 Then
                    config.Open(openIds)
                End If
            Catch
            End Try

            Return config
        End Function

        Private Shared Function IsPreviewMarkedOpen(preview As WorksetPreview) As Boolean
            If preview Is Nothing Then Return False
            Try
                Dim prop = preview.GetType().GetProperty("IsOpen")
                If prop Is Nothing Then Return False
                Dim raw = prop.GetValue(preview, Nothing)
                If TypeOf raw Is Boolean Then Return CBool(raw)
            Catch
            End Try
            Return False
        End Function

        Private Shared Function CanApplyViaTransmission(row As RevitLinkPathRow) As Boolean
            If row Is Nothing Then Return False
            If IsNewLinkRow(row) Then Return False

            Dim targetPath As String = NormalizeUserVisiblePath(row.TargetLinkPath)
            If String.IsNullOrWhiteSpace(targetPath) Then Return False
            If Not IsServerPath(targetPath) AndAlso Not File.Exists(targetPath) Then Return False
            If IsLikelyCloudPath(targetPath) Then Return False

            If IsLikelyCloudRow(row) Then Return False

            Return True
        End Function

        Private Shared Sub ApplyRowsViaTransmission(hostPath As String,
                                                    rows As IList(Of RevitLinkPathRow))
            If String.IsNullOrWhiteSpace(hostPath) OrElse rows Is Nothing OrElse rows.Count = 0 Then Return

            Dim modelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(hostPath)
            Dim transData As TransmissionData = TransmissionData.ReadTransmissionData(modelPath)
            If transData Is Nothing Then
                MarkRows(rows, "error", "TransmissionData를 읽지 못했습니다.")
                Return
            End If

            Dim changed As Boolean = False
            For Each row In rows
                If row Is Nothing Then Continue For

                Dim refInt As Integer
                If Not Integer.TryParse(NormalizeElementIdText(row.ReferenceElementId), refInt) Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "링크 식별자를 해석하지 못했습니다."
                    Continue For
                End If

                Dim refId As New ElementId(refInt)
                Dim extRef As ExternalFileReference = Nothing
                Try
                    extRef = transData.GetLastSavedReferenceData(refId)
                Catch
                    extRef = Nothing
                End Try
                If extRef Is Nothing Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "닫힌 문서 링크 정보를 찾지 못했습니다."
                    Continue For
                End If

                Try
                    If extRef.ExternalFileReferenceType <> ExternalFileReferenceType.RevitLink Then
                        row.ApplyStatus = "skip"
                        row.ApplyMessage = "외부 파일 링크가 아니어서 건너뜁니다."
                        Continue For
                    End If
                Catch
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "링크 유형을 확인하지 못했습니다."
                    Continue For
                End Try

                Dim targetPath As String = NormalizeUserVisiblePath(row.TargetLinkPath)
                If String.IsNullOrWhiteSpace(targetPath) Then
                    row.ApplyStatus = "skip"
                    row.ApplyMessage = "대상 경로가 비어 있습니다."
                    Continue For
                End If

                Dim storedPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(extRef))
                Dim absolutePath As String = SafeModelPathToUserVisiblePath(TryGetAbsoluteModelPath(extRef))
                Dim visiblePath As String = If(Not String.IsNullOrWhiteSpace(absolutePath), absolutePath, storedPath)
                If Not IsOriginalLinkReferenceStillCurrent(row, visiblePath, storedPath) Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "엑셀의 원본 링크 경로와 현재 RVT의 링크 경로가 달라 건너뜁니다. 최신 상태로 다시 추출해 주세요."
                    Continue For
                End If

                Dim targetPathType As PathType = ResolvePathType(targetPath)
                Dim targetModelPath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(targetPath)
                transData.SetDesiredReferenceData(refId, targetModelPath, targetPathType, True)
                row.TargetPathType = targetPathType.ToString()
                row.ApplyStatus = ""
                row.ApplyMessage = "닫힌 문서 경로 반영 대기"
                changed = True
            Next

            If Not changed Then Return

            transData.IsTransmitted = True
            TransmissionData.WriteTransmissionData(modelPath, transData)
        End Sub

        Private Shared Function RefreshRowAfterTransmission(doc As Document,
                                                            row As RevitLinkPathRow) As Boolean
            If doc Is Nothing OrElse row Is Nothing Then Return False

            Dim refInt As Integer
            If Not Integer.TryParse(NormalizeElementIdText(row.ReferenceElementId), refInt) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "링크 식별자를 해석하지 못했습니다."
                Return False
            End If

            Dim linkType As RevitLinkType = TryCast(doc.GetElement(New ElementId(refInt)), RevitLinkType)
            If linkType Is Nothing Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "대상 링크를 다시 확인하지 못했습니다."
                Return False
            End If

            Dim storedPath As String = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            If Not IsOriginalLinkReferenceStillCurrent(row, GetCurrentVisibleLinkPath(linkType, row), storedPath, linkType.Name) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "엑셀의 원본 링크 경로와 현재 RVT의 링크 경로가 달라 건너뜁니다. 최신 상태로 다시 추출해 주세요."
                Return False
            End If

            row.CurrentLinkPath = GetCurrentVisibleLinkPath(linkType, row)
            row.StoredLinkPath = storedPath
            row.CurrentPathType = SafePathTypeText(linkType)
            Dim targetPath As String = NormalizeUserVisiblePath(row.TargetLinkPath)
            Dim targetPathType As PathType = ResolvePathType(targetPath)
            If Not IsResolvedToTargetPath(row.CurrentLinkPath, row.StoredLinkPath, targetPath, targetPathType) Then
                row.ApplyStatus = "error"
                row.ApplyMessage = "닫힌 문서 반영 후 실제 링크 경로가 대상 경로로 바뀌지 않았습니다."
                Return False
            End If
            row.LinkFileName = SafeFileName(If(Not String.IsNullOrWhiteSpace(row.CurrentLinkPath), row.CurrentLinkPath, row.TargetLinkPath))
            row.TypeWorksetNames = GetLinkTypeWorksetNamesText(doc, linkType.Id)
            row.InstanceWorksetNames = GetLinkInstanceWorksetNamesText(doc, linkType.Id)
            If String.IsNullOrWhiteSpace(SafeStr(row.ApplyTypeWorksetNames)) Then
                row.ApplyTypeWorksetNames = row.TypeWorksetNames
            End If
            If String.IsNullOrWhiteSpace(SafeStr(row.ApplyInstanceWorksetNames)) Then
                row.ApplyInstanceWorksetNames = row.InstanceWorksetNames
            End If
            row.ApplyStatus = "changed"
            row.ApplyMessage = "닫힌 문서 기준 링크 경로 반영 완료"
            Return True
        End Function

        Private Shared Function TryReloadLinkFromTarget(doc As Document,
                                                        linkType As RevitLinkType,
                                                        targetModelPath As ModelPath,
                                                        config As WorksetConfiguration,
                                                        ByRef loadResultText As String,
                                                        ByRef loadErrorText As String,
                                                        ByRef retryApplied As Boolean) As Boolean
            loadResultText = ""
            loadErrorText = ""
            retryApplied = False

            If doc Is Nothing OrElse linkType Is Nothing OrElse targetModelPath Is Nothing Then Return False

            Try
                Dim result As LinkLoadResult = linkType.LoadFrom(targetModelPath, config)
                Return EvaluateLinkLoadResult(result, loadResultText)
            Catch ex As Exception
                loadErrorText = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function PrimeLinkTypeForReload(doc As Document,
                                                       linkType As RevitLinkType,
                                                       ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If doc Is Nothing OrElse linkType Is Nothing Then Return False

            Try
                If RevitLinkType.IsLoaded(doc, linkType.Id) Then
                    Return True
                End If
            Catch
            End Try

            Try
                Dim loadResult As LinkLoadResult = linkType.Load()
                Dim loadResultText As String = ""
                If EvaluateLinkLoadResult(loadResult, loadResultText) Then
                    Return True
                End If

                Try
                    If RevitLinkType.IsLoaded(doc, linkType.Id) Then
                        Return True
                    End If
                Catch
                End Try

                errorMessage = If(String.IsNullOrWhiteSpace(loadResultText),
                                  "링크를 다시 로드하지 못했습니다.",
                                  "링크를 다시 로드하지 못했습니다. (" & loadResultText & ")")
                Return False
            Catch ex As Exception
                Try
                    If RevitLinkType.IsLoaded(doc, linkType.Id) Then
                        Return True
                    End If
                Catch
                End Try

                errorMessage = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function EvaluateLinkLoadResult(loadResult As LinkLoadResult,
                                                       ByRef loadResultText As String) As Boolean
            loadResultText = ""
            If loadResult Is Nothing Then Return True

            Try
                loadResultText = SafeStr(loadResult.LoadResult.ToString())
                Return (loadResult.LoadResult = LinkLoadResultType.LinkLoaded)
            Catch
                Return True
            End Try
        End Function

        Private Shared Function IsClosedWorksetReloadException(ex As Exception) As Boolean
            Dim message As String = SafeStr(If(ex Is Nothing, "", ex.Message)).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(message) Then Return False

            Return message.Contains("closed workset") OrElse
                   message.Contains("closed wokset") OrElse
                   (message.Contains("닫힌") AndAlso message.Contains("웍셋"))
        End Function

        Private Shared Function IsLikelyCloudPath(pathText As String) As Boolean
            Dim text As String = SafeStr(pathText).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return False

            Dim lowered As String = text.ToLowerInvariant()
            Return lowered.Contains("://") OrElse
                   lowered.Contains("autodesk docs") OrElse
                   lowered.Contains("bim 360") OrElse
                   lowered.Contains("acc://") OrElse
                   lowered.Contains("cloud") OrElse
                   lowered.Contains("forgedm") OrElse
                   lowered.Contains("dmitem")
        End Function

        Private Shared Function TryExtractBasicFileInfo(pathText As String) As BasicFileInfo
            Try
                Return BasicFileInfo.Extract(pathText)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function IsAlreadyOpen(app As RevitApp, userVisiblePath As String) As Boolean
            Return TryFindOpenDocument(app, userVisiblePath) IsNot Nothing
        End Function

        Private Shared Function TryFindOpenDocument(app As RevitApp, userVisiblePath As String) As Document
            If app Is Nothing OrElse String.IsNullOrWhiteSpace(userVisiblePath) Then Return Nothing
            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For
                    If String.Equals(NormalizeComparePath(doc.PathName), NormalizeComparePath(userVisiblePath), StringComparison.OrdinalIgnoreCase) Then
                        Return doc
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function CaptureOpenDocumentHandles(app As RevitApp) As HashSet(Of Integer)
            Dim docHandles As New HashSet(Of Integer)()
            If app Is Nothing Then Return docHandles

            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For
                    docHandles.Add(System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(doc))
                Next
            Catch
            End Try

            Return docHandles
        End Function

        Private Shared Sub CloseDocumentsOpenedDuring(app As RevitApp,
                                                      initialHandles As ISet(Of Integer))
            If app Is Nothing Then Return

            Dim initial As ISet(Of Integer) = If(initialHandles, New HashSet(Of Integer)())
            Dim docsToClose As New List(Of Document)()

            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For

                    Dim docHandle As Integer = 0
                    Try
                        docHandle = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(doc)
                    Catch
                        Continue For
                    End Try

                    If initial.Contains(docHandle) Then Continue For
                    docsToClose.Add(doc)
                Next
            Catch
                Return
            End Try

            For Each doc In docsToClose
                SafeClose(doc)
            Next
        End Sub

        Private Shared Sub SafeClose(doc As Document)
            If doc Is Nothing Then Return
            Try
                doc.Close(False)
            Catch
            End Try
        End Sub

        Private Shared Sub TryDeleteFile(pathText As String)
            Dim normalized As String = NormalizeUserVisiblePath(pathText)
            If String.IsNullOrWhiteSpace(normalized) OrElse Not File.Exists(normalized) Then Return
            Try
                File.Delete(normalized)
            Catch
            End Try
        End Sub

        Private Shared Function AppendMessage(currentMessage As String, nextMessage As String) As String
            If String.IsNullOrWhiteSpace(currentMessage) Then Return SafeStr(nextMessage)
            If String.IsNullOrWhiteSpace(nextMessage) Then Return SafeStr(currentMessage)
            Return currentMessage & " / " & nextMessage
        End Function

        Private Shared Sub MergeDocumentRowsIntoExtractedRows(app As RevitApp,
                                                              hostPath As String,
                                                              hostName As String,
                                                              rows As IList(Of RevitLinkPathRow),
                                                              progress As Action(Of Integer, String),
                                                              total As Integer,
                                                              itemIndex As Integer)
            If app Is Nothing OrElse String.IsNullOrWhiteSpace(hostPath) OrElse rows Is Nothing Then Return

            Dim doc As Document = TryFindOpenDocument(app, hostPath)
            Dim openPath As String = hostPath
            Dim createdLocal As Boolean = False
            Dim openedHere As Boolean = False

            Try
                If doc Is Nothing Then
                    Dim basicInfo As BasicFileInfo = TryExtractBasicFileInfo(hostPath)
                    If basicInfo IsNot Nothing AndAlso basicInfo.IsCentral Then
                        ReportWeightedProgress(progress, total, itemIndex, 0.60R, $"[{itemIndex + 1}/{total}] 센트럴 로컬 파일 생성 중... {hostName}")
                        openPath = CreateNewLocalPath(hostPath)
                        createdLocal = True
                    End If

                    ReportWeightedProgress(progress, total, itemIndex, 0.68R, $"[{itemIndex + 1}/{total}] RVT 열기 중 (웍셋 닫기)... {hostName}")
                    doc = OpenProjectDocument(app, openPath, closeAllWorksets:=True)
                    openedHere = (doc IsNot Nothing)
                Else
                    ReportWeightedProgress(progress, total, itemIndex, 0.68R, $"[{itemIndex + 1}/{total}] 이미 열린 문서에서 링크 확인 중... {hostName}")
                End If

                If doc Is Nothing Then Return

                Dim documentRows As List(Of RevitLinkPathRow) = ExtractRowsFromOpenDocument(doc, hostPath, hostName, progress, total, itemIndex)
                MergeExtractedRows(rows, documentRows)
            Finally
                If openedHere Then
                    SafeClose(doc)
                End If
                If createdLocal Then
                    TryDeleteFile(openPath)
                End If
            End Try
        End Sub

        Private Shared Sub MergeExtractedRows(targetRows As IList(Of RevitLinkPathRow),
                                              documentRows As IEnumerable(Of RevitLinkPathRow))
            If targetRows Is Nothing OrElse documentRows Is Nothing Then Return

            Dim existingByRefId As New Dictionary(Of String, RevitLinkPathRow)(StringComparer.OrdinalIgnoreCase)
            For Each row In targetRows
                If row Is Nothing Then Continue For
                Dim key As String = NormalizeElementIdText(row.ReferenceElementId)
                If String.IsNullOrWhiteSpace(key) Then Continue For
                If Not existingByRefId.ContainsKey(key) Then
                    existingByRefId(key) = row
                End If
            Next

            For Each docRow In documentRows
                If docRow Is Nothing Then Continue For

                Dim key As String = NormalizeElementIdText(docRow.ReferenceElementId)
                Dim existing As RevitLinkPathRow = Nothing
                If Not String.IsNullOrWhiteSpace(key) AndAlso existingByRefId.TryGetValue(key, existing) Then
                    MergeExtractedRow(existing, docRow)
                Else
                    targetRows.Add(docRow)
                    If Not String.IsNullOrWhiteSpace(key) Then
                        existingByRefId(key) = docRow
                    End If
                End If
            Next
        End Sub

        Private Shared Sub MergeExtractedRow(target As RevitLinkPathRow,
                                             source As RevitLinkPathRow)
            If target Is Nothing OrElse source Is Nothing Then Return

            target.HostFileName = ChoosePreferredValue(target.HostFileName, source.HostFileName)
            target.HostFilePath = ChoosePreferredValue(target.HostFilePath, source.HostFilePath)
            target.ReferenceElementId = ChoosePreferredValue(target.ReferenceElementId, source.ReferenceElementId)
            target.LinkName = ChoosePreferredValue(target.LinkName, source.LinkName)
            target.LinkFileName = ChoosePreferredValue(target.LinkFileName, source.LinkFileName)
            target.TypeWorksetNames = ChoosePreferredValue(target.TypeWorksetNames, source.TypeWorksetNames)
            target.InstanceWorksetNames = ChoosePreferredValue(target.InstanceWorksetNames, source.InstanceWorksetNames)
            target.ApplyTypeWorksetNames = ChoosePreferredValue(target.ApplyTypeWorksetNames, source.ApplyTypeWorksetNames)
            target.ApplyInstanceWorksetNames = ChoosePreferredValue(target.ApplyInstanceWorksetNames, source.ApplyInstanceWorksetNames)
            target.CurrentLinkPath = ChoosePreferredValue(target.CurrentLinkPath, source.CurrentLinkPath)
            target.StoredLinkPath = ChoosePreferredValue(target.StoredLinkPath, source.StoredLinkPath)
            target.CurrentPathType = ChoosePreferredValue(target.CurrentPathType, source.CurrentPathType)
        End Sub

        Private Shared Function ChoosePreferredValue(currentValue As String,
                                                     candidateValue As String) As String
            If Not String.IsNullOrWhiteSpace(SafeStr(candidateValue)) Then
                Return candidateValue
            End If
            Return currentValue
        End Function

        Private Shared Function NeedsHostWorksetRetry(row As RevitLinkPathRow) As Boolean
            If row Is Nothing Then Return False
            If String.Equals(SafeStr(row.ApplyStatus), "changed", StringComparison.OrdinalIgnoreCase) Then Return False
            Return IsClosedWorksetText(row.ApplyMessage)
        End Function

        Private Shared Function IsClosedWorksetText(message As String) As Boolean
            Dim text As String = SafeStr(message).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(text) Then Return False

            Return text.Contains("closed workset") OrElse
                   text.Contains("closed wokset") OrElse
                   (text.Contains("닫힌") AndAlso text.Contains("웍셋"))
        End Function

        Private Shared Function CollectRetryHostWorksetNames(hostDoc As Document,
                                                             rows As IEnumerable(Of RevitLinkPathRow)) As List(Of String)
            Dim names As New List(Of String)()
            If rows Is Nothing Then Return names

            For Each row In rows
                If row Is Nothing Then Continue For
                For Each name In CollectApplyWorksetNames(row, hostDoc)
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                        names.Add(name)
                    End If
                Next
            Next

            Return names
        End Function

        Private Shared Function GetLinkTypeWorksetNamesText(hostDoc As Document,
                                                            linkTypeId As ElementId) As String
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then Return ""

            Dim names As New List(Of String)()
            For Each worksetId In CollectLinkTypeWorksetIds(hostDoc, linkTypeId)
                Dim name As String = ResolveWorksetName(hostDoc, worksetId)
                If String.IsNullOrWhiteSpace(name) Then Continue For
                If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                    names.Add(name)
                End If
            Next

            Return String.Join(", ", names)
        End Function

        Private Shared Function GetLinkInstanceWorksetNamesText(hostDoc As Document,
                                                                linkTypeId As ElementId) As String
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing OrElse linkTypeId = ElementId.InvalidElementId Then Return ""

            Dim names As New List(Of String)()
            For Each worksetId In CollectLinkInstanceWorksetIds(hostDoc, linkTypeId)
                Dim name As String = ResolveWorksetName(hostDoc, worksetId)
                If String.IsNullOrWhiteSpace(name) Then Continue For
                If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                    names.Add(name)
                End If
            Next

            Return String.Join(", ", names)
        End Function

        Private Shared Function CollectApplyWorksetNames(rows As IEnumerable(Of RevitLinkPathRow)) As List(Of String)
            Dim names As New List(Of String)()
            If rows Is Nothing Then Return names

            For Each row In rows
                If row Is Nothing Then Continue For
                For Each name In ResolveApplyWorksetNames(row)
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                        names.Add(name)
                    End If
                Next
            Next

            Return names
        End Function

        Private Shared Function CollectApplyWorksetNames(row As RevitLinkPathRow,
                                                         hostDoc As Document) As IEnumerable(Of String)
            Dim preferred = ResolveApplyWorksetNames(row).ToList()
            If preferred.Count > 0 Then Return preferred
            Return CollectRowHostWorksetNames(hostDoc, row)
        End Function

        Private Shared Iterator Function ResolveApplyWorksetNames(row As RevitLinkPathRow) As IEnumerable(Of String)
            If row Is Nothing Then Return

            Dim typeText As String = If(Not String.IsNullOrWhiteSpace(SafeStr(row.ApplyTypeWorksetNames)),
                                        row.ApplyTypeWorksetNames,
                                        row.TypeWorksetNames)
            Dim instanceText As String = If(Not String.IsNullOrWhiteSpace(SafeStr(row.ApplyInstanceWorksetNames)),
                                            row.ApplyInstanceWorksetNames,
                                            row.InstanceWorksetNames)

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each name In ParseWorksetNames(typeText)
                If seen.Add(name) Then Yield name
            Next
            For Each name In ParseWorksetNames(instanceText)
                If seen.Add(name) Then Yield name
            Next
        End Function

        Private Shared Iterator Function ParseWorksetNames(text As String) As IEnumerable(Of String)
            Dim raw As String = SafeStr(text)
            If String.IsNullOrWhiteSpace(raw) Then Return

            For Each token In raw.Split(New String() {",", ";", vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                Dim name As String = token.Trim()
                If String.IsNullOrWhiteSpace(name) Then Continue For
                Yield name
            Next
        End Function

        Private Shared Function ResolveFirstWorksetName(text As String) As String
            For Each name In ParseWorksetNames(text)
                If Not String.IsNullOrWhiteSpace(name) Then Return name.Trim()
            Next
            Return ""
        End Function

        Private Shared Function CollectRowHostWorksetNames(hostDoc As Document,
                                                           row As RevitLinkPathRow) As IEnumerable(Of String)
            Dim names As New List(Of String)()
            If hostDoc Is Nothing OrElse row Is Nothing Then Return names

            Dim refId = ParseElementIdOrInvalid(row.ReferenceElementId)
            If refId Is Nothing OrElse refId = ElementId.InvalidElementId Then Return names

            For Each worksetId In CollectLinkHostWorksetIds(hostDoc, refId)
                Dim name As String = ResolveWorksetName(hostDoc, worksetId)
                If String.IsNullOrWhiteSpace(name) Then Continue For
                If Not names.Any(Function(x) String.Equals(x, name, StringComparison.OrdinalIgnoreCase)) Then
                    names.Add(name)
                End If
            Next

            Return names
        End Function

        Private Shared Function EnsureEditableHostWorksets(hostDoc As Document,
                                                           linkTypeId As ElementId,
                                                           ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If hostDoc Is Nothing Then Return True

            Try
                If Not hostDoc.IsWorkshared Then Return True
            Catch
                Return True
            End Try

            Try
                Dim requestedIds = CollectLinkHostWorksetIds(hostDoc, linkTypeId)
                If requestedIds.Count = 0 Then Return True

                Dim checkedOut = WorksharingUtils.CheckoutWorksets(hostDoc, requestedIds)
                If checkedOut Is Nothing Then
                    errorMessage = "링크가 배치된 호스트 웍셋 checkout 결과를 확인하지 못했습니다."
                    Return False
                End If

                Dim missingNames As New List(Of String)()
                For Each worksetId In requestedIds
                    If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Continue For
                    If Not checkedOut.Contains(worksetId) Then
                        missingNames.Add(ResolveWorksetName(hostDoc, worksetId))
                    End If
                Next

                If missingNames.Count > 0 Then
                    errorMessage = "링크가 들어있는 호스트 웍셋을 checkout하지 못했습니다: " & String.Join(", ", missingNames.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
                    Return False
                End If

                Return True
            Catch ex As Exception
                errorMessage = "호스트 웍셋 checkout 실패: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function EnsureEditableNamedWorksets(hostDoc As Document,
                                                            worksetNames As IEnumerable(Of String),
                                                            ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If hostDoc Is Nothing OrElse worksetNames Is Nothing Then Return True

            Try
                If Not hostDoc.IsWorkshared Then Return True
            Catch
                Return True
            End Try

            Dim requestedIds As New HashSet(Of WorksetId)()
            Dim missingNames As New List(Of String)()
            For Each rawName In worksetNames
                Dim name As String = SafeStr(rawName).Trim()
                If String.IsNullOrWhiteSpace(name) Then Continue For

                Dim worksetId As WorksetId = FindUserWorksetIdByName(hostDoc, name)
                If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then
                    missingNames.Add(name)
                Else
                    requestedIds.Add(worksetId)
                End If
            Next

            If missingNames.Count > 0 Then
                errorMessage = "지정한 호스트 웍셋을 찾지 못했습니다: " & String.Join(", ", missingNames)
                Return False
            End If

            If requestedIds.Count = 0 Then Return True

            Try
                Dim checkedOut = WorksharingUtils.CheckoutWorksets(hostDoc, requestedIds)
                If checkedOut Is Nothing Then
                    errorMessage = "지정한 호스트 웍셋 checkout 결과를 확인하지 못했습니다."
                    Return False
                End If

                Dim notCheckedOut As New List(Of String)()
                For Each worksetId In requestedIds
                    If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Continue For
                    If Not checkedOut.Contains(worksetId) Then
                        notCheckedOut.Add(ResolveWorksetName(hostDoc, worksetId))
                    End If
                Next

                If notCheckedOut.Count > 0 Then
                    errorMessage = "지정한 호스트 웍셋을 checkout하지 못했습니다: " & String.Join(", ", notCheckedOut.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
                    Return False
                End If

                Return True
            Catch ex As Exception
                errorMessage = "지정한 호스트 웍셋 checkout 실패: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function CollectLinkHostWorksetIds(hostDoc As Document,
                                                          linkTypeId As ElementId) As ICollection(Of WorksetId)
            Dim ids As New HashSet(Of WorksetId)()
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing Then Return ids

            For Each worksetId In CollectLinkTypeWorksetIds(hostDoc, linkTypeId)
                ids.Add(worksetId)
            Next

            For Each worksetId In CollectLinkInstanceWorksetIds(hostDoc, linkTypeId)
                ids.Add(worksetId)
            Next

            Return ids
        End Function

        Private Shared Function CollectLinkTypeWorksetIds(hostDoc As Document,
                                                          linkTypeId As ElementId) As ICollection(Of WorksetId)
            Dim ids As New HashSet(Of WorksetId)()
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing Then Return ids

            Try
                Dim linkType As RevitLinkType = TryCast(hostDoc.GetElement(linkTypeId), RevitLinkType)
                If linkType IsNot Nothing AndAlso linkType.WorksetId IsNot Nothing AndAlso linkType.WorksetId <> WorksetId.InvalidWorksetId Then
                    ids.Add(linkType.WorksetId)
                End If
            Catch
            End Try

            Return ids
        End Function

        Private Shared Function CollectLinkInstanceWorksetIds(hostDoc As Document,
                                                              linkTypeId As ElementId) As ICollection(Of WorksetId)
            Dim ids As New HashSet(Of WorksetId)()
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing Then Return ids

            Try
                For Each inst In CollectLinkInstances(hostDoc, linkTypeId)
                    If inst Is Nothing Then Continue For
                    Try
                        Dim wsId As WorksetId = inst.WorksetId
                        If wsId IsNot Nothing AndAlso wsId <> WorksetId.InvalidWorksetId Then
                            ids.Add(wsId)
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            Return ids
        End Function

        Private Shared Function CollectLinkInstances(hostDoc As Document,
                                                     linkTypeId As ElementId) As List(Of RevitLinkInstance)
            Dim instances As New List(Of RevitLinkInstance)()
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing Then Return instances

            Try
                instances =
                    New FilteredElementCollector(hostDoc).
                        OfClass(GetType(RevitLinkInstance)).
                        WhereElementIsNotElementType().
                        Cast(Of RevitLinkInstance)().
                        Where(Function(inst) inst IsNot Nothing AndAlso inst.GetTypeId() = linkTypeId).
                        ToList()
            Catch
            End Try

            Return instances
        End Function

        Private Shared Function ResolveWorksetName(hostDoc As Document,
                                                   worksetId As WorksetId) As String
            If hostDoc Is Nothing OrElse worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Return ""

            Try
                Dim table As WorksetTable = hostDoc.GetWorksetTable()
                If table IsNot Nothing Then
                    Dim ws As Workset = table.GetWorkset(worksetId)
                    If ws IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ws.Name) Then
                        Return ws.Name
                    End If
                End If
            Catch
            End Try

            Try
                Return worksetId.IntegerValue.ToString(CultureInfo.InvariantCulture)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function FindUserWorksetIdByName(hostDoc As Document,
                                                        worksetName As String) As WorksetId
            If hostDoc Is Nothing OrElse String.IsNullOrWhiteSpace(worksetName) Then Return WorksetId.InvalidWorksetId

            Try
                Dim collector As New FilteredWorksetCollector(hostDoc)
                For Each ws As Workset In collector.OfKind(WorksetKind.UserWorkset)
                    If ws Is Nothing Then Continue For
                    If String.Equals(SafeStr(ws.Name).Trim(), worksetName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        Return ws.Id
                    End If
                Next
            Catch
            End Try

            Return WorksetId.InvalidWorksetId
        End Function

        Private Shared Function SetElementWorksetByName(element As Element,
                                                        worksetName As String) As Boolean
            If element Is Nothing OrElse String.IsNullOrWhiteSpace(worksetName) Then Return False

            Dim doc As Document = Nothing
            Try
                doc = element.Document
            Catch
                doc = Nothing
            End Try
            If doc Is Nothing Then Return False

            Dim worksetId As WorksetId = FindUserWorksetIdByName(doc, worksetName)
            If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Return False

            Try
                Dim p As Parameter = element.Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                If p Is Nothing OrElse p.IsReadOnly Then Return False
                Return p.Set(worksetId.IntegerValue)
            Catch
                Return False
            End Try
        End Function

        Private Shared Sub MarkRows(rows As IEnumerable(Of RevitLinkPathRow),
                                    status As String,
                                    message As String,
                                    Optional overwriteChanged As Boolean = False)
            If rows Is Nothing Then Return
            For Each row In rows
                If row Is Nothing Then Continue For
                If Not overwriteChanged AndAlso String.Equals(row.ApplyStatus, "changed", StringComparison.OrdinalIgnoreCase) Then Continue For
                row.ApplyStatus = status
                row.ApplyMessage = message
            Next
        End Sub

        Private Shared Function IsNewLinkRow(row As RevitLinkPathRow) As Boolean
            If row Is Nothing Then Return False
            Return String.IsNullOrWhiteSpace(SafeStr(row.ReferenceElementId)) AndAlso
                   String.IsNullOrWhiteSpace(SafeStr(row.LinkName)) AndAlso
                   Not String.IsNullOrWhiteSpace(SafeStr(row.TargetLinkPath))
        End Function

        Private Shared Function IsDeleteLinkRow(row As RevitLinkPathRow) As Boolean
            If row Is Nothing Then Return False
            Return Not String.IsNullOrWhiteSpace(SafeStr(row.ReferenceElementId)) AndAlso
                   String.IsNullOrWhiteSpace(NormalizeUserVisiblePath(row.TargetLinkPath))
        End Function

        Private Shared Function IsLikelyCloudRow(row As RevitLinkPathRow) As Boolean
            If row Is Nothing Then Return False

            Dim currentPathType As String = SafeStr(row.CurrentPathType).Trim()
            If String.Equals(currentPathType, "Cloud", StringComparison.OrdinalIgnoreCase) Then Return True

            If IsLikelyCloudPath(row.CurrentLinkPath) Then Return True
            If IsLikelyCloudPath(row.StoredLinkPath) Then Return True

            Return False
        End Function

        Private Shared Function ShouldBypassOriginalLinkReferenceCheck(row As RevitLinkPathRow,
                                                                       linkType As RevitLinkType,
                                                                       actualVisiblePath As String,
                                                                       actualStoredPath As String) As Boolean
            If IsLikelyCloudRow(row) Then Return True
            If IsLikelyCloudPath(actualVisiblePath) OrElse IsLikelyCloudPath(actualStoredPath) Then Return True
            If linkType IsNot Nothing AndAlso IsCloudLinkType(linkType, actualVisiblePath, actualStoredPath) Then Return True
            Return False
        End Function

        Private Shared Function IsOriginalLinkReferenceStillCurrent(row As RevitLinkPathRow,
                                                                    actualVisiblePath As String,
                                                                    actualStoredPath As String,
                                                                    Optional actualLinkName As String = "") As Boolean
            If row Is Nothing OrElse IsNewLinkRow(row) Then Return True

            Dim expectedCurrent As String = NormalizeUserVisiblePath(row.CurrentLinkPath)
            Dim expectedStored As String = NormalizeUserVisiblePath(row.StoredLinkPath)
            If String.IsNullOrWhiteSpace(expectedCurrent) AndAlso String.IsNullOrWhiteSpace(expectedStored) Then
                Return IsOriginalLinkNameStillCurrent(row, actualLinkName)
            End If

            Dim currentPath As String = NormalizeUserVisiblePath(actualVisiblePath)
            Dim storedPath As String = NormalizeUserVisiblePath(actualStoredPath)

            If PathsMatch(expectedCurrent, currentPath) OrElse PathsMatch(expectedCurrent, storedPath) Then Return True
            If PathsMatch(expectedStored, currentPath) OrElse PathsMatch(expectedStored, storedPath) Then Return True
            Return False
        End Function

        Private Shared Function IsOriginalLinkNameStillCurrent(row As RevitLinkPathRow,
                                                               actualLinkName As String) As Boolean
            If row Is Nothing Then Return True

            Dim expectedName As String = SafeStr(row.LinkName).Trim()
            Dim currentName As String = SafeStr(actualLinkName).Trim()
            If String.IsNullOrWhiteSpace(expectedName) OrElse String.IsNullOrWhiteSpace(currentName) Then Return True

            Return String.Equals(expectedName, currentName, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function PathsMatch(left As String, right As String) As Boolean
            If String.IsNullOrWhiteSpace(left) OrElse String.IsNullOrWhiteSpace(right) Then Return False
            Return String.Equals(NormalizeComparePath(left), NormalizeComparePath(right), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsResolvedToTargetPath(currentPath As String,
                                                       storedPath As String,
                                                       targetPath As String,
                                                       targetPathType As PathType) As Boolean
            Dim normalizedTarget As String = NormalizeUserVisiblePath(targetPath)
            If String.IsNullOrWhiteSpace(normalizedTarget) Then Return False

            If targetPathType = PathType.Relative Then
                Return True
            End If

            Return PathsMatch(normalizedTarget, currentPath) OrElse
                   PathsMatch(normalizedTarget, storedPath)
        End Function

        Private Shared Function ResolveNewLinkPlacement(value As String) As ImportPlacement
            Dim text As String = SafeStr(value).Trim().Replace("-"c, "_"c).Replace(" "c, "_"c).ToLowerInvariant()
            Select Case text
                Case "center", "centered", "center_to_center", "centertocenter"
                    Return ImportPlacement.Centered
                Case "shared", "shared_coordinates", "sharedcoordinates"
                    Return ImportPlacement.Shared
                Case "site", "project_base_point", "projectbasepoint"
                    Return ImportPlacement.Site
                Case Else
                    Return ImportPlacement.Origin
            End Select
        End Function

        Private Shared Sub ResolveMissingHostFilePaths(rows As IList(Of RevitLinkPathRow),
                                                       Optional hostPathHints As IEnumerable(Of String) = Nothing)
            If rows Is Nothing Then Return

            Dim pathsByName As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim duplicateNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If hostPathHints IsNot Nothing Then
                For Each hintPath In hostPathHints
                    RegisterHostPathHint(pathsByName, duplicateNames, hintPath, "")
                Next
            End If

            For Each row In rows
                If row Is Nothing Then Continue For

                row.HostFilePath = NormalizeUserVisiblePath(row.HostFilePath)
                RegisterHostPathHint(pathsByName, duplicateNames, row.HostFilePath, row.HostFileName)
            Next

            For Each row In rows
                If row Is Nothing OrElse Not String.IsNullOrWhiteSpace(row.HostFilePath) Then Continue For

                Dim hostNameText As String = NormalizeUserVisiblePath(row.HostFileName)
                If String.IsNullOrWhiteSpace(hostNameText) Then Continue For

                If IsServerPath(hostNameText) OrElse IsRootedPath(hostNameText) Then
                    row.HostFilePath = hostNameText
                    row.HostFileName = SafeFileName(hostNameText)
                    Continue For
                End If

                Dim hostKey As String = SafeFileName(hostNameText)
                If String.IsNullOrWhiteSpace(hostKey) Then Continue For
                If duplicateNames.Contains(hostKey) Then
                    row.ApplyStatus = "error"
                    row.ApplyMessage = "동일한 HostFileName을 가진 RVT가 여러 개라 호스트를 특정할 수 없습니다. HostFilePath를 입력해 주세요."
                    Continue For
                End If

                Dim resolvedPath As String = Nothing
                If pathsByName.TryGetValue(hostKey, resolvedPath) Then
                    row.HostFilePath = resolvedPath
                End If
            Next
        End Sub

        Private Shared Sub RegisterHostPathHint(pathsByName As IDictionary(Of String, String),
                                                duplicateNames As ISet(Of String),
                                                rawHostPath As String,
                                                rawHostName As String)
            If pathsByName Is Nothing OrElse duplicateNames Is Nothing Then Return

            Dim hostPath As String = NormalizeUserVisiblePath(rawHostPath)
            If String.IsNullOrWhiteSpace(hostPath) Then Return

            Dim hostKey As String = SafeFileName(If(Not String.IsNullOrWhiteSpace(rawHostName), rawHostName, hostPath))
            If String.IsNullOrWhiteSpace(hostKey) Then Return

            Dim existingPath As String = Nothing
            If pathsByName.TryGetValue(hostKey, existingPath) Then
                If Not PathsMatch(existingPath, hostPath) Then duplicateNames.Add(hostKey)
            Else
                pathsByName(hostKey) = hostPath
            End If
        End Sub

        Private Shared Function IsRootedPath(pathText As String) As Boolean
            Try
                Return Path.IsPathRooted(NormalizeUserVisiblePath(pathText))
            Catch
                Return False
            End Try
        End Function

        Private Shared Function CloneRows(rows As IEnumerable(Of RevitLinkPathRow)) As List(Of RevitLinkPathRow)
            Dim cloned As New List(Of RevitLinkPathRow)()
            If rows Is Nothing Then Return cloned

            For Each row In rows
                If row Is Nothing Then Continue For
                cloned.Add(New RevitLinkPathRow With {
                    .HostFileName = SafeStr(row.HostFileName),
                    .HostFilePath = SafeStr(row.HostFilePath),
                    .ReferenceElementId = SafeStr(row.ReferenceElementId),
                    .LinkName = SafeStr(row.LinkName),
                    .LinkFileName = SafeStr(row.LinkFileName),
                    .TypeWorksetNames = SafeStr(row.TypeWorksetNames),
                    .InstanceWorksetNames = SafeStr(row.InstanceWorksetNames),
                    .ApplyTypeWorksetNames = SafeStr(row.ApplyTypeWorksetNames),
                    .ApplyInstanceWorksetNames = SafeStr(row.ApplyInstanceWorksetNames),
                    .CurrentLinkPath = SafeStr(row.CurrentLinkPath),
                    .StoredLinkPath = SafeStr(row.StoredLinkPath),
                    .CurrentPathType = SafeStr(row.CurrentPathType),
                    .TargetLinkPath = SafeStr(row.TargetLinkPath),
                    .TargetPathType = SafeStr(row.TargetPathType),
                    .ApplyStatus = SafeStr(row.ApplyStatus),
                    .ApplyMessage = SafeStr(row.ApplyMessage)
                })
            Next

            Return cloned
        End Function

        Private Shared Function NormalizePaths(paths As IList(Of String)) As List(Of String)
            Dim results As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If paths Is Nothing Then Return results

            For Each raw In paths
                Dim path As String = NormalizeUserVisiblePath(raw)
                If String.IsNullOrWhiteSpace(path) Then Continue For
                If seen.Add(path) Then results.Add(path)
            Next

            Return results
        End Function

        Private Shared Function NormalizeUserVisiblePath(value As String) As String
            Dim path As String = SafeStr(value).Trim()
            If String.IsNullOrWhiteSpace(path) Then Return ""

            path = path.Trim(""""c)
            If path.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim u As New Uri(path)
                    If u IsNot Nothing AndAlso u.IsFile Then path = u.LocalPath
                Catch
                End Try
            End If

            If path.Contains("://") Then
                Return path.Trim()
            End If

            path = path.Replace("/"c, "\"c)
            Return path.Trim()
        End Function

        Private Shared Function NormalizeComparePath(path As String) As String
            Return NormalizeUserVisiblePath(path).TrimEnd("\"c).ToLowerInvariant()
        End Function

        Private Shared Function ResolvePathType(targetPath As String) As PathType
            If IsServerPath(targetPath) Then Return PathType.Server
            If Path.IsPathRooted(targetPath) Then Return PathType.Absolute
            Return PathType.Relative
        End Function

        Private Shared Function IsServerPath(path As String) As Boolean
            Dim s As String = NormalizeUserVisiblePath(path)
            If String.IsNullOrWhiteSpace(s) Then Return False
            Return s.StartsWith("RSN:\", StringComparison.OrdinalIgnoreCase) OrElse
                   s.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function TryGetAbsoluteModelPath(extRef As ExternalFileReference) As ModelPath
            If extRef Is Nothing Then Return Nothing
            Try
                Return extRef.GetAbsolutePath()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryGetAbsoluteModelPath(linkType As RevitLinkType) As ModelPath
            Return TryGetAbsoluteModelPath(TryGetExternalFileReference(linkType))
        End Function

        Private Shared Function TryGetStoredModelPath(extRef As ExternalFileReference) As ModelPath
            If extRef Is Nothing Then Return Nothing
            Try
                Return extRef.GetPath()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryGetStoredModelPath(linkType As RevitLinkType) As ModelPath
            Return TryGetStoredModelPath(TryGetExternalFileReference(linkType))
        End Function

        Private Shared Function TryGetExternalFileReference(linkType As RevitLinkType) As ExternalFileReference
            If linkType Is Nothing Then Return Nothing
            Try
                Return linkType.GetExternalFileReference()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryGetExternalResourceReference(linkType As RevitLinkType) As ExternalResourceReference
            If linkType Is Nothing Then Return Nothing
            Try
                Dim resourceRefs = linkType.GetExternalResourceReferences()
                If resourceRefs Is Nothing Then Return Nothing
                For Each pair In resourceRefs
                    If pair.Value IsNot Nothing Then Return pair.Value
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function GetCurrentVisibleLinkPath(linkType As RevitLinkType,
                                                          row As RevitLinkPathRow) As String
            Dim currentVisiblePath As String = SafeModelPathToUserVisiblePath(TryGetAbsoluteModelPath(linkType))
            If String.IsNullOrWhiteSpace(currentVisiblePath) Then
                currentVisiblePath = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            End If
            If String.IsNullOrWhiteSpace(currentVisiblePath) Then
                currentVisiblePath = GetExternalResourceDisplayPath(TryGetExternalResourceReference(linkType))
            End If
            If String.IsNullOrWhiteSpace(currentVisiblePath) AndAlso row IsNot Nothing Then
                currentVisiblePath = NormalizeUserVisiblePath(row.CurrentLinkPath)
            End If
            Return SafeStr(currentVisiblePath).Trim()
        End Function

        Private Shared Function SafeModelPathToUserVisiblePath(modelPath As ModelPath) As String
            If modelPath Is Nothing Then Return ""
            Try
                Return SafeStr(ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath))
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafePathTypeText(extRef As ExternalFileReference) As String
            If extRef Is Nothing Then Return ""
            Try
                Return SafeStr(extRef.PathType.ToString())
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafePathTypeText(linkType As RevitLinkType) As String
            Dim extRef As ExternalFileReference = TryGetExternalFileReference(linkType)
            Dim pathTypeText As String = SafePathTypeText(extRef)
            If Not String.IsNullOrWhiteSpace(pathTypeText) Then
                Return pathTypeText
            End If

            If IsCloudLinkType(linkType) Then
                Return "Cloud"
            End If

            Return ""
        End Function

        Private Shared Function IsCloudLinkType(linkType As RevitLinkType,
                                                Optional currentVisiblePath As String = "",
                                                Optional storedPath As String = "") As Boolean
            If linkType Is Nothing Then Return False

            Dim extRef As ExternalFileReference = TryGetExternalFileReference(linkType)
            If extRef IsNot Nothing Then
                Dim pathTypeText As String = SafePathTypeText(extRef)
                If Not String.IsNullOrWhiteSpace(pathTypeText) Then
                    Return False
                End If
            End If

            Dim visiblePath As String = NormalizeUserVisiblePath(currentVisiblePath)
            If String.IsNullOrWhiteSpace(visiblePath) Then
                visiblePath = SafeModelPathToUserVisiblePath(TryGetAbsoluteModelPath(linkType))
            End If
            If String.IsNullOrWhiteSpace(visiblePath) Then
                visiblePath = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            End If
            If IsLikelyCloudPath(visiblePath) Then Return True

            Dim storedVisiblePath As String = NormalizeUserVisiblePath(storedPath)
            If String.IsNullOrWhiteSpace(storedVisiblePath) Then
                storedVisiblePath = SafeModelPathToUserVisiblePath(TryGetStoredModelPath(linkType))
            End If
            If IsLikelyCloudPath(storedVisiblePath) Then Return True

            Dim resourceRef As ExternalResourceReference = TryGetExternalResourceReference(linkType)
            If resourceRef Is Nothing Then Return False

            Dim resourceDisplayPath As String = GetExternalResourceDisplayPath(resourceRef)
            If IsLikelyCloudPath(resourceDisplayPath) Then Return True

            Return IsLikelyCloudPath(SerializeExternalResourceReference(resourceRef))
        End Function

        Private Shared Function GetExternalResourceDisplayPath(resourceRef As ExternalResourceReference) As String
            If resourceRef Is Nothing Then Return ""

            Try
                Dim inSessionPath As String = SafeStr(resourceRef.InSessionPath).Trim()
                If Not String.IsNullOrWhiteSpace(inSessionPath) Then
                    Return inSessionPath
                End If
            Catch
            End Try

            Dim shortName As String = GetExternalResourceShortName(resourceRef)
            If Not String.IsNullOrWhiteSpace(shortName) Then
                Return shortName
            End If

            Return SerializeExternalResourceReference(resourceRef)
        End Function

        Private Shared Function GetExternalResourceShortName(resourceRef As ExternalResourceReference) As String
            If resourceRef Is Nothing Then Return ""
            Try
                Return SafeStr(resourceRef.GetResourceShortDisplayName()).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SerializeExternalResourceReference(resourceRef As ExternalResourceReference) As String
            If resourceRef Is Nothing Then Return ""

            Dim parts As New List(Of String)()
            Try
                Dim info = resourceRef.GetReferenceInformation()
                If info IsNot Nothing Then
                    For Each kv In info.OrderBy(Function(x) SafeStr(x.Key), StringComparer.OrdinalIgnoreCase)
                        If String.IsNullOrWhiteSpace(kv.Key) OrElse String.IsNullOrWhiteSpace(kv.Value) Then Continue For
                        parts.Add($"{kv.Key}={kv.Value}")
                    Next
                End If
            Catch
            End Try

            If parts.Count = 0 Then
                Try
                    Dim versionText As String = SafeStr(resourceRef.Version)
                    If Not String.IsNullOrWhiteSpace(versionText) Then
                        parts.Add("Version=" & versionText)
                    End If
                Catch
                End Try
            End If

            Return String.Join("; ", parts)
        End Function

        Private Shared Function SafeElementIdText(refId As ElementId) As String
            If refId Is Nothing Then Return ""
            Try
                Return refId.IntegerValue.ToString(CultureInfo.InvariantCulture)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeFileName(path As String) As String
            Dim s As String = NormalizeUserVisiblePath(path)
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Try
                Return System.IO.Path.GetFileName(s)
            Catch
                Return s
            End Try
        End Function

        Private Shared Sub ReportProgress(progress As Action(Of Integer, String),
                                          total As Integer,
                                          current As Integer,
                                          message As String)
            If progress Is Nothing Then Return
            Dim safeTotal As Integer = Math.Max(1, total)
            Dim safeCurrent As Integer = Math.Max(0, Math.Min(current, safeTotal))
            Dim percent As Integer = CInt(Math.Round((CDbl(safeCurrent) / CDbl(safeTotal)) * 100.0R))
            progress(percent, SafeStr(message))
        End Sub

        Private Shared Sub ReportWeightedProgress(progress As Action(Of Integer, String),
                                                  total As Integer,
                                                  itemIndex As Integer,
                                                  itemFraction As Double,
                                                  message As String)
            If progress Is Nothing Then Return

            Dim safeTotal As Integer = Math.Max(1, total)
            Dim safeIndex As Integer = Math.Max(0, Math.Min(itemIndex, safeTotal - 1))
            Dim safeFraction As Double = Math.Max(0.0R, Math.Min(1.0R, itemFraction))
            Dim percent As Integer = CInt(Math.Round(((CDbl(safeIndex) + safeFraction) / CDbl(safeTotal)) * 100.0R))
            progress(Math.Max(0, Math.Min(100, percent)), SafeStr(message))
        End Sub

        Private Shared Function ParseElementIdOrInvalid(value As String) As ElementId
            Dim idValue As Integer
            If Integer.TryParse(NormalizeElementIdText(value), idValue) Then
                Return New ElementId(idValue)
            End If
            Return ElementId.InvalidElementId
        End Function

        Private Shared Function NormalizeElementIdText(value As String) As String
            Dim text As String = SafeStr(value).Trim()
            If String.IsNullOrWhiteSpace(text) Then Return ""

            If text.StartsWith(",", StringComparison.Ordinal) Then
                text = text.Substring(1).Trim()
            End If
            If text.EndsWith(",", StringComparison.Ordinal) Then
                text = text.Substring(0, text.Length - 1).Trim()
            End If

            If text.Contains(","c) Then
                Dim compact As String = text.Replace(",", String.Empty).Trim()
                Dim parsed As Long
                If Long.TryParse(compact, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    Return compact
                End If

                text = text.Split(","c).
                            Select(Function(part) SafeStr(part).Trim()).
                            FirstOrDefault(Function(part) Not String.IsNullOrWhiteSpace(part))
                If text Is Nothing Then Return ""
            End If

            Return text.Trim()
        End Function

        Private Shared Function SafeStr(value As String) As String
            Return If(value, "")
        End Function

        Private Shared Function FindHeaderIndex(headers As IDictionary(Of Integer, String), headerName As String) As Integer
            If headers Is Nothing OrElse String.IsNullOrWhiteSpace(headerName) Then Return -1
            For Each kv In headers
                If String.Equals(SafeStr(kv.Value).Trim(), headerName, StringComparison.OrdinalIgnoreCase) Then
                    Return kv.Key
                End If
            Next
            Return -1
        End Function

        Private Shared Function GetCellText(cell As ICell, formatter As DataFormatter) As String
            If cell Is Nothing Then Return ""
            Try
                Return If(formatter Is Nothing, SafeStr(cell.ToString()), formatter.FormatCellValue(cell)).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetCellTextByHeader(headers As IDictionary(Of Integer, String),
                                                    row As IRow,
                                                    formatter As DataFormatter,
                                                    headerName As String) As String
            If row Is Nothing Then Return ""
            Dim colIndex As Integer = FindHeaderIndex(headers, headerName)
            If colIndex < 0 Then Return ""
            Return GetCellText(row.GetCell(colIndex), formatter)
        End Function

    End Class

End Namespace

