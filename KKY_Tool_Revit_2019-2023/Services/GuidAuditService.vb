Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application
Imports RvtDB = Autodesk.Revit.DB

Namespace Services

    ''' <summary>
    ''' GUID Audit 기능 포팅(Service 계층)
    '''  - 모드 1: 프로젝트 파라미터 vs 공유 파라미터 파일 GUID 비교
    '''  - 모드 2: 로드 패밀리 공유 파라미터 vs 공유 파라미터 파일 GUID 비교
    ''' </summary>
    Public NotInheritable Class GuidAuditService

        Private Sub New()
        End Sub

        Public Class RunResult
            Public Property Mode As Integer
            Public Property Project As DataTable
            Public Property FamilyDetail As DataTable
            Public Property FamilyIndex As DataTable
            Public Property RunId As String
            Public Property IncludeFamily As Boolean
        End Class

        Public Class CleanupSettings
            Public Property CloseAllWorksetsOnOpen As Boolean = True
            Public Property UseSyncComment As Boolean = False
            Public Property SyncComment As String = String.Empty
        End Class

        Public Class CleanupFileResult
            Public Property FilePath As String = String.Empty
            Public Property FileName As String = String.Empty
            Public Property Status As String = String.Empty
            Public Property RequestedDeleteCount As Integer
            Public Property DeletedCount As Integer
            Public Property ProjectDeletedCount As Integer
            Public Property FamilyDeletedCount As Integer
            Public Property WasCentralFile As Boolean
            Public Property UsedLocalFile As Boolean
            Public Property SynchronizePerformed As Boolean
            Public Property Message As String = String.Empty
        End Class

        Public Class CleanupResult
            Public Property Ok As Boolean
            Public Property Message As String = String.Empty
            Public Property SourceExcelPath As String = String.Empty
            Public Property InstructionCount As Integer
            Public Property DeletedCount As Integer
            Public Property SuccessCount As Integer
            Public Property FailCount As Integer
            Public Property NoChangeCount As Integer
            Public Property Files As List(Of CleanupFileResult) = New List(Of CleanupFileResult)()
        End Class

        Private Class TargetFile
            Public Property Path As String = String.Empty
            Public Property Name As String = String.Empty
        End Class

        Private Class CleanupInstruction
            Public Property DeleteScope As String = String.Empty
            Public Property DeleteKey As String = String.Empty
            Public Property RvtPath As String = String.Empty
            Public Property ParamName As String = String.Empty
            Public Property ParamKind As String = String.Empty
            Public Property RvtGuid As String = String.Empty
            Public Property FileGuid As String = String.Empty
            Public Property BoundCategories As String = String.Empty
            Public Property FamilyName As String = String.Empty
            Public Property FamilyCategory As String = String.Empty
            Public Property FamilyGuid As String = String.Empty
            Public Property SheetName As String = String.Empty
            Public Property RowNumber As Integer
        End Class

        Private Class ProjectBindingCandidate
            Public Property Definition As Definition
            Public Property ParamName As String = String.Empty
            Public Property ParamKind As String = String.Empty
            Public Property ParamGuid As String = String.Empty
        End Class

        ''' <summary>
        ''' GUID Audit 실행
        ''' </summary>
        Public Shared Function Run(app As UIApplication,
                                   mode As Integer,
                                   rvtPaths As IEnumerable(Of String),
                                   progress As Action(Of Double, String),
                                   Optional warn As Action(Of String) = Nothing,
                                   Optional includeFamily As Boolean = False,
                                   Optional includeAnnotation As Boolean = False) As RunResult

            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim defMap = SharedParamReader.ReadSharedParamNameGuidMap(app.Application)
            If defMap Is Nothing OrElse defMap.Count = 0 Then
                Throw New InvalidOperationException("공유 파라미터 파일이 설정되어 있지 않거나 읽을 수 없습니다. (Revit 옵션에서 Shared Parameter 파일 경로 확인)")
            End If

            Dim targets = BuildTargets(app, rvtPaths)
            If targets.Count = 0 Then
                Throw New InvalidOperationException("검토할 RVT 파일이 없습니다.")
            End If

            Dim runId As String = Guid.NewGuid().ToString("N")
            Dim total As Integer = targets.Count
            Dim projectTable As DataTable = Nothing
            Dim familyDetail As DataTable = Nothing
            Dim famIndex As DataTable = Nothing

            For i As Integer = 0 To total - 1
                Dim target = targets(i)
                Dim openedByMe As Boolean = False
                Dim doc As Document = Nothing
                Dim openError As String = ""

                Try
                    ReportProgress(progress, total, i + 1, 0.02R, $"문서 여는 중... {i + 1}/{total} {target.Name}")
                    doc = ResolveOrOpenDocument(app, app.ActiveUIDocument?.Document, target.Path, openedByMe, openError)

                    If doc Is Nothing Then
                        Dim failProj = Auditors.MakeFailureSummaryTable(1)
                        Dim note = BuildOpenFailNotes(openError, target.Path)
                        Dim shortReason = ShortenReason(note)
                        If warn IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(note) Then
                            warn(note)
                        End If
                        ReportProgress(progress, total, i + 1, 0.08R, $"문서 열기 실패: {target.Name} - {shortReason}")
                        Auditors.AddOpenFailRow(failProj, target.Name, target.Path, "Project", "OPEN_FAIL", note)
                        projectTable = MergeTable(projectTable, failProj)
                        If includeFamily Then
                            Dim failFam = Auditors.MakeFailureSummaryTable(2)
                            Auditors.AddOpenFailRow(failFam, target.Name, target.Path, "Family", "OPEN_FAIL", note)
                            familyDetail = MergeTable(familyDetail, failFam)
                        End If
                        Continue For
                    End If

                    Dim rvtName As String = GetRvtName(doc, target.Path)
                    Dim captureIndex As Integer = i
                    Dim captureName As String = rvtName

                    Dim proj = Auditors.RunProjectParameterAudit(doc, defMap, rvtName, target.Path,
                                                                 Function(cur, tot) As Object
                                                                     Dim frac As Double = 0.1R + 0.8R * SafeRatio(cur, tot)
                                                                     ReportProgress(progress, total, captureIndex + 1, frac, $"[{captureName}] 프로젝트 파라미터 ({cur}/{tot})")
                                                                     Return Nothing
                                                                 End Function)
                    projectTable = MergeTable(projectTable, proj)

                    If includeFamily Then
                        Dim famPack = Auditors.RunFamilyAudit(doc, defMap, rvtName, target.Path,
                                                              Function(cur, tot, famName) As Object
                                                                  Dim frac As Double = 0.1R + 0.8R * SafeRatio(cur, tot)
                                                                  ReportProgress(progress, total, captureIndex + 1, frac, $"[{captureName}] 패밀리 처리 중 ({cur}/{tot}) {famName}")
                                                                  Return Nothing
                                                              End Function,
                                                              includeAnnotation)
                        familyDetail = MergeTable(familyDetail, famPack.Detail)
                        famIndex = MergeTable(famIndex, famPack.Index)
                    End If

                    ReportProgress(progress, total, captureIndex + 1, 1.0R, $"완료: {captureIndex + 1}/{total} {captureName}")

                Catch ex As Exception
                    Dim fail = Auditors.MakeFailureSummaryTable(1)
                    Dim note = BuildExceptionNotes(ex, target.Path)
                    ReportProgress(progress, total, i + 1, 0.08R, $"문서 처리 실패: {target.Name} - {ShortenReason(note)}")
                    Auditors.AddOpenFailRow(fail, target.Name, target.Path, "Project", "ERROR", note)
                    projectTable = MergeTable(projectTable, fail)
                    If includeFamily Then
                        Dim failFam = Auditors.MakeFailureSummaryTable(2)
                        Auditors.AddOpenFailRow(failFam, target.Name, target.Path, "Family", "ERROR", note)
                        familyDetail = MergeTable(familyDetail, failFam)
                    End If

                Finally
                    If openedByMe AndAlso doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            ' ResultTableFilter.KeepOnlyIssues("guid", projectTable)
            If includeFamily Then
                ' ResultTableFilter.KeepOnlyIssues("guid", familyDetail)

                Dim famSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                If familyDetail IsNot Nothing AndAlso familyDetail.Columns.Contains("FamilyName") Then
                    For Each r As DataRow In familyDetail.Rows
                        Dim fn As String = Convert.ToString(r("FamilyName")).Trim()
                        If Not String.IsNullOrWhiteSpace(fn) Then famSet.Add(fn)
                    Next
                End If
                ResultTableFilter.KeepOnlyByNameSet(famIndex, "FamilyName", famSet)
            End If

            Dim res As New RunResult() With {
                .Mode = mode,
                .Project = If(projectTable, Auditors.MakeFailureSummaryTable(1)),
                .FamilyDetail = If(includeFamily, familyDetail, Nothing),
                .FamilyIndex = If(includeFamily, famIndex, Nothing),
                .RunId = runId,
                .IncludeFamily = includeFamily
            }
            Return res
        End Function

        Public Shared Function CleanupFromExcel(app As UIApplication,
                                                excelPath As String,
                                                settings As CleanupSettings,
                                                progress As Action(Of Double, String),
                                                Optional warn As Action(Of String) = Nothing) As CleanupResult

            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
            If String.IsNullOrWhiteSpace(excelPath) OrElse Not File.Exists(excelPath) Then
                Throw New FileNotFoundException("삭제용 엑셀 파일을 찾을 수 없습니다.", excelPath)
            End If

            Dim effectiveSettings As CleanupSettings = If(settings, New CleanupSettings())
            Dim instructions = ReadCleanupInstructions(excelPath)
            If instructions.Count = 0 Then
                Throw New InvalidOperationException("삭제용 엑셀에서 '삭제여부'가 입력된 항목을 찾지 못했습니다.")
            End If

            Dim groupedByFile = instructions.
                GroupBy(Function(x) x.RvtPath, StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(g) Path.GetFileName(g.Key), StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim result As New CleanupResult() With {
                .SourceExcelPath = excelPath,
                .InstructionCount = instructions.Count
            }

            Dim total As Integer = Math.Max(1, groupedByFile.Count)
            For i As Integer = 0 To groupedByFile.Count - 1
                Dim group = groupedByFile(i)
                Dim requestedPath As String = If(group.Key, "").Trim()
                Dim fileName As String = SafeFileName(requestedPath)
                Dim fileResult As New CleanupFileResult() With {
                    .FilePath = requestedPath,
                    .FileName = fileName,
                    .RequestedDeleteCount = group.Count()
                }

                Dim openPath As String = requestedPath
                Dim createdLocal As Boolean = False
                Dim doc As Document = Nothing

                Try
                    If String.IsNullOrWhiteSpace(requestedPath) OrElse Not File.Exists(requestedPath) Then
                        fileResult.Status = "Fail"
                        fileResult.Message = "RVT 파일을 찾을 수 없습니다."
                        result.Files.Add(fileResult)
                        Continue For
                    End If

                    If IsAlreadyOpen(app.Application, requestedPath) Then
                        fileResult.Status = "Fail"
                        fileResult.Message = "이미 열려 있는 문서라 정리를 진행할 수 없습니다."
                        result.Files.Add(fileResult)
                        Continue For
                    End If

                    Dim fileInfo = TryExtractBasicFileInfo(requestedPath)
                    If fileInfo IsNot Nothing AndAlso fileInfo.IsCentral Then
                        openPath = CreateNewLocalPath(requestedPath)
                        createdLocal = True
                        fileResult.WasCentralFile = True
                        fileResult.UsedLocalFile = True
                    End If

                    ReportProgress(progress, groupedByFile.Count, i + 1, 0.05R, $"정리 대상 문서 여는 중... {i + 1}/{groupedByFile.Count} {fileName}")
                    doc = OpenProjectDocument(app.Application, openPath, effectiveSettings.CloseAllWorksetsOnOpen)
                    If doc Is Nothing Then
                        fileResult.Status = "Fail"
                        fileResult.Message = "문서를 열지 못했습니다."
                        result.Files.Add(fileResult)
                        Continue For
                    End If

                    Dim notes As New List(Of String)()
                    Dim projectInstructions = group.
                        Where(Function(x) String.Equals(x.DeleteScope, "Project", StringComparison.OrdinalIgnoreCase)).
                        ToList()
                    Dim familyInstructions = group.
                        Where(Function(x) String.Equals(x.DeleteScope, "Family", StringComparison.OrdinalIgnoreCase)).
                        ToList()

                    If projectInstructions.Count > 0 Then
                        ReportProgress(progress, groupedByFile.Count, i + 1, 0.35R, $"[{fileName}] 프로젝트 파라미터 정리 중")
                        fileResult.ProjectDeletedCount = DeleteProjectParameters(doc, requestedPath, projectInstructions, notes)
                    End If

                    If familyInstructions.Count > 0 Then
                        ReportProgress(progress, groupedByFile.Count, i + 1, 0.75R, $"[{fileName}] 로드 패밀리 파라미터 정리 중")
                        fileResult.FamilyDeletedCount = DeleteFamilyParameters(doc, requestedPath, familyInstructions, notes, warn)
                    End If

                    fileResult.DeletedCount = fileResult.ProjectDeletedCount + fileResult.FamilyDeletedCount

                    If fileResult.DeletedCount > 0 Then
                        If doc.IsWorkshared Then
                            Dim syncError As String = String.Empty
                            Dim syncComment As String = If(effectiveSettings.UseSyncComment, If(effectiveSettings.SyncComment, String.Empty), String.Empty)
                            If SyncWithCentral(doc, syncComment, syncError) Then
                                fileResult.SynchronizePerformed = True
                                fileResult.Status = "Success"
                                notes.Add("동기화 완료")
                            Else
                                fileResult.Status = "Fail"
                                notes.Add("동기화 실패: " & syncError)
                            End If
                        Else
                            doc.Save()
                            fileResult.SynchronizePerformed = True
                            fileResult.Status = "Success"
                            notes.Add("저장 완료")
                        End If
                    Else
                        fileResult.Status = "NoChange"
                        If notes.Count = 0 Then
                            notes.Add("삭제된 파라미터가 없습니다.")
                        End If
                    End If

                    fileResult.Message = String.Join(" / ", notes.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))

                Catch ex As Exception
                    fileResult.Status = "Fail"
                    fileResult.Message = ex.Message
                Finally
                    If doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                    If createdLocal Then
                        TryDeleteFile(openPath)
                    End If
                End Try

                result.Files.Add(fileResult)
                ReportProgress(progress, groupedByFile.Count, i + 1, 1.0R, $"정리 완료: {i + 1}/{groupedByFile.Count} {fileName}")
            Next

            result.DeletedCount = result.Files.Sum(Function(x) x.DeletedCount)
            result.SuccessCount = result.Files.Where(Function(x) String.Equals(x.Status, "Success", StringComparison.OrdinalIgnoreCase)).Count()
            result.FailCount = result.Files.Where(Function(x) String.Equals(x.Status, "Fail", StringComparison.OrdinalIgnoreCase)).Count()
            result.NoChangeCount = result.Files.Where(Function(x) String.Equals(x.Status, "NoChange", StringComparison.OrdinalIgnoreCase)).Count()
            result.Ok = result.SuccessCount > 0 OrElse result.NoChangeCount > 0

            If result.FailCount > 0 AndAlso result.SuccessCount = 0 AndAlso result.NoChangeCount = 0 Then
                result.Message = "삭제 정리에 실패했습니다."
            ElseIf result.DeletedCount > 0 Then
                result.Message = $"파라미터 GUID 정리가 완료되었습니다. 삭제 {result.DeletedCount}건"
            Else
                result.Message = "정리할 삭제 대상은 있었지만 실제 삭제된 파라미터는 없습니다."
            End If

            Return result
        End Function

        ''' <summary>엑셀 내보내기 (AutoFit 사용 안 함)</summary>
        Public Shared Function Export(table As DataTable,
                                      sheetName As String,
                                      Optional excelMode As String = "fast",
                                      Optional progressChannel As String = Nothing,
                                      Optional exportLocale As String = "ko") As String
            If table Is Nothing Then Return String.Empty
            Dim doAutoFit As Boolean = False
            Try
                If String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase) AndAlso table.Rows.Count <= 30000 Then
                    doAutoFit = True
                End If
            Catch
                doAutoFit = False
            End Try
            ' ResultTableFilter.KeepOnlyIssues("guid", table)
            ExcelCore.EnsureMessageRow(table, "오류가 없습니다.")

            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = $"{sheetName}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
                If sfd.ShowDialog() <> DialogResult.OK Then Return String.Empty
                ExcelCore.SaveXlsx(sfd.FileName, sheetName, table, doAutoFit, sheetKey:=sheetName, progressKey:=progressChannel, exportKind:="guid", exportLocale:=exportLocale)
                Return sfd.FileName
            End Using
        End Function

        ''' <summary>엑셀 내보내기 (다중 시트)</summary>
        Public Shared Function ExportMulti(sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                           Optional excelMode As String = "fast",
                                           Optional progressChannel As String = Nothing,
                                           Optional exportLocale As String = "ko") As String
            If sheets Is Nothing OrElse sheets.Count = 0 Then Return String.Empty
            Dim doAutoFit As Boolean = False
            Try
                If String.Equals(excelMode, "normal", StringComparison.OrdinalIgnoreCase) Then
                    doAutoFit = True
                End If
            Catch
                doAutoFit = False
            End Try
            For Each kv In sheets
                ' ResultTableFilter.KeepOnlyIssues("guid", kv.Value)
                ExcelCore.EnsureMessageRow(kv.Value, "오류가 없습니다.")
            Next

            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = $"GuidAudit_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
                If sfd.ShowDialog() <> DialogResult.OK Then Return String.Empty
                ExcelCore.SaveXlsxMulti(sfd.FileName, sheets, doAutoFit, progressChannel, exportKind:="guid", exportLocale:=exportLocale)
                Return sfd.FileName
            End Using
        End Function

        Public Shared Function PrepareExportTable(source As DataTable, mode As Integer) As DataTable
            Dim baseTable As DataTable = If(source, Auditors.MakeFailureSummaryTable(mode))
            Dim exportTable As DataTable = baseTable.Clone()

            If source IsNot Nothing Then
                For Each r As DataRow In source.Rows
                    exportTable.ImportRow(r)
                Next
            End If

            ' ResultTableFilter.KeepOnlyIssues("guid", exportTable)

            If exportTable.Columns.Contains("BoundCategories") Then
                exportTable.Columns("BoundCategories").SetOrdinal(exportTable.Columns.Count - 1)
            End If

            EnsureDeleteActionColumn(exportTable)
            AppendHiddenDeleteKeyColumns(exportTable, mode)

            If exportTable.Columns.Contains("RvtPath") Then
                exportTable.Columns.Remove("RvtPath")
            End If

            TrimAndRenameGuidExportColumns(exportTable, mode)

            ExcelCore.EnsureMessageRow(exportTable, "오류가 없습니다.")

            Return exportTable
        End Function

        Private Shared Sub AppendHiddenDeleteKeyColumns(table As DataTable, mode As Integer)
            If table Is Nothing Then Return

            Dim hiddenSpecs As List(Of KeyValuePair(Of String, Func(Of DataRow, Object))) =
                BuildHiddenGuidExportColumns(mode)

            If hiddenSpecs Is Nothing OrElse hiddenSpecs.Count = 0 Then Return

            For Each spec In hiddenSpecs
                If table.Columns.Contains(spec.Key) Then
                    table.Columns.Remove(spec.Key)
                End If

                Dim added As DataColumn = table.Columns.Add(spec.Key, GetType(String))
                ExcelCore.MarkColumnHidden(added)
                added.SetOrdinal(table.Columns.Count - 1)
            Next

            For Each row As DataRow In table.Rows
                For Each spec In hiddenSpecs
                    Dim value As Object = Nothing
                    Try
                        value = spec.Value(row)
                    Catch
                        value = String.Empty
                    End Try
                    row(spec.Key) = If(value, String.Empty)
                Next
            Next
        End Sub

        Private Shared Function BuildHiddenGuidExportColumns(mode As Integer) As List(Of KeyValuePair(Of String, Func(Of DataRow, Object)))
            Dim items As New List(Of KeyValuePair(Of String, Func(Of DataRow, Object)))()

            If mode = 1 Then
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__DeleteScope", Function(r) "Project"))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__RvtPath", Function(r) GetRowString(r, "RvtPath")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__ParamName", Function(r) GetRowString(r, "ParamName")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__ParamKind", Function(r) GetRowString(r, "ParamKind")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__RvtGuid", Function(r) GetRowString(r, "RvtGuid")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__FileGuid", Function(r) GetRowString(r, "FileGuid")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__BoundCategories", Function(r) GetRowString(r, "BoundCategories")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__DeleteKey",
                    Function(r) BuildProjectDeleteKey(r)))
            Else
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__DeleteScope", Function(r) "Family"))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__RvtPath", Function(r) GetRowString(r, "RvtPath")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__FamilyName", Function(r) GetRowString(r, "FamilyName")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__FamilyCategory", Function(r) GetRowString(r, "FamilyCategory")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__ParamName", Function(r) GetRowString(r, "ParamName")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__ParamKind", Function(r) GetRowString(r, "ParamKind")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__FamilyGuid", Function(r) GetRowString(r, "FamilyGuid")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__FileGuid", Function(r) GetRowString(r, "FileGuid")))
                items.Add(New KeyValuePair(Of String, Func(Of DataRow, Object))("__DeleteKey",
                    Function(r) BuildFamilyDeleteKey(r)))
            End If

            Return items
        End Function

        Private Shared Function BuildProjectDeleteKey(row As DataRow) As String
            Return String.Join("|", New String() {
                "Project",
                GetRowString(row, "RvtPath"),
                GetRowString(row, "ParamName"),
                GetRowString(row, "ParamKind"),
                GetRowString(row, "RvtGuid"),
                GetRowString(row, "FileGuid"),
                GetRowString(row, "BoundCategories")
            })
        End Function

        Private Shared Function BuildFamilyDeleteKey(row As DataRow) As String
            Return String.Join("|", New String() {
                "Family",
                GetRowString(row, "RvtPath"),
                GetRowString(row, "FamilyName"),
                GetRowString(row, "ParamName"),
                GetRowString(row, "ParamKind"),
                GetRowString(row, "FamilyGuid"),
                GetRowString(row, "FileGuid")
            })
        End Function

        Private Shared Function GetRowString(row As DataRow, columnName As String) As String
            If row Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return String.Empty
            If row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then Return String.Empty

            Dim value As Object = Nothing
            Try
                value = row(columnName)
            Catch
                value = Nothing
            End Try

            If value Is Nothing OrElse value Is DBNull.Value Then Return String.Empty
            Return Convert.ToString(value).Trim()
        End Function

        Private Shared Sub TrimAndRenameGuidExportColumns(table As DataTable, mode As Integer)
            If table Is Nothing Then Return

            If mode = 1 Then
                RemoveColumnIfExists(table, "ParamGroup")
                RenameColumnIfExists(table, "RvtName", "파일명")
                RenameColumnIfExists(table, "ParamName", "파라미터명")
                RenameColumnIfExists(table, "ParamKind", "구분")
                RenameColumnIfExists(table, "BoundCategories", "적용카테고리")
                RenameColumnIfExists(table, "RvtGuid", "현재 GUID")
                RenameColumnIfExists(table, "FileGuid", "기준 GUID")
                RenameColumnIfExists(table, "Result", "검토결과")
                RenameColumnIfExists(table, "Notes", "비고")

                SetColumnOrder(table, New String() {
                    "삭제여부",
                    "파일명",
                    "파라미터명",
                    "구분",
                    "적용카테고리",
                    "현재 GUID",
                    "기준 GUID",
                    "검토결과",
                    "비고"
                })
            Else
                RemoveColumnIfExists(table, "IsShared")
                RenameColumnIfExists(table, "RvtName", "파일명")
                RenameColumnIfExists(table, "FamilyName", "패밀리명")
                RenameColumnIfExists(table, "FamilyCategory", "패밀리카테고리")
                RenameColumnIfExists(table, "ParamName", "파라미터명")
                RenameColumnIfExists(table, "ParamKind", "구분")
                RenameColumnIfExists(table, "FamilyGuid", "현재 GUID")
                RenameColumnIfExists(table, "FileGuid", "기준 GUID")
                RenameColumnIfExists(table, "Result", "검토결과")
                RenameColumnIfExists(table, "Notes", "비고")

                SetColumnOrder(table, New String() {
                    "삭제여부",
                    "파일명",
                    "패밀리명",
                    "패밀리카테고리",
                    "파라미터명",
                    "구분",
                    "현재 GUID",
                    "기준 GUID",
                    "검토결과",
                    "비고"
                })
            End If

            TranslateGuidExportValues(table)
        End Sub

        Private Shared Sub TranslateGuidExportValues(table As DataTable)
            If table Is Nothing OrElse table.Rows.Count = 0 Then Return

            For Each row As DataRow In table.Rows
                If table.Columns.Contains("구분") Then
                    row("구분") = TranslateParamKind(GetRowString(row, "구분"))
                End If

                If table.Columns.Contains("검토결과") Then
                    row("검토결과") = TranslateGuidResult(GetRowString(row, "검토결과"))
                End If
            Next
        End Sub

        Private Shared Function TranslateParamKind(value As String) As String
            Select Case If(value, "").Trim().ToUpperInvariant()
                Case "SHARED"
                    Return "공유"
                Case "PROJECT"
                    Return "프로젝트"
                Case "BUILTIN"
                    Return "내장"
                Case "FAMILY", "FAMILY_PARAM"
                    Return "패밀리"
                Case Else
                    Return value
            End Select
        End Function

        Private Shared Function TranslateGuidResult(value As String) As String
            Select Case If(value, "").Trim().ToUpperInvariant()
                Case "OK"
                    Return "일치"
                Case "OK(MULTI_IN_FILE)"
                    Return "일치(기준파일 중복)"
                Case "MISMATCH"
                    Return "불일치"
                Case "NOT_FOUND_IN_FILE"
                    Return "기준파일 없음"
                Case "PROJECT_PARAM"
                    Return "프로젝트 파라미터"
                Case "FAMILY_PARAM"
                    Return "패밀리 파라미터"
                Case "BUILTIN"
                    Return "내장 파라미터"
                Case "GUID_FAIL"
                    Return "GUID 추출 실패"
                Case "OPEN_FAIL"
                    Return "열기 실패"
                Case "ERROR"
                    Return "오류"
                Case Else
                    Return value
            End Select
        End Function

        Private Shared Sub RenameColumnIfExists(table As DataTable, sourceName As String, targetName As String)
            If table Is Nothing OrElse String.IsNullOrWhiteSpace(sourceName) OrElse String.IsNullOrWhiteSpace(targetName) Then Return
            If Not table.Columns.Contains(sourceName) Then Return
            If table.Columns.Contains(targetName) Then Return
            table.Columns(sourceName).ColumnName = targetName
        End Sub

        Private Shared Sub RemoveColumnIfExists(table As DataTable, columnName As String)
            If table Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return
            If table.Columns.Contains(columnName) Then
                table.Columns.Remove(columnName)
            End If
        End Sub

        Private Shared Sub SetColumnOrder(table As DataTable, orderedColumns As IEnumerable(Of String))
            If table Is Nothing OrElse orderedColumns Is Nothing Then Return

            Dim ordinal As Integer = 0
            For Each columnName As String In orderedColumns
                If String.IsNullOrWhiteSpace(columnName) Then Continue For
                If Not table.Columns.Contains(columnName) Then Continue For
                table.Columns(columnName).SetOrdinal(ordinal)
                ordinal += 1
            Next
        End Sub

        Private Shared Sub EnsureDeleteActionColumn(table As DataTable)
            If table Is Nothing Then Return

            Dim columnName As String = "삭제여부"
            If Not table.Columns.Contains(columnName) Then
                table.Columns.Add(columnName, GetType(String))
            End If

            For Each row As DataRow In table.Rows
                row(columnName) = String.Empty
            Next

            table.Columns(columnName).SetOrdinal(0)
        End Sub

        Private Shared Function ReadCleanupInstructions(excelPath As String) As List(Of CleanupInstruction)
            Dim results As New List(Of CleanupInstruction)()
            Dim formatter As New DataFormatter()

            Using stream As New FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using workbook As IWorkbook = New XSSFWorkbook(stream)
                    For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                        Dim sheet = workbook.GetSheetAt(sheetIndex)
                        If sheet Is Nothing Then Continue For

                        Dim headerRow = sheet.GetRow(sheet.FirstRowNum)
                        If headerRow Is Nothing OrElse headerRow.LastCellNum <= 0 Then Continue For

                        Dim headers As New Dictionary(Of Integer, String)()
                        For col As Integer = 0 To CInt(headerRow.LastCellNum) - 1
                            headers(col) = formatter.FormatCellValue(headerRow.GetCell(col)).Trim()
                        Next

                        Dim deleteCol = FindHeaderIndex(headers, "삭제여부")
                        If deleteCol < 0 Then deleteCol = FindHeaderIndex(headers, "Delete")
                        Dim scopeCol = FindHeaderIndex(headers, "__DeleteScope")
                        Dim keyCol = FindHeaderIndex(headers, "__DeleteKey")
                        Dim pathCol = FindHeaderIndex(headers, "__RvtPath")

                        If deleteCol < 0 OrElse scopeCol < 0 OrElse keyCol < 0 OrElse pathCol < 0 Then Continue For

                        For rowIndex As Integer = sheet.FirstRowNum + 1 To sheet.LastRowNum
                            Dim row = sheet.GetRow(rowIndex)
                            If row Is Nothing Then Continue For

                            Dim deleteValue As String = GetCellText(row.GetCell(deleteCol), formatter)
                            If Not IsDeleteAction(deleteValue) Then Continue For

                            Dim instruction As New CleanupInstruction() With {
                                .DeleteScope = GetCellText(row.GetCell(scopeCol), formatter),
                                .DeleteKey = GetCellText(row.GetCell(keyCol), formatter),
                                .RvtPath = GetCellText(row.GetCell(pathCol), formatter),
                                .ParamName = GetCellTextByHeader(headers, row, formatter, "__ParamName"),
                                .ParamKind = GetCellTextByHeader(headers, row, formatter, "__ParamKind"),
                                .RvtGuid = GetCellTextByHeader(headers, row, formatter, "__RvtGuid"),
                                .FileGuid = GetCellTextByHeader(headers, row, formatter, "__FileGuid"),
                                .BoundCategories = GetCellTextByHeader(headers, row, formatter, "__BoundCategories"),
                                .FamilyName = GetCellTextByHeader(headers, row, formatter, "__FamilyName"),
                                .FamilyCategory = GetCellTextByHeader(headers, row, formatter, "__FamilyCategory"),
                                .FamilyGuid = GetCellTextByHeader(headers, row, formatter, "__FamilyGuid"),
                                .SheetName = sheet.SheetName,
                                .RowNumber = rowIndex + 1
                            }

                            If String.IsNullOrWhiteSpace(instruction.DeleteScope) OrElse
                               String.IsNullOrWhiteSpace(instruction.DeleteKey) OrElse
                               String.IsNullOrWhiteSpace(instruction.RvtPath) Then
                                Continue For
                            End If

                            results.Add(instruction)
                        Next
                    Next
                End Using
            End Using

            Return results.
                GroupBy(Function(x) x.DeleteKey, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.First()).
                ToList()
        End Function

        Private Shared Function FindHeaderIndex(headers As IDictionary(Of Integer, String), headerName As String) As Integer
            If headers Is Nothing OrElse String.IsNullOrWhiteSpace(headerName) Then Return -1
            For Each kv In headers
                If String.Equals(If(kv.Value, "").Trim(), headerName, StringComparison.OrdinalIgnoreCase) Then
                    Return kv.Key
                End If
            Next
            Return -1
        End Function

        Private Shared Function GetCellText(cell As ICell, formatter As DataFormatter) As String
            If cell Is Nothing Then Return String.Empty
            If formatter Is Nothing Then
                Try
                    Return If(cell.ToString(), String.Empty).Trim()
                Catch
                    Return String.Empty
                End Try
            End If

            Try
                Return formatter.FormatCellValue(cell).Trim()
            Catch
                Try
                    Return If(cell.ToString(), String.Empty).Trim()
                Catch
                    Return String.Empty
                End Try
            End Try
        End Function

        Private Shared Function GetCellTextByHeader(headers As IDictionary(Of Integer, String),
                                                    row As IRow,
                                                    formatter As DataFormatter,
                                                    headerName As String) As String
            If row Is Nothing Then Return String.Empty
            Dim colIndex As Integer = FindHeaderIndex(headers, headerName)
            If colIndex < 0 Then Return String.Empty
            Return GetCellText(row.GetCell(colIndex), formatter)
        End Function

        Private Shared Function IsDeleteAction(value As String) As Boolean
            Select Case If(value, "").Trim().ToUpperInvariant()
                Case "DELETE", "삭제", "Y", "YES", "1", "TRUE"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function DeleteProjectParameters(doc As Document,
                                                        requestedPath As String,
                                                        instructions As IList(Of CleanupInstruction),
                                                        notes As IList(Of String)) As Integer
            If doc Is Nothing OrElse instructions Is Nothing OrElse instructions.Count = 0 Then Return 0

            Dim candidates = CollectProjectBindingCandidates(doc)
            If candidates.Count = 0 Then
                If notes IsNot Nothing Then notes.Add("삭제 가능한 프로젝트 파라미터를 찾지 못했습니다.")
                Return 0
            End If

            Dim deleted As Integer = 0
            Using tx As New Transaction(doc, "KKY GUID Cleanup - Project Parameters")
                tx.Start()

                For Each instruction In instructions
                    Dim candidate = candidates.FirstOrDefault(Function(x) ProjectCandidateMatches(x, instruction))
                    If candidate Is Nothing Then
                        If notes IsNot Nothing Then notes.Add($"프로젝트 파라미터 없음: {instruction.ParamName}")
                        Continue For
                    End If

                    Try
                        If doc.ParameterBindings.Remove(candidate.Definition) Then
                            deleted += 1
                        ElseIf notes IsNot Nothing Then
                            notes.Add($"프로젝트 파라미터 삭제 실패: {instruction.ParamName}")
                        End If
                    Catch ex As Exception
                        If notes IsNot Nothing Then notes.Add($"프로젝트 파라미터 삭제 실패: {instruction.ParamName} - {ex.Message}")
                    End Try
                Next

                If deleted > 0 Then
                    tx.Commit()
                Else
                    tx.RollBack()
                End If
            End Using

            Return deleted
        End Function

        Private Shared Function CollectProjectBindingCandidates(doc As Document) As List(Of ProjectBindingCandidate)
            Dim list As New List(Of ProjectBindingCandidate)()
            If doc Is Nothing Then Return list

            Dim sharedByName As New Dictionary(Of String, Guid)(StringComparer.OrdinalIgnoreCase)
            Try
                For Each spe As SharedParameterElement In New FilteredElementCollector(doc).OfClass(GetType(SharedParameterElement)).Cast(Of SharedParameterElement)()
                    Dim key As String = NormalizeName(SafeParamElementName(spe))
                    If String.IsNullOrWhiteSpace(key) OrElse sharedByName.ContainsKey(key) Then Continue For
                    sharedByName(key) = spe.GuidValue
                Next
            Catch
            End Try

            Dim iter As DefinitionBindingMapIterator = doc.ParameterBindings.ForwardIterator()
            iter.Reset()
            While iter.MoveNext()
                Dim def As Definition = Nothing
                Try
                    def = iter.Key
                Catch
                    def = Nothing
                End Try
                If def Is Nothing Then Continue While

                Dim candidate As New ProjectBindingCandidate() With {
                    .Definition = def,
                    .ParamName = If(def.Name, String.Empty),
                    .ParamKind = "Project",
                    .ParamGuid = String.Empty
                }

                If TypeOf def Is ExternalDefinition Then
                    candidate.ParamKind = "Shared"
                    Try
                        candidate.ParamGuid = DirectCast(def, ExternalDefinition).GUID.ToString()
                    Catch
                        candidate.ParamGuid = String.Empty
                    End Try
                Else
                    Dim sharedGuid As Guid = Guid.Empty
                    If sharedByName.TryGetValue(NormalizeName(candidate.ParamName), sharedGuid) Then
                        candidate.ParamKind = "Shared"
                        If sharedGuid <> Guid.Empty Then candidate.ParamGuid = sharedGuid.ToString()
                    End If
                End If

                list.Add(candidate)
            End While

            Return list
        End Function

        Private Shared Function ProjectCandidateMatches(candidate As ProjectBindingCandidate, instruction As CleanupInstruction) As Boolean
            If candidate Is Nothing OrElse instruction Is Nothing Then Return False
            If Not String.Equals(NormalizeName(candidate.ParamName), NormalizeName(instruction.ParamName), StringComparison.OrdinalIgnoreCase) Then Return False
            If Not String.Equals(If(candidate.ParamKind, ""), If(instruction.ParamKind, ""), StringComparison.OrdinalIgnoreCase) Then Return False

            If String.Equals(candidate.ParamKind, "Shared", StringComparison.OrdinalIgnoreCase) Then
                Return String.Equals(If(candidate.ParamGuid, ""), If(instruction.RvtGuid, ""), StringComparison.OrdinalIgnoreCase)
            End If

            Return True
        End Function

        Private Shared Function DeleteFamilyParameters(doc As Document,
                                                       requestedPath As String,
                                                       instructions As IList(Of CleanupInstruction),
                                                       notes As IList(Of String),
                                                       warn As Action(Of String)) As Integer
            If doc Is Nothing OrElse instructions Is Nothing OrElse instructions.Count = 0 Then Return 0

            Dim deleted As Integer = 0
            Dim familyGroups = instructions.
                GroupBy(Function(x) x.FamilyName, StringComparer.OrdinalIgnoreCase).
                ToList()

            For Each familyGroup In familyGroups
                Dim familyName As String = If(familyGroup.Key, "").Trim()
                If String.IsNullOrWhiteSpace(familyName) Then Continue For

                Dim family As Family = FindEditableFamily(doc, familyName)
                If family Is Nothing Then
                    If notes IsNot Nothing Then notes.Add($"로드된 패밀리를 찾지 못했습니다: {familyName}")
                    Continue For
                End If

                Dim blockedInstructions = familyGroup.
                    Where(Function(x) String.Equals(x.ParamKind, "BuiltIn", StringComparison.OrdinalIgnoreCase)).
                    ToList()
                For Each blocked In blockedInstructions
                    If notes IsNot Nothing Then notes.Add($"내장 파라미터는 삭제할 수 없습니다: {familyName} / {blocked.ParamName}")
                Next

                Dim editableInstructions = familyGroup.
                    Where(Function(x) Not String.Equals(x.ParamKind, "BuiltIn", StringComparison.OrdinalIgnoreCase)).
                    ToList()
                If editableInstructions.Count = 0 Then Continue For

                Dim famDoc As Document = Nothing
                Try
                    famDoc = doc.EditFamily(family)
                    If famDoc Is Nothing OrElse Not famDoc.IsFamilyDocument Then
                        If notes IsNot Nothing Then notes.Add($"패밀리 편집 문서를 열지 못했습니다: {familyName}")
                        Continue For
                    End If

                    Dim fm As FamilyManager = famDoc.FamilyManager
                    If fm Is Nothing Then
                        If notes IsNot Nothing Then notes.Add($"FamilyManager 없음: {familyName}")
                        Continue For
                    End If

                    Dim deletedInFamily As Integer = 0
                    Using tx As New Transaction(famDoc, "KKY GUID Cleanup - Family Parameters")
                        tx.Start()

                        For Each instruction In editableInstructions
                            Dim matchedParameter As FamilyParameter = Nothing
                            For Each fp As FamilyParameter In fm.Parameters
                                If FamilyParameterMatches(fp, instruction) Then
                                    matchedParameter = fp
                                    Exit For
                                End If
                            Next

                            If matchedParameter Is Nothing Then
                                If notes IsNot Nothing Then notes.Add($"패밀리 파라미터 없음: {familyName} / {instruction.ParamName}")
                                Continue For
                            End If

                            Try
                                fm.RemoveParameter(matchedParameter)
                                deletedInFamily += 1
                            Catch ex As Exception
                                If notes IsNot Nothing Then notes.Add($"패밀리 파라미터 삭제 실패: {familyName} / {instruction.ParamName} - {ex.Message}")
                            End Try
                        Next

                        If deletedInFamily > 0 Then
                            tx.Commit()
                        Else
                            tx.RollBack()
                        End If
                    End Using

                    If deletedInFamily > 0 Then
                        famDoc.LoadFamily(doc, New CleanupFamilyLoadOptions())
                        deleted += deletedInFamily
                    End If
                Catch ex As Exception
                    If notes IsNot Nothing Then notes.Add($"패밀리 정리 실패: {familyName} - {ex.Message}")
                    If warn IsNot Nothing Then warn($"패밀리 정리 실패: {familyName} - {ex.Message}")
                Finally
                    If famDoc IsNot Nothing Then
                        Try
                            famDoc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            Return deleted
        End Function

        Private Shared Function FindEditableFamily(doc As Document, familyName As String) As Family
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(familyName) Then Return Nothing

            Return New FilteredElementCollector(doc).
                OfClass(GetType(Family)).
                Cast(Of Family)().
                FirstOrDefault(Function(f)
                                   If f Is Nothing Then Return False
                                   If Not String.Equals(If(f.Name, ""), familyName, StringComparison.OrdinalIgnoreCase) Then Return False
                                   Try
                                       If f.IsInPlace Then Return False
                                   Catch
                                   End Try
                                   Try
                                       Return f.IsEditable
                                   Catch
                                       Return True
                                   End Try
                               End Function)
        End Function

        Private Shared Function FamilyParameterMatches(fp As FamilyParameter, instruction As CleanupInstruction) As Boolean
            If fp Is Nothing OrElse instruction Is Nothing Then Return False

            Dim name As String = String.Empty
            Try
                name = fp.Definition.Name
            Catch
                name = String.Empty
            End Try
            If Not String.Equals(NormalizeName(name), NormalizeName(instruction.ParamName), StringComparison.OrdinalIgnoreCase) Then Return False

            Dim kind As String = GetFamilyParamKind(fp)
            If Not String.Equals(kind, instruction.ParamKind, StringComparison.OrdinalIgnoreCase) Then Return False

            If String.Equals(kind, "Shared", StringComparison.OrdinalIgnoreCase) Then
                Dim guidValue As Guid = Guid.Empty
                If Not TryGetFamilyParameterGuid(fp, guidValue) Then Return False
                Return String.Equals(guidValue.ToString(), If(instruction.FamilyGuid, ""), StringComparison.OrdinalIgnoreCase)
            End If

            Return True
        End Function

        Private Shared Function SafeParamElementName(pe As Element) As String
            Try
                Return pe.Name
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function GetFamilyParamKind(fp As FamilyParameter) As String
            If fp Is Nothing Then Return "None"
            Dim isSharedFlag As Boolean = False
            Try : isSharedFlag = fp.IsShared : Catch : isSharedFlag = False : End Try
            If isSharedFlag Then Return "Shared"
            Dim idVal As Integer = 0
            Try : idVal = fp.Id.IntegerValue : Catch : idVal = 0 : End Try
            If idVal < 0 Then Return "BuiltIn"
            Return "Family"
        End Function

        Private Shared Function TryGetFamilyParameterGuid(fp As FamilyParameter, ByRef g As Guid) As Boolean
            g = Guid.Empty
            If fp Is Nothing Then Return False

            Dim t = fp.GetType()
            Dim p = t.GetProperty("GUID", BindingFlags.Public Or BindingFlags.Instance)
            If p Is Nothing Then Return False

            Dim v = p.GetValue(fp, Nothing)
            If v Is Nothing Then Return False

            If TypeOf v Is Guid Then
                g = DirectCast(v, Guid)
                Return g <> Guid.Empty
            End If

            Return False
        End Function

        Private Shared Function OpenProjectDocument(app As RevitApp,
                                                    userVisiblePath As String,
                                                    closeAllWorksets As Boolean) As Document
            Dim modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)
            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If closeAllWorksets Then
                Dim worksetConfig As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                openOpts.SetOpenWorksetsConfiguration(worksetConfig)
            End If

            Return app.OpenDocumentFile(modelPath, openOpts)
        End Function

        Private Shared Function SyncWithCentral(doc As Document,
                                                comment As String,
                                                ByRef err As String) As Boolean
            err = String.Empty
            If doc Is Nothing OrElse Not doc.IsWorkshared Then
                err = "Workshared 문서가 아닙니다."
                Return False
            End If

            Try
                Dim twc As New TransactWithCentralOptions()
                Dim swc As New SynchronizeWithCentralOptions()
                swc.Comment = If(comment, String.Empty)
                Try
                    Dim relinquish As New RelinquishOptions(True)
                    swc.SetRelinquishOptions(relinquish)
                Catch
                End Try

                doc.SynchronizeWithCentral(twc, swc)
                Return True
            Catch ex As Exception
                err = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function CreateNewLocalPath(centralPath As String) As String
            Dim localRoot = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "GuidCleanup", DateTime.Now.ToString("yyyyMMdd"))
            Directory.CreateDirectory(localRoot)

            Dim fileName = Path.GetFileNameWithoutExtension(centralPath) & "_" & Environment.UserName & "_" & DateTime.Now.ToString("HHmmssfff") & ".rvt"
            Dim localPath = Path.Combine(localRoot, fileName)

            Dim sourcePath = ModelPathUtils.ConvertUserVisiblePathToModelPath(centralPath)
            Dim targetPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath)
            WorksharingUtils.CreateNewLocal(sourcePath, targetPath)
            Return localPath
        End Function

        Private Shared Function TryExtractBasicFileInfo(pathText As String) As BasicFileInfo
            Try
                Return BasicFileInfo.Extract(pathText)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function IsAlreadyOpen(app As RevitApp, userVisiblePath As String) As Boolean
            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For
                    If String.Equals(doc.PathName, userVisiblePath, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
            End Try
            Return False
        End Function

        Private Shared Sub TryDeleteFile(pathText As String)
            If String.IsNullOrWhiteSpace(pathText) Then Return
            Try
                If File.Exists(pathText) Then
                    File.Delete(pathText)
                End If
            Catch
            End Try
        End Sub

        Private NotInheritable Class CleanupFamilyLoadOptions
            Implements IFamilyLoadOptions

            Public Function OnFamilyFound(familyInUse As Boolean,
                                          ByRef overwriteParameterValues As Boolean) As Boolean _
                Implements IFamilyLoadOptions.OnFamilyFound
                overwriteParameterValues = True
                Return True
            End Function

            Public Function OnSharedFamilyFound(sharedFamily As Family,
                                                familyInUse As Boolean,
                                                ByRef source As FamilySource,
                                                ByRef overwriteParameterValues As Boolean) As Boolean _
                Implements IFamilyLoadOptions.OnSharedFamilyFound
                source = FamilySource.Family
                overwriteParameterValues = True
                Return True
            End Function
        End Class

        Private Shared Function BuildTargets(app As UIApplication, rvtPaths As IEnumerable(Of String)) As List(Of TargetFile)
            Dim list As New List(Of TargetFile)()
            Dim dedup As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim requested As IEnumerable(Of String) = If(rvtPaths, Enumerable.Empty(Of String)())
            For Each p In requested
                If String.IsNullOrWhiteSpace(p) Then Continue For
                Dim full As String = p
                Try
                    If System.IO.Path.IsPathRooted(p) Then
                        full = System.IO.Path.GetFullPath(p)
                    Else
                        full = p.Trim()
                    End If
                Catch
                    full = p
                End Try
                If dedup.Add(full) Then
                    list.Add(New TargetFile() With {.Path = full, .Name = SafeFileName(full)})
                End If
            Next

            If list.Count = 0 Then
                Dim ap As String = ""
                Try : ap = app.ActiveUIDocument?.Document?.PathName : Catch : ap = "" : End Try
                list.Add(New TargetFile() With {.Path = ap, .Name = SafeFileName(ap)})
            End If

            Return list
        End Function

        Private Shared Function SafeFileName(p As String) As String
            If String.IsNullOrWhiteSpace(p) Then Return "(Active/Unsaved)"
            Try
                Return System.IO.Path.GetFileName(p)
            Catch
                Return p
            End Try
        End Function

        Private Shared Function GetRvtName(doc As Document, path As String) As String
            If Not String.IsNullOrWhiteSpace(path) Then
                Try
                    Return System.IO.Path.GetFileName(path)
                Catch
                End Try
            End If
            Try
                Return doc.Title
            Catch
                Return "(Doc)"
            End Try
        End Function

        Private Shared Function MergeTable(master As DataTable, part As DataTable) As DataTable
            If part Is Nothing Then Return master
            If master Is Nothing Then master = part.Clone()
            For Each r As DataRow In part.Rows
                master.ImportRow(r)
            Next
            Return master
        End Function

        Private Shared Function SafeRatio(cur As Integer, tot As Integer) As Double
            If tot <= 0 Then Return 0
            Return Math.Max(0, Math.Min(1.0R, CDbl(cur) / CDbl(tot)))
        End Function

        Private Shared Function NormalizeName(s As String) As String
            If s Is Nothing Then Return String.Empty
            Dim value As String = s.Replace(ChrW(&HA0), " ")
            value = value.Trim()
            If value.Length = 0 Then Return String.Empty
            Try
                value = Regex.Replace(value, "\s+", " ")
            Catch
            End Try
            Return value
        End Function

        Private Shared Sub ReportProgress(cb As Action(Of Double, String),
                                          totalFiles As Integer,
                                          fileIndex As Integer,
                                          docProgress As Double,
                                          text As String)
            If cb Is Nothing Then Return
            Dim safeTotal As Integer = Math.Max(1, totalFiles)
            Dim idx As Integer = Math.Max(0, fileIndex - 1)
            Dim ratio As Double = (idx + Math.Max(0.0R, Math.Min(1.0R, docProgress))) / safeTotal
            Dim pct As Double = Math.Max(0, Math.Min(100, Math.Round(ratio * 1000.0R) / 10.0R))
            cb(pct, text)
        End Sub

        Private Shared Function BuildOpenFailNotes(reason As String, inputPath As String) As String
            Dim trimmed = If(reason, "").Trim()
            Dim hasPathInReason As Boolean = False
            Try
                hasPathInReason = Not String.IsNullOrWhiteSpace(inputPath) AndAlso
                                  trimmed.IndexOf(inputPath, StringComparison.OrdinalIgnoreCase) >= 0
            Catch
                hasPathInReason = False
            End Try
            Dim pathPart = If(String.IsNullOrWhiteSpace(inputPath) OrElse hasPathInReason, "", $" [Path: {inputPath}]")
            If String.IsNullOrWhiteSpace(trimmed) Then
                Return $"문서 열기 실패{pathPart}"
            End If
            Return $"{trimmed}{pathPart}"
        End Function

        Private Shared Function BuildExceptionNotes(ex As Exception, inputPath As String) As String
            If ex Is Nothing Then Return BuildOpenFailNotes(String.Empty, inputPath)

            Dim hrPart As String = ""
            Try
                hrPart = $" (0x{ex.HResult:X8})"
            Catch
                hrPart = ""
            End Try

            Return BuildOpenFailNotes($"{ex.Message}{hrPart}", inputPath)
        End Function

        Private Shared Function ShortenReason(reason As String) As String
            If String.IsNullOrWhiteSpace(reason) Then Return String.Empty
            Dim firstLine As String = reason.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ").Trim()
            If firstLine.Length > 120 Then
                Return firstLine.Substring(0, 117) & "..."
            End If
            Return firstLine
        End Function

        '=========================================================
        ' Central(Workshared) => Detach + CloseAllWorksets
        '=========================================================
        Private Shared Function ResolveOrOpenDocument(uiApp As UIApplication, activeDoc As Document, path As String, ByRef openedByMe As Boolean, ByRef failureReason As String) As Document
            openedByMe = False
            failureReason = String.Empty

            Dim requested As String = If(path, "").Trim()

            Dim isRooted As Boolean = False
            Try
                isRooted = System.IO.Path.IsPathRooted(requested)
            Catch
                isRooted = False
            End Try

            Dim allowNameMatch As Boolean = (Not isRooted) AndAlso requested.IndexOf(":"c) = -1 AndAlso requested.IndexOf("\"c) = -1

            If String.IsNullOrWhiteSpace(requested) Then
                Return activeDoc
            End If

            If IsMatchingDoc(activeDoc, requested, allowNameMatch) Then
                Return activeDoc
            End If

            Dim opened = FindOpenDocument(uiApp, requested, allowNameMatch)
            If opened IsNot Nothing Then Return opened

            If allowNameMatch Then
                failureReason = $"Invalid path: {requested}"
                Return Nothing
            End If

            If Not isRooted Then
                failureReason = $"Invalid path: {requested}"
                Return Nothing
            End If

            Try
                If System.IO.Path.IsPathRooted(requested) AndAlso Not File.Exists(requested) Then
                    failureReason = $"File not found: {requested}"
                    Return Nothing
                End If
            Catch ex As Exception
                failureReason = BuildExceptionNotes(ex, requested)
                Return Nothing
            End Try

            Dim mp As ModelPath = Nothing
            Try
                mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(requested)
            Catch ex As Exception
                failureReason = BuildExceptionNotes(ex, requested)
                mp = Nothing
            End Try
            If mp Is Nothing Then
                If String.IsNullOrWhiteSpace(failureReason) Then failureReason = $"경로 변환 실패 [Path: {requested}]"
                Return Nothing
            End If

            Dim preferDetach As Boolean = False
            Try
                Dim bfi = BasicFileInfo.Extract(requested)
                If bfi Is Nothing Then
                    preferDetach = True
                ElseIf bfi.IsWorkshared Then
                    preferDetach = True
                End If
            Catch
                preferDetach = True
            End Try

            Dim attempts As New List(Of OpenOptions)()
            If preferDetach Then attempts.Add(CreateDetachOptions())
            attempts.Add(New OpenOptions())

            Dim app = uiApp.Application
            For Each opt In attempts
                Try
                    Dim d = app.OpenDocumentFile(mp, opt)
                    openedByMe = True
                    failureReason = String.Empty
                    Return d
                Catch ex As Exception
                    failureReason = BuildExceptionNotes(ex, requested)
                End Try
            Next

            openedByMe = False
            Return Nothing
        End Function

        Private Shared Function CreateDetachOptions() As OpenOptions
            Dim opt As New OpenOptions()
            Try
                opt.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                Dim wc As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opt.SetOpenWorksetsConfiguration(wc)
            Catch
            End Try
            Return opt
        End Function

        Private Shared Function FindOpenDocument(uiApp As UIApplication, requested As String, allowNameMatch As Boolean) As Document
            If uiApp Is Nothing Then Return Nothing
            Try
                For Each d As Document In uiApp.Application.Documents
                    If IsMatchingDoc(d, requested, allowNameMatch) Then Return d
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function IsMatchingDoc(doc As Document, requested As String, allowNameMatch As Boolean) As Boolean
            If doc Is Nothing Then Return False

            Dim dp As String = ""
            Try : dp = doc.PathName : Catch : dp = "" : End Try
            If Not String.IsNullOrWhiteSpace(dp) AndAlso String.Equals(dp, requested, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            If allowNameMatch Then
                Dim fileOnly As String = ""
                Try
                    fileOnly = Path.GetFileName(dp)
                Catch
                    fileOnly = ""
                End Try
                If Not String.IsNullOrWhiteSpace(fileOnly) AndAlso String.Equals(fileOnly, requested, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                Dim title As String = ""
                Try : title = doc.Title : Catch : title = "" : End Try
                If Not String.IsNullOrWhiteSpace(title) AndAlso String.Equals(title, requested, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            End If

            Return False
        End Function

        '=========================================================
        ' 내부: Audit 로직 (기존 구현 이동)
        '=========================================================
        Private NotInheritable Class SharedParamReader

            Public Shared Function ReadSharedParamNameGuidMap(app As Autodesk.Revit.ApplicationServices.Application) As Dictionary(Of String, List(Of Guid))
                Dim defFile As DefinitionFile = Nothing
                Try
                    defFile = app.OpenSharedParameterFile()
                Catch
                    defFile = Nothing
                End Try

                If defFile Is Nothing Then Return Nothing

                Dim map As New Dictionary(Of String, List(Of Guid))(StringComparer.OrdinalIgnoreCase)

                For Each grp As DefinitionGroup In defFile.Groups
                    For Each d As Definition In grp.Definitions
                        Dim g As Guid = Guid.Empty
                        If Not TryGetDefinitionGuid(d, g) Then Continue For

                        Dim name = NormalizeName(d.Name)
                        If Not map.ContainsKey(name) Then map(name) = New List(Of Guid)()
                        map(name).Add(g)
                    Next
                Next

                Return map
            End Function

            Private Shared Function TryGetDefinitionGuid(d As Definition, ByRef g As Guid) As Boolean
                g = Guid.Empty
                If d Is Nothing Then Return False

                Dim t = d.GetType()
                Dim p = t.GetProperty("GUID", BindingFlags.Public Or BindingFlags.Instance)
                If p Is Nothing Then Return False

                Dim v = p.GetValue(d, Nothing)
                If v Is Nothing Then Return False

                If TypeOf v Is Guid Then
                    g = DirectCast(v, Guid)
                    Return g <> Guid.Empty
                End If

                Return False
            End Function

        End Class

        Private NotInheritable Class FamilyAuditPack
            Public Property Summary As DataTable
            Public Property Detail As DataTable
            Public Property Index As DataTable
        End Class

        Private NotInheritable Class Auditors

            Public Shared Function MakeFailureSummaryTable(mode As Integer) As DataTable
                If mode = 1 Then
                    Dim dt As New DataTable("ProjectParams")
                    dt.Columns.Add("RvtName", GetType(String))
                    dt.Columns.Add("RvtPath", GetType(String))
                    dt.Columns.Add("ParamName", GetType(String))
                    dt.Columns.Add("ParamKind", GetType(String))
                    dt.Columns.Add("ParamGroup", GetType(String))
                    dt.Columns.Add("BoundCategories", GetType(String))
                    dt.Columns.Add("RvtGuid", GetType(String))
                    dt.Columns.Add("FileGuid", GetType(String))
                    dt.Columns.Add("Result", GetType(String))
                    dt.Columns.Add("Notes", GetType(String))
                    Return dt
                Else
                    Dim dt As New DataTable("FamilyParams")
                    dt.Columns.Add("RvtName", GetType(String))
                    dt.Columns.Add("RvtPath", GetType(String))
                    dt.Columns.Add("FamilyName", GetType(String))
                    dt.Columns.Add("FamilyCategory", GetType(String))
                    dt.Columns.Add("ParamName", GetType(String))
                    dt.Columns.Add("IsShared", GetType(String))
                    dt.Columns.Add("FamilyGuid", GetType(String))
                    dt.Columns.Add("FileGuid", GetType(String))
                    dt.Columns.Add("Result", GetType(String))
                    dt.Columns.Add("Notes", GetType(String))
                    Return dt
                End If
            End Function

            Public Shared Sub AddOpenFailRow(dt As DataTable, rvtName As String, rvtPath As String, scope As String, result As String, notes As String)
                Dim r = dt.NewRow()
                If dt.Columns.Contains("RvtName") Then r("RvtName") = If(rvtName, "")
                If dt.Columns.Contains("RvtPath") Then r("RvtPath") = If(rvtPath, "")
                If dt.Columns.Contains("FamilyName") Then r("FamilyName") = ""
                If dt.Columns.Contains("FamilyCategory") Then r("FamilyCategory") = ""
                If dt.Columns.Contains("ParamName") Then r("ParamName") = ""
                If dt.Columns.Contains("ParamKind") Then r("ParamKind") = ""
                If dt.Columns.Contains("ParamGroup") Then r("ParamGroup") = ""
                If dt.Columns.Contains("BoundCategories") Then r("BoundCategories") = ""
                If dt.Columns.Contains("RvtGuid") Then r("RvtGuid") = ""
                If dt.Columns.Contains("IsShared") Then r("IsShared") = ""
                If dt.Columns.Contains("FamilyGuid") Then r("FamilyGuid") = ""
                If dt.Columns.Contains("FileGuid") Then r("FileGuid") = ""
                If dt.Columns.Contains("Result") Then r("Result") = result
                If dt.Columns.Contains("Notes") Then r("Notes") = notes
                dt.Rows.Add(r)
            End Sub

            Public Shared Function RunProjectParameterAudit(doc As Document,
                                                            fileMap As Dictionary(Of String, List(Of Guid)),
                                                            rvtName As String,
                                                            rvtPath As String,
                                                            Optional progress As Action(Of Integer, Integer) = Nothing) As DataTable

                Dim dt As New DataTable("ProjectParams")
                dt.Columns.Add("RvtName", GetType(String))
                dt.Columns.Add("RvtPath", GetType(String))
                dt.Columns.Add("ParamName", GetType(String))
                dt.Columns.Add("ParamKind", GetType(String))
                dt.Columns.Add("ParamGroup", GetType(String))
                dt.Columns.Add("BoundCategories", GetType(String))
                dt.Columns.Add("RvtGuid", GetType(String))
                dt.Columns.Add("FileGuid", GetType(String))
                dt.Columns.Add("Result", GetType(String))
                dt.Columns.Add("Notes", GetType(String))

                Dim allowedCategoryNames As HashSet(Of String) = BuildAllowedCategoryNameSet(doc)

                Dim speByName As New Dictionary(Of String, List(Of Guid))(StringComparer.OrdinalIgnoreCase)
                Try
                    For Each spe As SharedParameterElement In New FilteredElementCollector(doc).OfClass(GetType(SharedParameterElement)).Cast(Of SharedParameterElement)()
                        Dim key As String = NormalizeName(SafeParamElementName(spe))
                        Dim g As Guid = Guid.Empty
                        Try
                            g = spe.GuidValue
                        Catch
                            g = Guid.Empty
                        End Try
                        If g = Guid.Empty Then Continue For
                        If Not speByName.ContainsKey(key) Then speByName(key) = New List(Of Guid)()
                        speByName(key).Add(g)
                    Next
                Catch
                End Try

                Dim bindings As BindingMap = doc.ParameterBindings
                Dim iter As DefinitionBindingMapIterator = bindings.ForwardIterator()
                iter.Reset()

                Dim idx As Integer = 0
                Dim total As Integer = 0
                Try
                    While iter.MoveNext()
                        total += 1
                    End While
                Catch
                    total = 0
                End Try

                Try
                    iter.Reset()
                Catch
                End Try

                While True
                    Dim moved As Boolean = False
                    Try
                        moved = iter.MoveNext()
                    Catch
                        Exit While
                    End Try
                    If Not moved Then Exit While

                    idx += 1
                    If progress IsNot Nothing Then progress(idx, Math.Max(1, total))

                    Dim def As Definition = Nothing
                    Dim binding As ElementBinding = Nothing
                    Try
                        def = iter.Key
                        binding = TryCast(iter.Current, ElementBinding)
                    Catch
                        def = Nothing
                        binding = Nothing
                    End Try

                    If def Is Nothing Then Continue While

                    Dim name As String = ""
                    Try : name = def.Name : Catch : name = "" : End Try
                    Dim normName As String = NormalizeName(name)

                    Dim kind As String = "Project"
                    Dim projGuid As String = ""
                    Dim fileGuid As String = ""
                    Dim result As String = ""
                    Dim notes As String = ""

                    Dim isShared As Boolean = TypeOf def Is ExternalDefinition
                    Dim docGuid As Guid = Guid.Empty
                    Dim docGuids As List(Of Guid) = Nothing
                    If isShared Then
                        kind = "Shared"
                        Try
                            docGuid = DirectCast(def, ExternalDefinition).GUID
                        Catch
                            docGuid = Guid.Empty
                        End Try
                        If docGuid <> Guid.Empty Then docGuids = New List(Of Guid)() From {docGuid}
                    Else
                        Dim list As List(Of Guid) = Nothing
                        If speByName.TryGetValue(normName, list) Then
                            isShared = True
                            kind = "Shared"
                            docGuids = New List(Of Guid)(list)
                            docGuid = docGuids.FirstOrDefault()
                        End If
                    End If

                    If isShared Then
                        projGuid = If(docGuid = Guid.Empty, "", docGuid.ToString())
                        Dim fileGuids As List(Of Guid) = Nothing
                        If fileMap IsNot Nothing AndAlso fileMap.TryGetValue(normName, fileGuids) Then
                            fileGuid = String.Join("; ", fileGuids.Select(Function(x) x.ToString()).Distinct().ToArray())
                            If docGuids Is Nothing Then docGuids = New List(Of Guid)()
                            Dim hit As Boolean = False
                            For Each g In fileGuids
                                If docGuids.Any(Function(x) x = g) Then
                                    hit = True
                                    Exit For
                                End If
                            Next
                            result = If(hit, If(fileGuids.Count > 1, "OK(MULTI_IN_FILE)", "OK"), "MISMATCH")
                            If result = "MISMATCH" Then
                                notes = "RVT의 GUID와 Shared Parameter 파일 GUID 불일치"
                            End If
                        Else
                            result = "NOT_FOUND_IN_FILE"
                            notes = "Shared Parameter 파일에서 동일 이름을 찾지 못함"
                        End If

                        If result = "OK" OrElse result = "OK(MULTI_IN_FILE)" OrElse result = "MISMATCH" Then
                            If fileGuids IsNot Nothing AndAlso fileGuids.Count > 1 Then
                                notes = AppendNote(notes, "파일 내 동일 이름 GUID 여러 개")
                            End If
                            If docGuids IsNot Nothing AndAlso docGuids.Count > 1 Then
                                notes = AppendNote(notes, "문서 내 동일 이름 GUID 여러 개")
                            End If
                        End If
                    Else
                        result = "PROJECT_PARAM"
                    End If

                    Dim r = dt.NewRow()
                    r("RvtName") = If(rvtName, "")
                    r("RvtPath") = If(rvtPath, "")
                    r("ParamName") = name
                    r("ParamKind") = kind
                    r("ParamGroup") = SafeParameterGroupName(def)
                    r("BoundCategories") = FormatBoundCategories(binding, allowedCategoryNames)
                    r("RvtGuid") = projGuid
                    r("FileGuid") = fileGuid
                    r("Result") = result
                    r("Notes") = notes
                    dt.Rows.Add(r)
                End While

                If dt.Columns.Contains("BoundCategories") Then
                    dt.Columns("BoundCategories").SetOrdinal(dt.Columns.Count - 1)
                End If

                Return dt
            End Function

            Private Shared Function AppendNote(existing As String, note As String) As String
                If String.IsNullOrWhiteSpace(existing) Then Return note
                If String.IsNullOrWhiteSpace(note) Then Return existing
                Return existing & "; " & note
            End Function

            Private Shared Function SafeParameterGroupName(def As Definition) As String
                If def Is Nothing Then Return ""

                ' Revit 2019~2023: InternalDefinition.ParameterGroup + LabelUtils.GetLabelFor(BuiltInParameterGroup)
                Try
                    Dim idef As InternalDefinition = TryCast(def, InternalDefinition)
                    If idef IsNot Nothing Then
                        Dim pi As PropertyInfo = idef.GetType().GetProperty("ParameterGroup", BindingFlags.Public Or BindingFlags.Instance)
                        If pi IsNot Nothing Then
                            Dim pgObj As Object = pi.GetValue(idef, Nothing)
                            If pgObj IsNot Nothing Then
                                Dim miLabelFor As MethodInfo =
                                    GetType(LabelUtils).GetMethod("GetLabelFor",
                                                                  BindingFlags.Public Or BindingFlags.Static,
                                                                  Nothing,
                                                                  New Type() {pgObj.GetType()},
                                                                  Nothing)
                                If miLabelFor IsNot Nothing Then
                                    Dim labelObj As Object = miLabelFor.Invoke(Nothing, New Object() {pgObj})
                                    Dim label As String = TryCast(labelObj, String)
                                    If Not String.IsNullOrWhiteSpace(label) Then Return label
                                End If

                                Dim fallback As String = pgObj.ToString()
                                If Not String.IsNullOrWhiteSpace(fallback) Then Return fallback
                            End If
                        End If
                    End If
                Catch
                End Try

                ' Revit 2025: Definition.GetGroupTypeId + LabelUtils.GetLabelForGroup(ForgeTypeId)
                Try
                    Dim miGetGroupTypeId As MethodInfo =
                        def.GetType().GetMethod("GetGroupTypeId", BindingFlags.Public Or BindingFlags.Instance)

                    If miGetGroupTypeId IsNot Nothing Then
                        Dim groupIdObj As Object = miGetGroupTypeId.Invoke(def, Nothing)
                        If groupIdObj IsNot Nothing Then
                            Dim miLabelForGroup As MethodInfo =
                                GetType(LabelUtils).GetMethod("GetLabelForGroup",
                                                              BindingFlags.Public Or BindingFlags.Static,
                                                              Nothing,
                                                              New Type() {groupIdObj.GetType()},
                                                              Nothing)

                            If miLabelForGroup IsNot Nothing Then
                                Dim labelObj As Object = miLabelForGroup.Invoke(Nothing, New Object() {groupIdObj})
                                Dim label As String = TryCast(labelObj, String)
                                If Not String.IsNullOrWhiteSpace(label) Then Return label
                            End If

                            Dim fallback As String = groupIdObj.ToString()
                            If Not String.IsNullOrWhiteSpace(fallback) Then Return fallback
                        End If
                    End If
                Catch
                End Try

                Return ""
            End Function

            Private Shared Function FormatBoundCategories(binding As ElementBinding,
                                                          allowedCategoryNames As HashSet(Of String)) As String
                If binding Is Nothing OrElse binding.Categories Is Nothing Then Return ""
                If allowedCategoryNames Is Nothing OrElse allowedCategoryNames.Count = 0 Then Return ""

                Dim topLevelNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim subByTop As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

                For Each cat As Category In binding.Categories
                    If cat Is Nothing Then Continue For

                    Dim currentName As String = SafeCategoryName(cat)
                    If String.IsNullOrWhiteSpace(currentName) Then Continue For

                    Dim parent As Category = Nothing
                    Try
                        parent = cat.Parent
                    Catch
                        parent = Nothing
                    End Try

                    If parent Is Nothing Then
                        If allowedCategoryNames.Contains(currentName) Then
                            topLevelNames.Add(currentName)
                        End If
                        Continue For
                    End If

                    Dim parentName As String = SafeCategoryName(parent)
                    If String.IsNullOrWhiteSpace(parentName) Then Continue For
                    If Not allowedCategoryNames.Contains(parentName) Then Continue For

                    topLevelNames.Add(parentName)
                    If Not subByTop.ContainsKey(parentName) Then
                        subByTop(parentName) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    End If
                    subByTop(parentName).Add(currentName)
                Next

                Dim labels As New List(Of String)()
                For Each top In topLevelNames.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    labels.Add($"[{top}]")

                    Dim subs As HashSet(Of String) = Nothing
                    If subByTop.TryGetValue(top, subs) AndAlso subs IsNot Nothing Then
                        For Each subName In subs.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                            labels.Add($"[{top}: {subName}]")
                        Next
                    End If
                Next

                Return String.Join(",", labels.ToArray())
            End Function

            Private Shared Function SafeCategoryName(cat As Category) As String
                If cat Is Nothing Then Return ""
                Try
                    Dim name As String = If(cat.Name, "")
                    If String.IsNullOrWhiteSpace(name) Then Return ""
                    Return name.Trim()
                Catch
                    Return ""
                End Try
            End Function

            Private Shared Function BuildAllowedCategoryNameSet(doc As Document) As HashSet(Of String)
                Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                If doc Is Nothing Then Return result

                Dim cats As Categories = Nothing
                Try
                    cats = doc.Settings.Categories
                Catch
                    cats = Nothing
                End Try
                If cats Is Nothing Then Return result

                For Each cat As Category In cats
                    AddAllowedCategoryName(cat, result)

                    Try
                        Dim subs As CategoryNameMap = cat.SubCategories
                        If subs IsNot Nothing Then
                            For Each subCat As Category In subs
                                AddAllowedCategoryName(subCat, result)
                            Next
                        End If
                    Catch
                    End Try
                Next

                Return result
            End Function

            Private Shared Sub AddAllowedCategoryName(cat As Category, allowedNames As HashSet(Of String))
                If cat Is Nothing OrElse allowedNames Is Nothing Then Return

                Dim name As String = ""
                Try
                    name = If(cat.Name, "")
                Catch
                    name = ""
                End Try
                If String.IsNullOrWhiteSpace(name) Then Return

                Dim trimmed As String = name.Trim()
                If trimmed.StartsWith("<", StringComparison.OrdinalIgnoreCase) Then Return
                If trimmed.IndexOf("line style", StringComparison.OrdinalIgnoreCase) >= 0 Then Return

                Dim canBind As Boolean = False
                Try
                    canBind = cat.AllowsBoundParameters
                Catch
                    canBind = False
                End Try
                If Not canBind Then Return

                allowedNames.Add(trimmed)
            End Sub

            Public Shared Function RunFamilyAudit(doc As Document,
                                                  fileMap As Dictionary(Of String, List(Of Guid)),
                                                  rvtName As String,
                                                  rvtPath As String,
                                                  Optional progress As Action(Of Integer, Integer, String) = Nothing,
                                                  Optional includeAnnotation As Boolean = False) As FamilyAuditPack

                Dim pack As New FamilyAuditPack()

                Dim dtDet As New DataTable("FamilyParamDetail")
                dtDet.Columns.Add("RvtName", GetType(String))
                dtDet.Columns.Add("RvtPath", GetType(String))
                dtDet.Columns.Add("FamilyName", GetType(String))
                dtDet.Columns.Add("FamilyCategory", GetType(String))
                dtDet.Columns.Add("ParamName", GetType(String))
                dtDet.Columns.Add("ParamKind", GetType(String))
                dtDet.Columns.Add("IsShared", GetType(String))
                dtDet.Columns.Add("FamilyGuid", GetType(String))
                dtDet.Columns.Add("FileGuid", GetType(String))
                dtDet.Columns.Add("Result", GetType(String))
                dtDet.Columns.Add("Notes", GetType(String))

                Dim dtIdx As New DataTable("FamilyIndex")
                dtIdx.Columns.Add("RvtName", GetType(String))
                dtIdx.Columns.Add("RvtPath", GetType(String))
                dtIdx.Columns.Add("FamilyName", GetType(String))
                dtIdx.Columns.Add("FamilyCategory", GetType(String))
                dtIdx.Columns.Add("TotalParamCount", GetType(Integer))
                dtIdx.Columns.Add("SharedParamCount", GetType(Integer))

                Dim fams = New FilteredElementCollector(doc).
                    OfClass(GetType(Family)).
                    Cast(Of Family)().
                    OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase).
                    ToList()

                Dim total As Integer = Math.Max(1, fams.Count)
                Dim idx As Integer = 0

                For Each fam As Family In fams
                    idx += 1

                    If progress IsNot Nothing Then progress(idx, total, fam.Name)

                    Dim famName = fam.Name
                    Dim famCat = ""
                    Try
                        If fam.FamilyCategory IsNot Nothing Then famCat = fam.FamilyCategory.Name
                    Catch
                        famCat = ""
                    End Try

                    Try
                        If fam.FamilyCategory IsNot Nothing Then
                            Dim catType As CategoryType
                            Try
                                catType = fam.FamilyCategory.CategoryType
                            Catch
                                catType = CType(-1, CategoryType)
                            End Try
                            If catType = CategoryType.Annotation AndAlso Not includeAnnotation Then
                                Continue For
                            End If
                        End If
                    Catch
                    End Try

                    Dim skip As Boolean = False
                    Try
                        If fam.IsInPlace Then skip = True
                    Catch
                        skip = False
                    End Try
                    If skip Then Continue For

                    Try
                        Dim p = fam.GetType().GetProperty("IsEditable", BindingFlags.Public Or BindingFlags.Instance)
                        If p IsNot Nothing Then
                            Dim v = p.GetValue(fam, Nothing)
                            If TypeOf v Is Boolean AndAlso Not DirectCast(v, Boolean) Then
                                Continue For
                            End If
                        End If
                    Catch
                    End Try

                    Dim famDoc As Document = Nothing
                    Try
                        Dim isInPlace As Boolean = False
                        Try
                            isInPlace = fam.IsInPlace
                        Catch
                            isInPlace = False
                        End Try
                        If isInPlace Then Continue For

                        Try
                            famDoc = doc.EditFamily(fam)
                        Catch ex As InvalidOperationException
                            famDoc = Nothing
                            Continue For
                        End Try

                        If famDoc Is Nothing OrElse Not famDoc.IsFamilyDocument Then
                            Continue For
                        End If

                        Dim fm As FamilyManager = famDoc.FamilyManager
                        If fm Is Nothing Then
                            AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "OPEN_FAIL", "FamilyManager 없음")
                            Continue For
                        End If

                        Dim totalParamCount As Integer = 0
                        Dim sharedCount As Integer = 0

                        For Each fp As FamilyParameter In fm.Parameters
                            If fp Is Nothing Then Continue For
                            totalParamCount += 1

                            Dim pName As String = ""
                            Try : pName = fp.Definition.Name : Catch : pName = "" : End Try
                            Dim normParamName As String = NormalizeName(pName)

                            Dim paramKind As String = GetFamilyParamKind(fp)
                            Dim isSharedBool As Boolean = String.Equals(paramKind, "Shared", StringComparison.OrdinalIgnoreCase)
                            If isSharedBool Then sharedCount += 1

                            Dim famGuid As String = ""
                            Dim fileGuid As String = ""
                            Dim res As String = ""
                            Dim notes As String = ""

                            If isSharedBool Then
                                Dim gFam As Guid = Guid.Empty
                                If TryGetFamilyParameterGuid(fp, gFam) Then
                                    famGuid = gFam.ToString()

                                    Dim fileGuids As List(Of Guid) = Nothing
                                    If fileMap.TryGetValue(normParamName, fileGuids) Then
                                        fileGuid = String.Join("; ", fileGuids.Select(Function(x) x.ToString()).Distinct().ToArray())

                                        If fileGuids.Any(Function(x) x = gFam) Then
                                            res = If(fileGuids.Count > 1, "OK(MULTI_IN_FILE)", "OK")
                                        Else
                                            res = "MISMATCH"
                                        End If
                                    Else
                                        res = "NOT_FOUND_IN_FILE"
                                    End If
                                Else
                                    res = "GUID_FAIL"
                                    notes = "FamilyParameter GUID 추출 실패"
                                End If
                            ElseIf String.Equals(paramKind, "BuiltIn", StringComparison.OrdinalIgnoreCase) Then
                                res = "BUILTIN"
                            ElseIf String.Equals(paramKind, "Family", StringComparison.OrdinalIgnoreCase) Then
                                res = "FAMILY_PARAM"
                            Else
                                res = "FAMILY_PARAM"
                            End If

                            If isSharedBool Then
                                If res = "NOT_FOUND_IN_FILE" Then
                                    notes = "Shared Parameter 파일에서 동일 이름을 찾지 못함"
                                ElseIf res = "MISMATCH" Then
                                    notes = "RVT의 GUID와 Shared Parameter 파일 GUID 불일치"
                                End If

                                If res = "OK" OrElse res = "OK(MULTI_IN_FILE)" OrElse res = "MISMATCH" Then
                                    Dim fileGuids As List(Of Guid) = Nothing
                                    If fileMap.TryGetValue(normParamName, fileGuids) AndAlso fileGuids IsNot Nothing AndAlso fileGuids.Count > 1 Then
                                        notes = AppendNote(notes, "파일 내 동일 이름 GUID 여러 개")
                                    End If
                                End If
                            End If

                            AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, pName,
                                         paramKind,
                                         If(isSharedBool, "Y", "N"),
                                         famGuid, fileGuid, res, notes)
                        Next

                        Dim rIdx = dtIdx.NewRow()
                        rIdx("RvtName") = If(rvtName, "")
                        rIdx("RvtPath") = If(rvtPath, "")
                        rIdx("FamilyName") = If(famName, "")
                        rIdx("FamilyCategory") = If(famCat, "")
                        rIdx("TotalParamCount") = totalParamCount
                        rIdx("SharedParamCount") = sharedCount
                        dtIdx.Rows.Add(rIdx)

                    Catch ex As Exception
                        AddDetailRow(dtDet, rvtName, rvtPath, famName, famCat, "", "N/A", "", "", "", "OPEN_FAIL", ex.Message)

                    Finally
                        If famDoc IsNot Nothing Then
                            Try
                                famDoc.Close(False)
                            Catch
                            End Try
                        End If
                    End Try
                Next

                pack.Summary = Nothing
                pack.Detail = dtDet
                pack.Index = dtIdx
                Return pack
            End Function

            Private Shared Sub AddDetailRow(dt As DataTable,
                                            rvtName As String,
                                            rvtPath As String,
                                            famName As String,
                                            famCat As String,
                                            pName As String,
                                            paramKind As String,
                                            isShared As String,
                                            famGuid As String,
                                            fileGuid As String,
                                            res As String,
                                            notes As String)
                Dim r = dt.NewRow()
                r("RvtName") = If(rvtName, "")
                r("RvtPath") = If(rvtPath, "")
                r("FamilyName") = If(famName, "")
                r("FamilyCategory") = If(famCat, "")
                r("ParamName") = If(pName, "")
                r("ParamKind") = If(paramKind, "")
                r("IsShared") = If(isShared, "")
                r("FamilyGuid") = If(famGuid, "")
                r("FileGuid") = If(fileGuid, "")
                r("Result") = If(res, "")
                r("Notes") = If(notes, "")
                dt.Rows.Add(r)
            End Sub

            Private Shared Function GetFamilyParamKind(fp As FamilyParameter) As String
                If fp Is Nothing Then Return "None"
                Dim isSharedFlag As Boolean = False
                Try : isSharedFlag = fp.IsShared : Catch : isSharedFlag = False : End Try
                If isSharedFlag Then Return "Shared"
                Dim idVal As Integer = 0
                Try : idVal = fp.Id.IntegerValue : Catch : idVal = 0 : End Try
                If idVal < 0 Then Return "BuiltIn"
                Return "Family"
            End Function

            Private Shared Function SafeParamElementName(pe As Element) As String
                Try
                    Return pe.Name
                Catch
                    Return ""
                End Try
            End Function


            Private Shared Function TryGetFamilyParameterGuid(fp As FamilyParameter, ByRef g As Guid) As Boolean
                g = Guid.Empty
                If fp Is Nothing Then Return False

                Dim t = fp.GetType()
                Dim p = t.GetProperty("GUID", BindingFlags.Public Or BindingFlags.Instance)
                If p Is Nothing Then Return False

                Dim v = p.GetValue(fp, Nothing)
                If v Is Nothing Then Return False

                If TypeOf v Is Guid Then
                    g = DirectCast(v, Guid)
                    Return g <> Guid.Empty
                End If

                Return False
            End Function

            Private Shared Function GetParamTypeName(def As Definition) As String
                If def Is Nothing Then
                    Return String.Empty
                End If

                Try
                    Dim p = def.GetType().GetProperty("ParameterType", BindingFlags.Public Or BindingFlags.Instance)
                    If p IsNot Nothing Then
                        Dim v = p.GetValue(def, Nothing)
                        If v IsNot Nothing Then Return v.ToString()
                    End If
                Catch
                End Try

                Try
                    Dim m = def.GetType().GetMethod("GetDataType", BindingFlags.Public Or BindingFlags.Instance)
                    If m IsNot Nothing Then
                        Dim v = m.Invoke(def, Nothing)
                        If v IsNot Nothing Then Return v.ToString()
                    End If
                Catch
                End Try

                Try
                    Dim p2 = def.GetType().GetProperty("DataType", BindingFlags.Public Or BindingFlags.Instance)
                    If p2 IsNot Nothing Then
                        Dim v = p2.GetValue(def, Nothing)
                        If v IsNot Nothing Then Return v.ToString()
                    End If
                Catch
                End Try

                Return String.Empty
            End Function

        End Class

    End Class

End Namespace
