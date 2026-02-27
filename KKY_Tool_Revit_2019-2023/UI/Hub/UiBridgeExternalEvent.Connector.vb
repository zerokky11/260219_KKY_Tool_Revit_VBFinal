Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Diagnostics
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports System.Windows.Forms ' WinForms 다이얼로그 사용
Imports KKY_Tool_Revit.Infrastructure

Namespace UI.Hub
    ' 커넥터 진단 (fix2 이벤트명/스키마 유지)
    Partial Public Class UiBridgeExternalEvent

        ' 최근 로드/실행 결과(엑셀 내보내기 시 기본 소스)
        Private lastConnRows As List(Of Dictionary(Of String, Object)) = Nothing

        ' 전체 커넥터 결과(엑셀 내보내기용) - Total/Detail 분리
        Private _connectorTotalRows As List(Of Dictionary(Of String, Object)) = Nothing
        Private _connectorDetailRows As List(Of Dictionary(Of String, Object)) = Nothing

        ' 추가 추출 파라미터
        Private _connectorExtraParams As List(Of String) = Nothing

        ' 추가 필터
        Private _connectorTargetFilter As String = String.Empty
        Private _connectorExcludeEndDummy As Boolean = False

        ' 마지막 실행 시 UI 단위(엑셀 헤더/거리 변환에 사용)
        ' - SaveExcel payload 에 unit 이 누락되는 케이스가 있어도, 직전 실행 설정으로 내보내기 일관성 유지
        Private _connectorUiUnit As String = "inch"
        Private _lastConnectorTol As Double = 1.0R
        Private _lastConnectorUnit As String = "inch"
        Private _lastConnectorParam As String = "Comments"
        Private _lastConnectorReviewParams As List(Of String) = New List(Of String) From {"Comments"}
        Private _lastConnectorTargetFilter As String = String.Empty
        Private _lastConnectorExcludeEndDummy As Boolean = False

        ' 디버그 로그를 웹(F12 콘솔)로 보내는 헬퍼
        Private Sub LogDebug(message As String)
            Try
                Dim ts As String = Date.Now.ToString("HH:mm:ss")
                SendToWeb("host:log", New With {
                    .message = $"[{ts}] {message}"
                })
            Catch
                ' 로깅 중 예외는 무시
            End Try
        End Sub

        ' 오류 로그용
        Private Sub LogError(message As String)
            Try
                Dim ts As String = Date.Now.ToString("HH:mm:ss")
                SendToWeb("host:error", New With {
                    .message = $"[{ts}] {message}"
                })
            Catch
            End Try
        End Sub

        Private Sub ReportConnectorProgress(pct As Double, text As String)
            SendToWeb("connector:progress", New With {.pct = pct, .text = text})
        End Sub

        Private Function SafePayloadSnapshot(payload As Object) As String
            If payload Is Nothing Then Return "(null)"
            Try
                Dim dict = TryCast(payload, IDictionary(Of String, Object))
                If dict IsNot Nothing Then
                    Dim parts As New List(Of String)()
                    For Each kv In dict
                        Dim v As Object = kv.Value
                        Dim text As String = If(v Is Nothing, "(null)", v.ToString())
                        parts.Add(kv.Key & "=" & text)
                    Next
                    Return "{" & String.Join(", ", parts) & "}"
                End If
                Return payload.ToString()
            Catch
                Return "(payload)"
            End Try
        End Function

#Region "핸들러 (Core에서 리플렉션으로 호출)"

        ' === connector:run ===
        Private Sub HandleConnectorRun(app As UIApplication, payload As Object)
            Try
                LogDebug("[connector] HandleConnectorRun 진입")
                LogDebug("[connector] payload 수신: " & SafePayloadSnapshot(payload))
                ReportConnectorProgress(0.1R, "커넥터 진단 시작...")

                Dim uidoc = app.ActiveUIDocument
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    LogError("[connector] 활성 문서가 없습니다.")
                    SendToWeb("revit:error", New With {.message = "활성 문서가 없습니다."})
                    SendToWeb("connector:done", New With {.ok = False, .message = "활성 문서가 없습니다."})
                    Return
                End If

                _connectorTotalRows = Nothing
                _connectorDetailRows = Nothing

                ' === payload 파싱 ===
                Dim tol As Double = 1.0 ' 기본 1 inch
                Dim unit As String = "inch"
                Dim param As String = "Comments"
                Dim paramsCsv As String = String.Empty
                Dim paramsArray As List(Of String) = Nothing
                Try
                    Dim vTol = GetProp(payload, "tol")
                    If vTol IsNot Nothing Then tol = Convert.ToDouble(vTol)
                Catch : End Try
                Try
                    Dim vUnit = TryCast(GetProp(payload, "unit"), String)
                    If Not String.IsNullOrEmpty(vUnit) Then unit = vUnit
                Catch : End Try
                Try
                    Dim vParam = TryCast(GetProp(payload, "param"), String)
                    If Not String.IsNullOrEmpty(vParam) Then param = vParam
                Catch : End Try
                Try
                    paramsCsv = TryCast(GetProp(payload, "paramsCsv"), String)
                Catch
                    paramsCsv = String.Empty
                End Try
                Try
                    paramsArray = ParseParamsArray(GetProp(payload, "params"))
                Catch
                    paramsArray = Nothing
                End Try

                Dim paramCsvNormalized As String = NormalizeParamsCsv(paramsCsv)
                If String.IsNullOrWhiteSpace(paramCsvNormalized) AndAlso paramsArray IsNot Nothing AndAlso paramsArray.Count > 0 Then
                    paramCsvNormalized = NormalizeParamsCsv(String.Join(",", paramsArray))
                End If
                If String.IsNullOrWhiteSpace(paramCsvNormalized) Then
                    paramCsvNormalized = NormalizeParamsCsv(param)
                End If
                If String.IsNullOrWhiteSpace(paramCsvNormalized) Then
                    paramCsvNormalized = "Comments"
                End If

                Dim reviewParams As List(Of String) = ParseReviewParamsCsv(paramCsvNormalized)
                If reviewParams Is Nothing OrElse reviewParams.Count = 0 Then
                    reviewParams = New List(Of String) From {"Comments"}
                End If
                paramCsvNormalized = String.Join(",", reviewParams)

                _connectorExtraParams = ParseExtraParams(TryCast(GetProp(payload, "extraParams"), String))
                Try
                    Dim vFilter = TryCast(GetProp(payload, "targetFilter"), String)
                    _connectorTargetFilter = If(vFilter, String.Empty)
                Catch
                    _connectorTargetFilter = String.Empty
                End Try
                Try
                    Dim vExclude = GetProp(payload, "excludeEndDummy")
                    If vExclude IsNot Nothing Then
                        _connectorExcludeEndDummy = Convert.ToBoolean(vExclude)
                    Else
                        _connectorExcludeEndDummy = False
                    End If
                Catch
                    _connectorExcludeEndDummy = False
                End Try
                LogDebug($"[connector] 파라미터 파싱 완료 (tol={tol}, unit={unit}, param={param}, paramsCsv={paramCsvNormalized}, extra={String.Join(",", _connectorExtraParams)} )")

                ' 직전 실행 단위 저장(엑셀 내보내기에서 기본값으로 사용)
                _connectorUiUnit = NormalizeUiUnit(unit)
                _lastConnectorTol = tol
                _lastConnectorUnit = unit
                _lastConnectorParam = param
                _lastConnectorReviewParams = reviewParams
                _lastConnectorTargetFilter = _connectorTargetFilter
                _lastConnectorExcludeEndDummy = _connectorExcludeEndDummy

                ' === 서비스 호출 ===
                LogDebug("[connector] 커넥터 수집/진단 실행 시작")
                Const PREVIEW_LIMIT As Integer = 150
                Dim rows As List(Of Dictionary(Of String, Object)) = Nothing
                Try
                    rows = Services.ConnectorDiagnosticsService.Run(app, tol, unit, paramCsvNormalized, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy, AddressOf ReportConnectorProgress)
                Catch ex As Exception
                    ' 네임스페이스 변동 대비 리플렉션 재시도
                    Try
                        Dim t = Type.GetType("KKY_Tool_Revit.Services.ConnectorDiagnosticsService, KKY_Tool_Revit")
                        If t Is Nothing Then t = Type.GetType("ConnectorDiagnosticsService")
                        If t IsNot Nothing Then
                            Dim m = t.GetMethod("Run", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static)
                            If m IsNot Nothing Then
                                Dim args As Object()
                                Dim ps = m.GetParameters()
                                If ps.Length >= 8 Then
                                    args = New Object() {app, tol, unit, paramCsvNormalized, _connectorExtraParams, _connectorTargetFilter, _connectorExcludeEndDummy, CType(AddressOf ReportConnectorProgress, Action(Of Double, String))}
                                ElseIf ps.Length >= 6 Then
                                    args = New Object() {app, tol, unit, paramCsvNormalized, _connectorExtraParams, CType(AddressOf ReportConnectorProgress, Action(Of Double, String))}
                                Else
                                    args = New Object() {app, tol, unit, paramCsvNormalized}
                                End If
                                rows = CType(m.Invoke(Nothing, args), List(Of Dictionary(Of String, Object)))
                            End If
                        End If
                    Catch
                    End Try
                End Try
                If rows Is Nothing Then rows = New List(Of Dictionary(Of String, Object))()

                Try
                    Dim svcLog = Services.ConnectorDiagnosticsService.LastDebug
                    If svcLog IsNot Nothing Then
                        For Each line In svcLog
                            LogDebug("[connector][svc] " & line)
                        Next
                    End If
                Catch
                End Try

                Dim totalRows = If(rows, New List(Of Dictionary(Of String, Object))()).Select(Function(r) CloneRow(r)).ToList()
                Dim filteredRows = totalRows.Where(Function(r) ShouldIncludeRow(r)).ToList()

                Dim mismatchAll = filteredRows.Where(Function(r) IsMismatchRow(r)).ToList()
                Dim nearAll = filteredRows.Where(Function(r) IsNearConnection(r)).ToList()


                ' ✅ 멀티 파라미터 중 "이슈 0건" 파라미터도 검토 여부를 알 수 있도록 안내행 추가
                Dim uiMsgRows = BuildNoIssueMessageRows(filteredRows, reviewParams)
                If uiMsgRows IsNot Nothing AndAlso uiMsgRows.Count > 0 Then
                    filteredRows = uiMsgRows.Concat(filteredRows).ToList()
                End If

                Dim mismatchPreview As List(Of Dictionary(Of String, Object)) = mismatchAll.Take(PREVIEW_LIMIT).ToList()
                Dim nearPreview As List(Of Dictionary(Of String, Object)) = nearAll.Take(PREVIEW_LIMIT).ToList()
                Dim previewRows As List(Of Dictionary(Of String, Object)) = filteredRows.Take(PREVIEW_LIMIT).ToList()

                _connectorTotalRows = filteredRows
                _connectorDetailRows = rows
                lastConnRows = filteredRows

                Dim mismatchCount As Integer = mismatchAll.Count
                Dim okCount As Integer = Math.Max(filteredRows.Count - mismatchCount, 0)
                LogDebug($"[connector] 규칙/비교 로직 적용 완료: 정상 {okCount}개, 경고/오류 {mismatchCount}개")
                LogDebug($"[connector] 커넥터 수집 완료: 결과 행 {filteredRows.Count}개 (Mismatch={mismatchAll.Count}, Near={nearAll.Count})")

                LogDebug("[connector] 결과 전송 준비 완료, connector:done/connector:loaded emit 직전")
                Dim hasMore As Boolean = filteredRows.Count > PREVIEW_LIMIT
                SendToWeb("connector:loaded", New With {
                    .rows = previewRows,
                    .total = filteredRows.Count,
                    .previewCount = previewRows.Count,
                    .hasMore = hasMore,
                    .mismatch = New With {
                        .rows = mismatchPreview,
                        .total = mismatchAll.Count,
                        .previewCount = mismatchPreview.Count,
                        .hasMore = mismatchAll.Count > PREVIEW_LIMIT
                    },
                    .near = New With {
                        .rows = nearPreview,
                        .total = nearAll.Count,
                        .previewCount = nearPreview.Count,
                        .hasMore = nearAll.Count > PREVIEW_LIMIT
                    },
                    .extraParams = _connectorExtraParams,
                    .reviewParams = reviewParams,
                    .paramsCsv = paramCsvNormalized
                })
                SendToWeb("connector:done", New With {
                    .rows = previewRows,
                    .total = filteredRows.Count,
                    .previewCount = previewRows.Count,
                    .hasMore = hasMore,
                    .mismatch = New With {
                        .rows = mismatchPreview,
                        .total = mismatchAll.Count,
                        .previewCount = mismatchPreview.Count,
                        .hasMore = mismatchAll.Count > PREVIEW_LIMIT
                    },
                    .near = New With {
                        .rows = nearPreview,
                        .total = nearAll.Count,
                        .previewCount = nearPreview.Count,
                        .hasMore = nearAll.Count > PREVIEW_LIMIT
                    },
                    .extraParams = _connectorExtraParams,
                    .reviewParams = reviewParams,
                    .paramsCsv = paramCsvNormalized
                })
                LogDebug("[connector] 결과 전송 완료, connector:done emit")
                LogDebug("[connector] HandleConnectorRun 정상 종료")

            Catch ex As Exception
                LogError("[connector] 검사 중 예외 발생: " & ex.ToString())
                SendToWeb("connector:done", New With {.ok = False, .message = ex.Message})
                SendToWeb("revit:error", New With {.message = "실행 실패: " & ex.Message})
            Finally
                ReportConnectorProgress(0R, String.Empty)
            End Try
        End Sub

        ' === connector:param-list ===
        Private Sub HandleConnectorParamList(app As UIApplication, payload As Object)
            Try
                Dim uidoc = app.ActiveUIDocument
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    SendToWeb("connector:param-list:done", New With {.ok = False, .message = "활성 문서가 없습니다.", .params = New List(Of String)()})
                    Return
                End If

                Dim names As New List(Of String)()
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                Try
                    Dim bindings = doc.ParameterBindings
                    If bindings IsNot Nothing Then
                        Dim it = bindings.ForwardIterator()
                        If it IsNot Nothing Then
                            it.Reset()
                            While it.MoveNext()
                                Dim def = TryCast(it.Key, Definition)
                                If def Is Nothing Then Continue While
                                Dim name As String = If(def.Name, String.Empty).Trim()
                                If name = "" Then Continue While
                                If seen.Add(name) Then names.Add(name)
                            End While
                        End If
                    End If
                Catch
                End Try

                names.Sort(StringComparer.OrdinalIgnoreCase)
                SendToWeb("connector:param-list:done", New With {.ok = True, .params = names})
            Catch ex As Exception
                SendToWeb("connector:param-list:done", New With {.ok = False, .message = ex.Message, .params = New List(Of String)()})
            End Try
        End Sub

        ' === connector:save-excel ===
        Private Sub HandleConnectorSaveExcel(app As UIApplication, payload As Object)
            Try
                Dim rows As List(Of Dictionary(Of String, Object)) = _connectorTotalRows
                If rows Is Nothing OrElse rows.Count = 0 Then rows = TryGetRowsFromPayload(payload)
                If rows Is Nothing OrElse rows.Count = 0 Then rows = lastConnRows

                If rows Is Nothing Then
                    SendToWeb("revit:error", New With {.message = "저장할 데이터가 없습니다."})
                    Return
                End If

                Dim filteredTotal = rows.Where(AddressOf ShouldExportToExcel).ToList()

                If filteredTotal Is Nothing OrElse filteredTotal.Count = 0 Then
                    System.Windows.Forms.MessageBox.Show("이슈 항목이 없습니다.", "검토 결과", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                ' UI 단위는 connector:save-excel payload 에 포함되는게 이상적이지만,
                ' 누락되는 케이스가 있어 직전 실행값(_connectorUiUnit)을 기본으로 사용.
                Dim uiUnitRaw As String = ""
                Try
                    uiUnitRaw = TryCast(GetProp(payload, "uiUnit"), String)
                Catch
                    uiUnitRaw = ""
                End Try
                If String.IsNullOrWhiteSpace(uiUnitRaw) Then
                    Try
                        uiUnitRaw = TryCast(GetProp(payload, "displayUnit"), String)
                    Catch
                        uiUnitRaw = ""
                    End Try
                End If
                If String.IsNullOrWhiteSpace(uiUnitRaw) Then
                    Try
                        uiUnitRaw = TryCast(GetProp(payload, "unit"), String)
                    Catch
                        uiUnitRaw = ""
                    End Try
                End If
                If String.IsNullOrWhiteSpace(uiUnitRaw) Then
                    uiUnitRaw = _connectorUiUnit
                End If
                Dim uiUnit As String = NormalizeUiUnit(uiUnitRaw)

                ' ✅ 납품 기준(BQC): UI의 선택/탭과 상관없이 아래 항목만 내보낸다.
                ' 1) 파라미터 불일치(Mismatch) / Shared Parameter 등록 필요
                ' 2) 대상과 거리가 0인데 미연결(Proximity + Distance=0)
                ' 3) ERROR
                Dim exportRows As List(Of Dictionary(Of String, Object)) = filteredTotal.Where(Function(r) ShouldExportIssueRow(r)).ToList()
                If _connectorExcludeEndDummy Then
                    exportRows = exportRows.Where(Function(r) Not ShouldExcludeEndDummyRow(r)).ToList()
                End If

                If exportRows Is Nothing OrElse exportRows.Count = 0 Then
                    System.Windows.Forms.MessageBox.Show("내보낼 항목이 없습니다.", "검토 결과", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                Dim mismatchCount As Integer = CountMismatches(exportRows)
                If exportRows Is Nothing OrElse exportRows.Count = 0 Then
                    exportRows = BuildEmptyConnectorRows()
                    mismatchCount = 0
                End If


                ' ✅ 선택한 reviewParams 중 이슈가 0건인 파라미터는 "오류가 없습니다" 안내행을 1건 추가(엑셀에서 검토 여부 확인 목적)
                Dim msgRows = BuildNoIssueMessageRows(exportRows, _lastConnectorReviewParams)
                If msgRows IsNot Nothing AndAlso msgRows.Count > 0 Then
                    ' placeholder(오류가 없습니다.)만 있을 경우 메시지 행으로 대체
                    If exportRows IsNot Nothing AndAlso exportRows.Count = 1 Then
                        Dim id1Text As String = ReadFieldInsensitive(exportRows(0), "Id1")
                        If Not String.IsNullOrWhiteSpace(id1Text) AndAlso id1Text.Contains("오류가 없습니다") Then
                            exportRows = msgRows
                        Else
                            exportRows = msgRows.Concat(exportRows).ToList()
                        End If
                    Else
                        exportRows = msgRows.Concat(exportRows).ToList()
                    End If
                End If

                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("connector:progress")
                Dim rvtBaseName As String = ""
                Try
                    Dim doc = app.ActiveUIDocument?.Document
                    If doc IsNot Nothing Then
                        Dim p As String = doc.PathName
                        If Not String.IsNullOrWhiteSpace(p) Then
                            rvtBaseName = System.IO.Path.GetFileNameWithoutExtension(p)
                        Else
                            rvtBaseName = doc.Title
                        End If
                    End If
                Catch
                    rvtBaseName = ""
                End Try

                Dim saved As String = SaveRowsToExcel(exportRows, mismatchCount, _connectorExtraParams, doAutoFit, "connector:progress", uiUnit, rvtBaseName, _lastConnectorReviewParams)

                SendToWeb("connector:saved", New With {.path = saved})

            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = "엑셀 내보내기 실패: " & ex.Message})
            End Try
        End Sub

#End Region

#Region "엑셀 입출력/유틸 (스키마 불변)"

        Private Function TryReadExcelAsDataTable() As DataTable
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Excel Files|*.xlsx;*.xls"
                ofd.Multiselect = False
                If ofd.ShowDialog() <> DialogResult.OK Then Return Nothing

                Dim filePath = ofd.FileName
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim wb As IWorkbook
                    Dim ext As String = System.IO.Path.GetExtension(filePath)
                    If ext IsNot Nothing AndAlso ext.Equals(".xls", StringComparison.OrdinalIgnoreCase) Then
                        wb = New HSSFWorkbook(fs)
                    Else
                        wb = New XSSFWorkbook(fs)
                    End If

                    Dim sh = wb.GetSheetAt(0)
                    Dim dt As New DataTable()

                    ' 헤더
                    Dim hr = sh.GetRow(sh.FirstRowNum)
                    If hr Is Nothing Then Return Nothing
                    For c = 0 To hr.LastCellNum - 1
                        Dim name = If(hr.GetCell(c)?.ToString(), $"C{c + 1}")
                        dt.Columns.Add(name)
                    Next

                    ' 데이터
                    For r = sh.FirstRowNum + 1 To sh.LastRowNum
                        Dim sr = sh.GetRow(r)
                        If sr Is Nothing Then Continue For
                        Dim dr = dt.NewRow()
                        For c = 0 To dt.Columns.Count - 1
                            dr(c) = If(sr.GetCell(c)?.ToString(), "")
                        Next
                        dt.Rows.Add(dr)
                    Next

                    Return dt
                End Using
            End Using
        End Function

        Private Function DataTableRows(dt As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            For Each r As DataRow In dt.Rows
                Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each c As DataColumn In dt.Columns
                    d(c.ColumnName) = If(r(c), "")
                Next
                list.Add(d)
            Next
            Return list
        End Function

        Private Function TryGetRowsFromPayload(payload As Object) As List(Of Dictionary(Of String, Object))
            If payload Is Nothing Then Return Nothing
            Try
                Dim d = TryCast(payload, IDictionary(Of String, Object))
                If d IsNot Nothing AndAlso d.ContainsKey("rows") Then
                    Return TryCast(d("rows"), List(Of Dictionary(Of String, Object)))
                End If
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function ReadField(r As Dictionary(Of String, Object), key As String) As String
            If r Is Nothing Then Return String.Empty
            If r.ContainsKey(key) AndAlso r(key) IsNot Nothing Then
                Return r(key).ToString()
            End If
            Return String.Empty
        End Function

        Private Shared Function ReadFieldInsensitive(r As Dictionary(Of String, Object), key As String) As String
            If r Is Nothing Then Return String.Empty
            For Each kv In r
                If kv.Key Is Nothing Then Continue For
                If String.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) Then
                    If kv.Value Is Nothing Then Return String.Empty
                    Return kv.Value.ToString()
                End If
            Next
            Return String.Empty
        End Function


        ' 선택한 reviewParams 중 existingRows에 한 건도 없는 파라미터에 대해 안내행을 생성한다.
        ' - ParamCompare: "[Param] 파라미터에 대한 연속성 오류가 없습니다."
        Private Shared Function BuildNoIssueMessageRows(existingRows As List(Of Dictionary(Of String, Object)),
                                                        reviewParams As List(Of String)) As List(Of Dictionary(Of String, Object))
            Dim result As New List(Of Dictionary(Of String, Object))()
            If reviewParams Is Nothing OrElse reviewParams.Count = 0 Then Return result

            Dim present As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If existingRows IsNot Nothing AndAlso existingRows.Count > 0 Then
                ' "오류가 없습니다." placeholder 1행만 있는 경우는 존재 파라미터 없음으로 처리
                If existingRows.Count = 1 Then
                    Dim id1Text As String = ReadFieldInsensitive(existingRows(0), "Id1")
                    If Not String.IsNullOrWhiteSpace(id1Text) AndAlso id1Text.Contains("오류가 없습니다") Then
                        ' ignore
                    Else
                        For Each r In existingRows
                            Dim p As String = ReadFieldInsensitive(r, "ParamName")
                            If Not String.IsNullOrWhiteSpace(p) Then present.Add(p.Trim())
                        Next
                    End If
                Else
                    For Each r In existingRows
                        Dim p As String = ReadFieldInsensitive(r, "ParamName")
                        If Not String.IsNullOrWhiteSpace(p) Then present.Add(p.Trim())
                    Next
                End If
            End If

            For Each raw In reviewParams
                Dim name As String = If(raw, "").Trim()
                If name = "" Then Continue For
                If present.Contains(name) Then Continue For

                Dim row As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                row("ParamName") = name
                row("Status") = "OK"
                row("ParamCompare") = $"[{name}] 파라미터에 대한 연속성 오류가 없습니다."
                result.Add(row)

                present.Add(name)
            Next

            Return result
        End Function

        ' ============================
        ' Connector Export: message/compare normalization
        ' ============================
        Private Shared Function NormalizeConnectorParamCompareForExport(row As Dictionary(Of String, Object)) As String
            If row Is Nothing Then Return String.Empty

            Dim status As String = ReadFieldInsensitive(row, "Status").Trim()
            Dim pc As String = ReadFieldInsensitive(row, "ParamCompare").Trim()

            ' Service가 Mismatch 시 ParamCompare에 메시지를 넣는 경우가 있어, 메시지 기반 보정
            If pc.IndexOf("연속성 오류가 없습니다", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Match"
            If pc.IndexOf("불일치", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Mismatch"

            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then
                If String.Equals(pc, "Match", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(pc, "Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(pc, "BothEmpty", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(pc, "N/A", StringComparison.OrdinalIgnoreCase) Then
                    Return pc
                End If
                Return "Match"
            End If

            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then
                Return "Mismatch"
            End If

            ' 비교 자체가 불가한 케이스만 N/A
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then
                Return "N/A"
            End If

            ' Proximity는 비교값이 있으면 그대로 사용(이전처럼 무조건 N/A로 떨어뜨리지 않음)
            If String.Equals(pc, "Match", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(pc, "Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(pc, "BothEmpty", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(pc, "N/A", StringComparison.OrdinalIgnoreCase) Then
                Return pc
            End If

            ' 최후: Value1/Value2로 비교 (pc가 비어있거나 알 수 없는 값인 경우)
            Dim v1 As String = ReadFieldInsensitive(row, "Value1").Trim()
            Dim v2 As String = ReadFieldInsensitive(row, "Value2").Trim()

            If v1.IndexOf("미등록", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               v2.IndexOf("미등록", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "N/A"
            End If

            If v1 = "" AndAlso v2 = "" Then Return "BothEmpty"
            If v1 <> "" OrElse v2 <> "" Then
                If String.Equals(v1, v2, StringComparison.Ordinal) Then Return "Match"
                Return "Mismatch"
            End If

            Return "N/A"
        End Function

        Private Shared Function BuildConnectorReviewTextForExport(row As Dictionary(Of String, Object)) As String
            If row Is Nothing Then Return String.Empty

            Dim status As String = ReadFieldInsensitive(row, "Status").Trim()
            Dim param As String = ReadFieldInsensitive(row, "ParamName").Trim()
            Dim pc As String = ReadFieldInsensitive(row, "ParamCompare").Trim()
            Dim err As String = ReadFieldInsensitive(row, "ErrorMessage").Trim()

            ' Proximity는 "연결 필요" 자체를 검토내용에 넣지 않고, 파라미터 비교 결과만 표시
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then
                Dim pcNorm As String = NormalizeConnectorParamCompareForExport(row)
                If String.Equals(pcNorm, "Mismatch", StringComparison.OrdinalIgnoreCase) Then
                    If param <> "" Then Return $"[{param}] 값이 서로 불일치. 확인이 필요합니다."
                    Return "값이 서로 불일치. 확인이 필요합니다."
                End If

                If String.Equals(pcNorm, "BothEmpty", StringComparison.OrdinalIgnoreCase) Then
                    If param <> "" Then Return $"[{param}] 파라미터 속성이 모두 누락되어있습니다."
                    Return "파라미터 속성이 모두 누락되어있습니다."
                End If

                ' Match / 기타는 OK로 취급
                If pc.IndexOf("연속성 오류가 없습니다", StringComparison.OrdinalIgnoreCase) >= 0 Then Return pc
                If param <> "" Then Return $"[{param}] 파라미터에 대한 연속성 오류가 없습니다."
                Return "연속성 오류가 없습니다."
            End If

            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then
                Dim pcNormOk As String = NormalizeConnectorParamCompareForExport(row)
                If String.Equals(pcNormOk, "BothEmpty", StringComparison.OrdinalIgnoreCase) Then
                    If param <> "" Then Return $"[{param}] 파라미터 속성이 모두 누락되어있습니다."
                    Return "파라미터 속성이 모두 누락되어있습니다."
                End If

                If pc.IndexOf("연속성 오류가 없습니다", StringComparison.OrdinalIgnoreCase) >= 0 Then Return pc
                If param <> "" Then Return $"[{param}] 파라미터에 대한 연속성 오류가 없습니다."
                Return "연속성 오류가 없습니다."
            End If

            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then
                If param <> "" Then Return $"[{param}] 값이 서로 불일치. 확인이 필요합니다."
                Return "값이 서로 불일치. 확인이 필요합니다."
            End If

            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then
                If param <> "" Then Return $"{param} : Shared Parameter 등록 필요"
                Return "Shared Parameter 등록 필요"
            End If

            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then
                Return "연결 대상 객체 없음"
            End If

            If String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then
                If err <> "" Then Return err
                If pc <> "" Then Return pc
                Return "ERROR"
            End If

            If pc.IndexOf("불일치", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               pc.IndexOf("연속성 오류가 없습니다", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               pc.IndexOf("오류", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return pc
            End If

            Return status
        End Function


        Private Shared Function ParseExtraParams(raw As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(raw) Then Return result

            Dim parts = raw.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each part In parts
                Dim name = part.Trim()
                If String.IsNullOrEmpty(name) Then Continue For
                If seen.Add(name) Then result.Add(name)
            Next
            Return result
        End Function

        Private Shared Function ParseParamsArray(raw As Object) As List(Of String)
            Dim result As New List(Of String)()
            If raw Is Nothing Then Return result

            Dim list = TryCast(raw, IEnumerable(Of Object))
            If list IsNot Nothing Then
                For Each item In list
                    If item Is Nothing Then Continue For
                    Dim name As String = item.ToString()
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    result.Add(name)
                Next
                Return result
            End If

            Dim listStr = TryCast(raw, IEnumerable(Of String))
            If listStr IsNot Nothing Then
                For Each name In listStr
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    result.Add(name)
                Next
            End If

            Return result
        End Function

        Private Shared Function ParseReviewParamsCsv(paramCsv As String) As List(Of String)
            Dim normalized As String = NormalizeParamsCsv(paramCsv)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(normalized) Then Return result
            For Each token In normalized.Split(","c)
                Dim name As String = If(token, String.Empty).Trim()
                If name = "" Then Continue For
                result.Add(name)
            Next
            Return result
        End Function

        Private Shared Function NormalizeParamsCsv(paramCsv As String) As String
            Dim tokens As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each raw In If(paramCsv, String.Empty).Split(","c)
                Dim name As String = If(raw, String.Empty).Trim()
                If name = "" Then Continue For
                If seen.Add(name) Then tokens.Add(name)
            Next
            Return String.Join(",", tokens)
        End Function

        Private Shared Function NormalizeUiUnit(raw As String) As String
            Dim u As String = If(raw, "").Trim().ToLowerInvariant()
            If u = "mm" OrElse u = "millimeter" OrElse u = "millimeters" Then Return "mm"
            Return "inch"
        End Function

        Private Shared Function IsNearConnection(r As Dictionary(Of String, Object)) As Boolean
            Dim conn As String = ReadField(r, "ConnectionType")
            If String.IsNullOrEmpty(conn) Then conn = ReadField(r, "Connection Type")
            If String.Equals(conn, "Near", StringComparison.OrdinalIgnoreCase) Then Return True
            If conn.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Dim status As String = ReadField(r, "Status")
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        ' Distance=0 이면서 Physical(연결됨)이 아닌 Proximity 케이스만 True
        ' - 서비스에서 "연결 대상 객체 없음"은 Distance 빈값으로 내려오도록 보정되어 있어, 여기서 0은 '실제로 대상이 있는' 0을 의미한다.
        Private Shared Function IsZeroDistanceNotConnected(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False

            Dim st As String = ReadField(row, "Status")
            Dim conn As String = ReadField(row, "ConnectionType")
            If String.IsNullOrEmpty(conn) Then conn = ReadField(row, "Connection Type")

            ' Proximity/미연결이 아닌 경우 제외
            If Not (String.Equals(st, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) OrElse
                    conn.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0) Then
                Return False
            End If

            ' Physical(연결 됨) 제외
            If conn.IndexOf("Physical", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return False
            End If

            ' 대상(Id2) 없는 행 제외
            Dim id2Raw As String = ReadField(row, "Id2")
            If String.IsNullOrWhiteSpace(id2Raw) Then Return False
            Dim id2 = id2Raw.Trim()
            If id2.StartsWith(",", StringComparison.Ordinal) Then id2 = id2.Substring(1).Trim()
            If String.IsNullOrWhiteSpace(id2) OrElse id2 = "0" Then Return False

            ' 거리 0 판정 (inch 기준)
            Dim distInch As Double = GetDistanceInch(row)
            If Double.IsNaN(distInch) Then Return False
            Return Math.Abs(distInch) < 0.0001R
        End Function

        Private Shared Function ShouldExcludeEndDummyRow(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            Dim family1 As String = ReadFieldInsensitive(row, "Family1")
            Dim family2 As String = ReadFieldInsensitive(row, "Family2")
            Dim type1 As String = ReadFieldInsensitive(row, "Type1")
            Dim type2 As String = ReadFieldInsensitive(row, "Type2")

            Dim parts As String() = {family1, family2, type1, type2}
            For Each part In parts
                If String.IsNullOrWhiteSpace(part) Then Continue For
                If part.IndexOf("End_", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
                If part.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function

        Private Shared Function ShouldExportIssueRow(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            Dim st = ReadField(row, "Status")
            If String.Equals(st, "ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            If IsMismatchRow(row) Then Return True
            Return IsZeroDistanceNotConnected(row)
        End Function

        Private Shared Function GetDistanceInch(row As Dictionary(Of String, Object)) As Double
            If row Is Nothing Then Return Double.NaN

            Dim v As Object = Nothing
            If row.ContainsKey("Distance (inch)") Then v = row("Distance (inch)")
            If v Is Nothing Then Return Double.NaN

            Try
                Return Convert.ToDouble(v)
            Catch
            End Try

            Dim s As String = v.ToString()
            If String.IsNullOrWhiteSpace(s) Then Return Double.NaN

            Dim d As Double
            If Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then Return d
            If Double.TryParse(s, d) Then Return d

            Return Double.NaN
        End Function

        Private Shared Function IsMismatchRow(r As Dictionary(Of String, Object)) As Boolean
            Dim status As String = ReadField(r, "Status")
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsMismatchStatus(status As String) As Boolean
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsIssueStatus(status As String) As Boolean
            If String.IsNullOrEmpty(status) Then Return False
            If String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "Parameter 없음", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsMatchOrOk(status As String) As Boolean
            If String.IsNullOrEmpty(status) Then Return False
            If String.Equals(status, "Match", StringComparison.OrdinalIgnoreCase) Then Return True
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return True
            Return False
        End Function

        Private Shared Function IsParamCompareMatch(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            Dim pc As String = ReadField(row, "ParamCompare")
            Return String.Equals(pc, "Match", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function ShouldExportToExcel(row As Dictionary(Of String, Object)) As Boolean
            If row Is Nothing Then Return False
            If IsParamCompareMatch(row) Then Return False
            Dim status As String = ReadField(row, "Status")
            Return IsIssueStatus(status)
        End Function

        Private Shared Function BuildEmptyConnectorRows() As List(Of Dictionary(Of String, Object))
            Dim row As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            row("Id1") = "오류가 없습니다."
            Return New List(Of Dictionary(Of String, Object)) From {row}
        End Function

        Private Shared Function ShouldIncludeRow(r As Dictionary(Of String, Object)) As Boolean
            If r Is Nothing Then Return False
            If IsParamCompareMatch(r) Then Return False
            If IsMismatchRow(r) Then Return True
            If IsNearConnection(r) Then Return True
            Dim status As String = ReadField(r, "Status")
            Return IsIssueStatus(status)
        End Function

        Private Function CountMismatches(rows As List(Of Dictionary(Of String, Object))) As Integer
            If rows Is Nothing Then Return 0
            Dim cnt As Integer = 0
            For Each row In rows
                Dim status As String = Nothing
                If row IsNot Nothing AndAlso row.ContainsKey("Status") AndAlso row("Status") IsNot Nothing Then
                    status = row("Status").ToString()
                End If
                If IsIssueStatus(status) Then
                    cnt += 1
                End If
            Next
            Return cnt
        End Function

        Private Function SaveRowsToExcel(totalRows As List(Of Dictionary(Of String, Object)),
                                         Optional mismatchCount As Integer = -1,
                                         Optional extraParams As List(Of String) = Nothing,
                                         Optional doAutoFit As Boolean = False,
                                         Optional progressChannel As String = Nothing,
                                         Optional uiUnit As String = "inch",
                                         Optional rvtBaseName As String = "",
                                         Optional reviewParams As List(Of String) = Nothing) As String
            Dim todayToken As String = Date.Now.ToString("yyMMdd")

            ' ✅ 파일명 건수는 "최종 엑셀에 실제로 저장되는 오류 행 수" 기준으로 계산
            Dim count As Integer = 0
            If totalRows IsNot Nothing Then
                count = totalRows.Count
                ' "오류가 없습니다." 안내 1행만 있는 경우는 0건 처리
                If count = 1 Then
                    Dim id1Text As String = SafeCellString(totalRows(0), "Id1")
                    If Not String.IsNullOrWhiteSpace(id1Text) AndAlso id1Text.Contains("오류가 없습니다") Then
                        count = 0
                    End If
                End If
            End If

            ' ✅ 기본 파일명: RVT 파일명 규칙 기반 + 고정 suffix + (n건)
            Dim defaultName As String = BuildTradeReviewDefaultExcelName(rvtBaseName, count)

            ' ✅ 멀티 RVT 실행 시(2개 이상) 기본 저장 파일명 규칙:
            '   - 규칙(ExtractTradePrefix) 일치: [첫번째 파일 규칙 prefix]+nFile_공종검토 (n = 파일수-1)
            '   - 규칙 불일치: Parameter 연속성검토_Selected n Files
            Try
                Dim filesInOrder As List(Of String) = CollectDistinctRvtFilesInOrder(totalRows)
                If filesInOrder IsNot Nothing AndAlso filesInOrder.Count >= 2 Then
                    Dim firstBase As String = System.IO.Path.GetFileNameWithoutExtension(filesInOrder(0))
                    Dim prefix As String = ExtractTradePrefix(firstBase)
                    If Not String.IsNullOrWhiteSpace(prefix) Then
                        Dim addN As Integer = Math.Max(0, filesInOrder.Count - 1)
                        defaultName = $"{prefix}+{addN}File_공종검토"
                    Else
                        defaultName = $"Parameter 연속성검토_Selected {filesInOrder.Count} Files"
                    End If
                    defaultName = SanitizeFileName(defaultName)
                    If Not defaultName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then
                        defaultName &= ".xlsx"
                    End If
                End If
            Catch
                ' ignore - fallback to existing naming
            End Try

            If String.IsNullOrWhiteSpace(defaultName) Then
                defaultName = $"{todayToken}_커넥터기반 속성값 검토 결과_{count}개.xlsx"
            End If
            Dim totalCount As Integer = If(totalRows, New List(Of Dictionary(Of String, Object))()).Count
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset(progressChannel)
            Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalCount, Nothing, True)
            LogAutoFitDecision(doAutoFit, "UiBridgeExternalEvent.SaveRowsToExcel")

            Try
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                    sfd.FileName = defaultName
                    If sfd.ShowDialog() <> DialogResult.OK Then Throw New OperationCanceledException()

                    Dim extrasSource As List(Of String) = If(extraParams, New List(Of String)())
                    If (extrasSource Is Nothing OrElse extrasSource.Count = 0) AndAlso totalRows IsNot Nothing AndAlso totalRows.Count > 0 Then
                        extrasSource = InferExtrasFromRow(totalRows(0))
                    End If
                    Dim extrasHeaders = BuildExtraHeaders(extrasSource)
                    Dim headersTotal = BuildHeaders(extrasHeaders)
                    Dim totalBase = totalRows.Select(Function(r) StripExtras(r, extrasHeaders)).ToList()

                    Dim sheetTables = BuildConnectorSheetTables(totalBase, headersTotal, extrasHeaders, uiUnit, reviewParams)

                    Dim savePath = sfd.FileName
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", totalCount, totalCount, Nothing, True)
                    SaveConnectorRowsMultiSheet(savePath, sheetTables, doAutoFit, progressChannel)
                    Global.KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey("connector", savePath, autoFit:=doAutoFit, excelMode:=If(doAutoFit, "normal", "fast"))

                    Try
                        ApplyConnectorReviewContentIssueStyles(savePath)
                    Catch
                        ' ignore
                    End Try

                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_SAVE", "파일 저장 중", totalCount, totalCount, Nothing, True)
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "AUTOFIT", If(doAutoFit, "AutoFit 적용", "빠른 모드: AutoFit 생략"), totalCount, totalCount, Nothing, True)
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기 완료", totalCount, totalCount, 100.0R, True)
                    Return savePath
                End Using
            Catch ex As OperationCanceledException
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "DONE", "엑셀 내보내기가 취소되었습니다.", 0, totalCount, 100.0R, True)
                Return String.Empty
            Catch ex As Exception
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "ERROR", ex.Message, 0, totalCount, Nothing, True)
                Throw
            End Try
        End Function

        Private Function BuildConnectorSheetTables(rows As List(Of Dictionary(Of String, Object)),
                                                  headersTotal As List(Of String),
                                                  extrasHeaders As List(Of String),
                                                  uiUnit As String,
                                                  reviewParams As List(Of String)) As List(Of KeyValuePair(Of String, DataTable))
            Dim sheets As New List(Of KeyValuePair(Of String, DataTable))()
            Dim baseRows = If(rows, New List(Of Dictionary(Of String, Object))())

            ' ✅ RVT를 2개 이상 검토한 경우: File 기준으로 시트를 분리한다.
            '    (단일 파일일 때는 기존 동작(ParamName 기준) 유지)
            Dim filesInOrder As List(Of String) = Nothing
            Try
                filesInOrder = CollectDistinctRvtFilesInOrder(baseRows)
            Catch
                filesInOrder = New List(Of String)()
            End Try

            If filesInOrder IsNot Nothing AndAlso filesInOrder.Count >= 2 Then
                ' File 값이 없는 행(예: "오류가 없습니다" 안내행)은 모든 파일 시트에 포함시킨다.
                Dim globalRows As List(Of Dictionary(Of String, Object)) =
                    baseRows.Where(Function(r) String.IsNullOrWhiteSpace(ReadFieldInsensitive(r, "File"))).ToList()

                Dim used As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each f In filesInOrder
                    Dim fileKey As String = If(f, "").Trim()

                    Dim rowsForFile As List(Of Dictionary(Of String, Object)) =
                        baseRows.Where(Function(r)
                                           Dim rf As String = ReadFieldInsensitive(r, "File")
                                           If String.IsNullOrWhiteSpace(rf) Then Return False
                                           Return String.Equals(rf.Trim(), fileKey, StringComparison.OrdinalIgnoreCase)
                                       End Function).ToList()

                    Dim merged As New List(Of Dictionary(Of String, Object))()
                    If globalRows IsNot Nothing AndAlso globalRows.Count > 0 Then merged.AddRange(globalRows)
                    If rowsForFile IsNot Nothing AndAlso rowsForFile.Count > 0 Then merged.AddRange(rowsForFile)

                    Dim sheetRaw As String = ""
                    Try
                        sheetRaw = System.IO.Path.GetFileNameWithoutExtension(fileKey)
                    Catch
                        sheetRaw = fileKey
                    End Try
                    If String.IsNullOrWhiteSpace(sheetRaw) Then sheetRaw = fileKey
                    Dim sheetName As String = SafeExcelSheetName(sheetRaw)

                    ' 시트명 중복 방지 (_2, _3...)
                    Dim baseName As String = sheetName
                    Dim idx As Integer = 2
                    While used.Contains(sheetName)
                        Dim suffix As String = "_" & idx.ToString()
                        Dim maxLen As Integer = 31 - suffix.Length
                        Dim head As String = baseName
                        If head.Length > maxLen Then head = head.Substring(0, maxLen)
                        sheetName = head & suffix
                        idx += 1
                    End While
                    used.Add(sheetName)

                    Dim dt As DataTable = BuildConnectorExportDataTable(headersTotal, merged, uiUnit)
                    NormalizeConnectorExportDataTableSchema(dt, extrasHeaders)
                    LogConnectorExportHeaders(dt)
                    If dt.Rows.Count = 0 Then
                        Global.KKY_Tool_Revit.Infrastructure.ExcelCore.EnsureNoDataRow(dt, "검토 결과가 없습니다.")
                    End If
                    sheets.Add(New KeyValuePair(Of String, DataTable)(sheetName, dt))
                Next

                Return sheets
            End If

            ' ✅ 기본(단일 파일/파일 정보 없음): ParamName 기준 시트 분리
            Dim grouped As New Dictionary(Of String, List(Of Dictionary(Of String, Object)))(StringComparer.OrdinalIgnoreCase)
            For Each row In baseRows
                Dim paramName As String = ReadField(row, "ParamName")
                If String.IsNullOrWhiteSpace(paramName) Then paramName = "Connector Diagnostics"
                If Not grouped.ContainsKey(paramName) Then grouped(paramName) = New List(Of Dictionary(Of String, Object))()
                grouped(paramName).Add(row)
            Next

            If grouped.Count = 0 Then
                Dim defaultSheetName As String = "Connector Diagnostics"
                If reviewParams IsNot Nothing AndAlso reviewParams.Count > 0 Then
                    defaultSheetName = reviewParams(0)
                End If
                grouped(defaultSheetName) = New List(Of Dictionary(Of String, Object))()
            End If

            For Each kv In grouped
                Dim dt As DataTable = BuildConnectorExportDataTable(headersTotal, kv.Value, uiUnit)
                NormalizeConnectorExportDataTableSchema(dt, extrasHeaders)
                LogConnectorExportHeaders(dt)
                If dt.Rows.Count = 0 Then
                    Global.KKY_Tool_Revit.Infrastructure.ExcelCore.EnsureNoDataRow(dt, "검토 결과가 없습니다.")
                End If
                sheets.Add(New KeyValuePair(Of String, DataTable)(SafeExcelSheetName(kv.Key), dt))
            Next

            Return sheets
        End Function

        Private Shared Function SafeExcelSheetName(raw As String) As String
            Dim name As String = If(raw, String.Empty).Trim()
            If name = "" Then name = "Connector Diagnostics"
            Dim invalidChars As Char() = New Char() {"["c, "]"c, ":"c, "*"c, "?"c, "/"c, "\"c}
            For Each ch In invalidChars
                name = name.Replace(ch, "_"c)
            Next
            If name.Length > 31 Then name = name.Substring(0, 31)
            If String.IsNullOrWhiteSpace(name) Then name = "Connector Diagnostics"
            Return name
        End Function

        Private Shared Sub SaveConnectorRowsMultiSheet(savePath As String,
                                                       sheetTables As List(Of KeyValuePair(Of String, DataTable)),
                                                       doAutoFit As Boolean,
                                                       progressChannel As String)
            Dim tables = If(sheetTables, New List(Of KeyValuePair(Of String, DataTable))())
            If tables.Count = 0 Then
                Dim dt As New DataTable("Connector Diagnostics")
                dt.Columns.Add("Id1", GetType(String))
                Global.KKY_Tool_Revit.Infrastructure.ExcelCore.EnsureNoDataRow(dt, "검토 결과가 없습니다.")
                tables.Add(New KeyValuePair(Of String, DataTable)("Connector Diagnostics", dt))
            End If

            Global.KKY_Tool_Revit.Infrastructure.ExcelCore.SaveXlsxMulti(savePath, tables, doAutoFit, progressChannel, sheetKeyOverride:="connector", exportKind:="connector")
        End Sub

        Private Shared Function BuildConnectorExportDataTable(headers As List(Of String),
                                                              rows As List(Of Dictionary(Of String, Object)),
                                                              uiUnit As String) As DataTable
            Dim dt As New DataTable("Connector Diagnostics")
            Dim safeHeaders As List(Of String) = If(headers, New List(Of String)())
            For Each h In safeHeaders
                Dim colName As String = If(h, "").Trim()
                If String.IsNullOrWhiteSpace(colName) Then colName = "Column"
                Dim uniqueName As String = colName
                Dim suffix As Integer = 1
                While dt.Columns.Contains(uniqueName)
                    suffix += 1
                    uniqueName = colName & "_" & suffix.ToString()
                End While
                dt.Columns.Add(uniqueName, GetType(String))
            Next

            Dim safeRows = If(rows, New List(Of Dictionary(Of String, Object))())
            For Each row In safeRows
                Dim dr = dt.NewRow()
                For Each col As DataColumn In dt.Columns
                    dr(col.ColumnName) = ReadConnectorExportCell(row, col.ColumnName, uiUnit)
                Next
                dt.Rows.Add(dr)
            Next

            Return dt
        End Function

        Private Shared Function ReadConnectorExportCell(row As Dictionary(Of String, Object),
                                                        header As String,
                                                        uiUnit As String) As String
            If row Is Nothing Then Return String.Empty

            ' ✅ 엑셀 F열(검토내용)에 메시지 출력, ParamCompare는 순수 비교값만 유지
            If String.Equals(header, "검토내용", StringComparison.Ordinal) Then
                Return BuildConnectorReviewTextForExport(row)
            End If
            If String.Equals(header, "ParamCompare", StringComparison.Ordinal) Then
                Return NormalizeConnectorParamCompareForExport(row)
            End If
            If String.Equals(header, "비고(답변)", StringComparison.Ordinal) Then
                Return ""
            End If

            If String.Equals(header, "Distance", StringComparison.OrdinalIgnoreCase) Then
                Dim distRaw As String = SafeCellString(row, header)
                If String.IsNullOrWhiteSpace(distRaw) Then Return String.Empty
                Dim distValue As Double
                If Double.TryParse(distRaw, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, distValue) OrElse
                   Double.TryParse(distRaw, distValue) Then
                    Return ConvertDistanceForUi(distValue, uiUnit)
                End If
                Return distRaw
            End If

            Return SafeCellString(row, header)
        End Function

        Private Shared Function ConvertDistanceForUi(distanceInch As Double, uiUnit As String) As String
            Dim v As Double = distanceInch
            If String.Equals(uiUnit, "mm", StringComparison.OrdinalIgnoreCase) Then
                v = distanceInch * 25.4R
            End If
            Return Math.Round(v, 3).ToString(Globalization.CultureInfo.InvariantCulture)
        End Function

        ' 테두리/헤더/색상 스타일 헬퍼 (같은 워크북 내 공유)
        Private Shared Function CreateBorderedStyle(wb As XSSFWorkbook) As ICellStyle
            Dim st As ICellStyle = wb.CreateCellStyle()
            ' 행 높이 자동 증가(82.5 등) 이슈 방지: WrapText를 명시적으로 끔
            st.WrapText = False
            st.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
            st.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
            Return st
        End Function

        Private Shared Function CreateHeaderStyle(wb As XSSFWorkbook, baseStyle As ICellStyle) As ICellStyle
            Dim st As XSSFCellStyle = CType(wb.CreateCellStyle(), XSSFCellStyle)
            st.CloneStyleFrom(baseStyle)
            st.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
            ' use available ctor for current NPOI version
            st.SetFillForegroundColor(New XSSFColor(New Byte() {&H2A, &H3B, &H52}))

            Dim f As XSSFFont = CType(wb.CreateFont(), XSSFFont)
            f.IsBold = True
            f.Color = IndexedColors.White.Index
            st.SetFont(f)
            Return st
        End Function

        Private Shared Function CreateFillStyle(wb As XSSFWorkbook, baseStyle As ICellStyle, rgb As Byte()) As ICellStyle
            Dim st As XSSFCellStyle = CType(wb.CreateCellStyle(), XSSFCellStyle)
            st.CloneStyleFrom(baseStyle)
            st.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
            ' same ctor form as header style to keep compatibility with the current NPOI version
            st.SetFillForegroundColor(New XSSFColor(rgb))
            Return st
        End Function

        Private Shared Function SafeCellString(row As Dictionary(Of String, Object), key As String) As String
            If row Is Nothing OrElse String.IsNullOrEmpty(key) OrElse Not row.ContainsKey(key) Then Return String.Empty
            Dim v = row(key)
            Return If(v Is Nothing, String.Empty, v.ToString())
        End Function

        Private Shared Function BuildBaseHeaders() As List(Of String)
            ' ✅ Connector 엑셀 내보내기 스키마(납품 요구사항)
            ' - Category2 ↔ Family1 사이에 "검토내용", "비고(답변)" 2열 추가(항상 빈값)
            ' - Status, ErrorMessage 컬럼은 엑셀에 출력하지 않음
            Return New List(Of String) From {
                "File",
                "Id1", "Id2",
                "Category1", "Category2",
                            "검토내용", "비고(답변)",
                "Family1", "Family2",
                "Distance (inch)",
                "ConnectionType",
                "ParamName",
                "Value1", "Value2",
                "ParamCompare"
            }
        End Function

        Private Shared Function BuildExtraHeaders(extras As IList(Of String)) As List(Of String)
            Dim list As New List(Of String)()
            If extras Is Nothing Then Return list

            For Each name In extras
                list.Add($"{name}(ID1)")
                list.Add($"{name}(ID2)")
            Next

            Return list
        End Function

        Private Shared Function BuildHeaders(extras As IList(Of String)) As List(Of String)
            Dim headers = BuildBaseHeaders()
            If extras IsNot Nothing Then
                For Each name In extras
                    headers.Add(name)
                Next
            End If
            Return headers
        End Function

        ' 엑셀 저장 직전: 컬럼 스키마를 강제로 맞춘다(헤더 꼬임 방지)
        Private Shared Sub NormalizeConnectorExportDataTableSchema(dt As DataTable, extrasHeaders As IList(Of String))
            If dt Is Nothing Then Return

            Dim desired As New List(Of String) From {
                "File",
                "Id1", "Id2",
                "Category1", "Category2",
                            "검토내용", "비고(답변)",
                "Family1", "Family2",
                "Distance (inch)",
                "ConnectionType",
                "ParamName",
                "Value1", "Value2",
                "ParamCompare"
            }

            If extrasHeaders IsNot Nothing Then
                For Each h In extrasHeaders
                    Dim name As String = If(h, "").Trim()
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    desired.Add(name)
                Next
            End If

            ' 1) 필요한 컬럼 추가 + 순서 고정
            For i As Integer = 0 To desired.Count - 1
                Dim name As String = desired(i)
                If Not dt.Columns.Contains(name) Then
                    dt.Columns.Add(name, GetType(String))
                End If
                Try
                    dt.Columns(name).SetOrdinal(i)
                Catch
                    ' SetOrdinal 실패는 무시(동일 컬럼명 충돌 등)
                End Try
            Next

            ' 2) 불필요 컬럼 제거(Status/ErrorMessage 등)
            For i As Integer = dt.Columns.Count - 1 To 0 Step -1
                Dim name As String = dt.Columns(i).ColumnName
                If Not desired.Contains(name) Then
                    dt.Columns.RemoveAt(i)
                End If
            Next

            ' 3) 신규 메모열은 항상 빈칸(Null → "")
            If dt.Columns.Contains("검토내용") Then
                For Each r As DataRow In dt.Rows
                    If r Is Nothing Then Continue For
                    If r.IsNull("검토내용") Then r("검토내용") = ""
                Next
            End If
            If dt.Columns.Contains("비고(답변)") Then
                For Each r As DataRow In dt.Rows
                    If r Is Nothing Then Continue For
                    If r.IsNull("비고(답변)") Then r("비고(답변)") = ""
                Next
            End If
        End Sub

        ' 엑셀 내보내기 헤더를 로그로 남김(실제 저장 스키마 확인용)
        Private Sub LogConnectorExportHeaders(dt As DataTable)
            Try
                If dt Is Nothing Then Return
                Dim cols As New List(Of String)()
                For Each c As DataColumn In dt.Columns
                    cols.Add(c.ColumnName)
                Next
                LogDebug("[connector][excel] export headers => " & String.Join(" | ", cols))
            Catch
            End Try
        End Sub


        Private Shared Function InferExtrasFromRow(row As Dictionary(Of String, Object)) As List(Of String)
            Dim extras As New List(Of String)()
            If row Is Nothing Then Return extras

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each key In row.Keys
                If key Is Nothing Then Continue For
                If key.EndsWith("(ID1)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = key.Substring(0, key.Length - "(ID1)".Length)
                    If seen.Add(name) Then extras.Add(name)
                End If
            Next
            Return extras
        End Function

        Private Shared Function BuildReviewRows(rows As List(Of Dictionary(Of String, Object)), extras As IList(Of String)) As List(Of Dictionary(Of String, Object))
            If rows Is Nothing Then Return New List(Of Dictionary(Of String, Object))()

            Dim dirRows As New Dictionary(Of String, List(Of Dictionary(Of String, Object)))(StringComparer.Ordinal)
            Dim pairMembers As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim id1 = ToIntLocal(ReadField(row, "Id1"))
                Dim id2 = ToIntLocal(ReadField(row, "Id2"))
                If id1 = 0 AndAlso id2 = 0 Then Continue For

                Dim dirKey = $"{id1}->{id2}"
                Dim pairKey = If(id1 <= id2, $"{id1}_{id2}", $"{id2}_{id1}")

                If Not dirRows.ContainsKey(dirKey) Then dirRows(dirKey) = New List(Of Dictionary(Of String, Object))()
                dirRows(dirKey).Add(row)

                If Not pairMembers.ContainsKey(pairKey) Then pairMembers(pairKey) = New HashSet(Of String)(StringComparer.Ordinal)
                pairMembers(pairKey).Add(dirKey)
            Next

            Dim result As New List(Of Dictionary(Of String, Object))()

            For Each kv In pairMembers
                Dim ids = kv.Key.Split("_"c)
                If ids.Length <> 2 Then Continue For
                Dim a As Integer = ToIntLocal(ids(0))
                Dim b As Integer = ToIntLocal(ids(1))
                Dim keyAB = $"{a}->{b}"
                Dim keyBA = $"{b}->{a}"

                Dim bestAB = SelectBestRow(dirRows, keyAB)
                Dim bestBA = SelectBestRow(dirRows, keyBA)

                If bestAB Is Nothing AndAlso bestBA IsNot Nothing Then bestAB = SwapRow(bestBA)
                If bestBA Is Nothing AndAlso bestAB IsNot Nothing Then bestBA = SwapRow(bestAB)

                If bestAB IsNot Nothing Then result.Add(AppendExtrasForId1(bestAB, extras))
                If bestBA IsNot Nothing Then result.Add(AppendExtrasForId1(bestBA, extras))
            Next

            Return result
        End Function

        Private Shared Function StripExtras(row As Dictionary(Of String, Object), Optional extras As IList(Of String) = Nothing) As Dictionary(Of String, Object)
            Dim headers = BuildHeaders(extras)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            For Each key In headers
                If row.ContainsKey(key) Then d(key) = row(key)
            Next
            Return d
        End Function

        Private Shared Function CloneRow(row As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            If row Is Nothing Then Return d
            For Each kv In row
                d(kv.Key) = kv.Value
            Next
            Return d
        End Function

        Private Shared Function RowPriority(row As Dictionary(Of String, Object)) As Integer
            Dim status = SafeCellString(row, "Status")
            Dim conn = SafeCellString(row, "ConnectionType")

            If String.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) Then Return 5
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return 4
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return 4
            If conn.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 OrElse String.Equals(conn, "Near", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return 2
            Return 1
        End Function

        Private Shared Function SelectBestRow(map As Dictionary(Of String, List(Of Dictionary(Of String, Object))), key As String) As Dictionary(Of String, Object)
            If Not map.ContainsKey(key) Then Return Nothing
            Dim best As Dictionary(Of String, Object) = Nothing

            For Each row In map(key)
                If row Is Nothing Then Continue For
                If best Is Nothing Then
                    best = row
                    Continue For
                End If

                Dim pr = RowPriority(row)
                Dim pb = RowPriority(best)
                If pr > pb Then
                    best = row
                ElseIf pr = pb Then
                    Dim dr = ToDoubleLocal(SafeCellString(row, "Distance (inch)"))
                    Dim db = ToDoubleLocal(SafeCellString(best, "Distance (inch)"))
                    If dr < db Then best = row
                End If
            Next

            Return best
        End Function

        Private Shared Function SwapRow(row As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            If row Is Nothing Then Return Nothing

            Dim swapped As New Dictionary(Of String, Object)(StringComparer.Ordinal)

            ' ✅ 값 보존: File / 메모열은 교환 대상이 아니므로 그대로 유지
            swapped("File") = SafeCellString(row, "File")
            swapped("검토내용") = SafeCellString(row, "검토내용")
            swapped("비고(답변)") = SafeCellString(row, "비고(답변)")

            ' ✅ ID1/ID2 쌍 스왑
            swapped("Id1") = SafeCellString(row, "Id2")
            swapped("Id2") = SafeCellString(row, "Id1")
            swapped("Category1") = SafeCellString(row, "Category2")
            swapped("Category2") = SafeCellString(row, "Category1")
            swapped("Family1") = SafeCellString(row, "Family2")
            swapped("Family2") = SafeCellString(row, "Family1")
            swapped("Distance (inch)") = SafeCellString(row, "Distance (inch)")
            swapped("ConnectionType") = SafeCellString(row, "ConnectionType")
            swapped("ParamName") = SafeCellString(row, "ParamName")
            swapped("Value1") = SafeCellString(row, "Value2")
            swapped("Value2") = SafeCellString(row, "Value1")
            swapped("ParamCompare") = SafeCellString(row, "ParamCompare")

            ' 내부(UI 렌더/필터)에서 사용할 수 있으므로 유지 (엑셀 헤더에서는 제외됨)
            swapped("Status") = SafeCellString(row, "Status")
            swapped("ErrorMessage") = SafeCellString(row, "ErrorMessage")

            ' Extra Params(ID1/ID2) 스왑
            For Each kv In row
                If kv.Key Is Nothing Then Continue For

                If kv.Key.EndsWith("(ID1)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = kv.Key.Substring(0, kv.Key.Length - "(ID1)".Length)
                    swapped($"{name}(ID1)") = SafeCellString(row, $"{name}(ID2)")
                ElseIf kv.Key.EndsWith("(ID2)", StringComparison.OrdinalIgnoreCase) Then
                    Dim name = kv.Key.Substring(0, kv.Key.Length - "(ID2)".Length)
                    swapped($"{name}(ID2)") = SafeCellString(row, $"{name}(ID1)")
                End If
            Next

            Return swapped
        End Function

        Private Shared Function AppendExtrasForId1(row As Dictionary(Of String, Object), extras As IList(Of String)) As Dictionary(Of String, Object)
            Dim d = StripExtras(row, extras)
            If extras IsNot Nothing Then
                For Each name In extras
                    If row.ContainsKey(name) Then d(name) = SafeCellString(row, name)
                    Dim key1 = $"{name}(ID1)"
                    Dim key2 = $"{name}(ID2)"
                    d(key1) = SafeCellString(row, key1)
                    d(key2) = SafeCellString(row, key2)
                    If Not d.ContainsKey(name) Then d(name) = SafeCellString(row, key1)
                Next
            End If
            Return d
        End Function

        Private Shared Function BuildTotalRows(rows As List(Of Dictionary(Of String, Object))) As List(Of Dictionary(Of String, Object))
            If rows Is Nothing Then Return New List(Of Dictionary(Of String, Object))()
            Dim pairRows As New Dictionary(Of String, Dictionary(Of String, Object))(StringComparer.Ordinal)

            For Each row In rows
                If row Is Nothing Then Continue For
                Dim id1 = ToIntLocal(ReadField(row, "Id1"))
                Dim id2 = ToIntLocal(ReadField(row, "Id2"))
                Dim key = If(id1 <= id2, $"{id1}_{id2}", $"{id2}_{id1}")

                If Not pairRows.ContainsKey(key) Then
                    pairRows(key) = row
                Else
                    Dim cur = row
                    Dim best = pairRows(key)
                    Dim pr = RowPriority(cur)
                    Dim pb = RowPriority(best)
                    If pr > pb Then
                        pairRows(key) = cur
                    ElseIf pr = pb Then
                        Dim dr = ToDoubleLocal(SafeCellString(cur, "Distance (inch)"))
                        Dim db = ToDoubleLocal(SafeCellString(best, "Distance (inch)"))
                        If dr < db Then pairRows(key) = cur
                    End If
                End If
            Next

            Return pairRows.Values.Select(Function(r) CloneRow(r)).ToList()
        End Function

        Private Shared Sub WriteSheet(wb As XSSFWorkbook,
                                      sheetName As String,
                                      headers As List(Of String),
                                      rows As List(Of Dictionary(Of String, Object)),
                                      headerStyle As ICellStyle,
                                      baseStyle As ICellStyle,
                                      matchStyle As ICellStyle,
                                      mismatchStyle As ICellStyle,
                                      nearStyle As ICellStyle,
                                      errorStyle As ICellStyle,
                                      Optional progressChannel As String = Nothing,
                                      Optional ByRef written As Integer = 0,
                                      Optional totalRows As Integer = 0,
                                      Optional doAutoFit As Boolean = False,
                                      Optional uiUnit As String = "inch")
            Dim sh = wb.CreateSheet(sheetName)

            Dim headerRow = sh.CreateRow(0)
            headerRow.Height = -1
            For i = 0 To headers.Count - 1
                Dim c = headerRow.CreateCell(i)
                Dim headerText As String = headers(i)
                If String.Equals(headerText, "Distance (inch)", StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(uiUnit, "mm", StringComparison.OrdinalIgnoreCase) Then
                    headerText = "Distance (mm)"
                End If
                c.SetCellValue(headerText)
                c.CellStyle = headerStyle
            Next

            sh.CreateFreezePane(0, 1)

            If headers.Count > 0 Then
                ' AutoFilter 범위는 헤더 + 데이터 전체 범위로 지정
                Dim lastRowIdx As Integer = 0
                If rows IsNot Nothing Then lastRowIdx = Math.Max(0, rows.Count)
                Dim range As New NPOI.SS.Util.CellRangeAddress(0, lastRowIdx, 0, headers.Count - 1)
                sh.SetAutoFilter(range)
            End If

            If rows IsNot Nothing Then
                Dim r As Integer = 1
                For Each row In rows
                    Dim sr = sh.CreateRow(r) : r += 1
                    sr.Height = -1

                    Dim statusVal As String = SafeCellString(row, "Status")
                    Dim connVal As String = SafeCellString(row, "ConnectionType")
                    Dim styleToUse As ICellStyle = baseStyle

                    If String.Equals(statusVal, "ERROR", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = errorStyle
                    ElseIf IsMismatchStatus(statusVal) Then
                        styleToUse = mismatchStyle
                    ElseIf String.Equals(statusVal, "OK", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = matchStyle
                    ElseIf String.Equals(statusVal, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(statusVal, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then
                        styleToUse = nearStyle
                    ElseIf String.Equals(connVal.Trim(), "Near", StringComparison.OrdinalIgnoreCase) OrElse connVal.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        styleToUse = nearStyle
                    End If

                    For c = 0 To headers.Count - 1
                        Dim key = headers(c)
                        Dim v As Object = Nothing
                        If row.ContainsKey(key) Then v = row(key)
                        Dim cell = sr.CreateCell(c)

                        Dim text As String = If(v Is Nothing, "", v.ToString())
                        If String.Equals(key, "Id2", StringComparison.OrdinalIgnoreCase) Then
                            Dim t = text.Trim()
                            If t = "" OrElse t = "0" Then
                                text = ""
                            ElseIf Not t.StartsWith(",", StringComparison.Ordinal) Then
                                text = "," & t
                            Else
                                text = t
                            End If
                        End If

                        If String.Equals(key, "Distance (inch)", StringComparison.OrdinalIgnoreCase) Then
                            Dim d As Double
                            Dim ok As Boolean = False
                            If v IsNot Nothing Then
                                Try
                                    d = Convert.ToDouble(v)
                                    ok = True
                                Catch
                                End Try
                            End If
                            If Not ok Then
                                Dim tmp As Double
                                If Double.TryParse(text.Trim(), tmp) Then
                                    d = tmp
                                    ok = True
                                End If
                            End If

                            If ok Then
                                If String.Equals(uiUnit, "mm", StringComparison.OrdinalIgnoreCase) Then
                                    d = d * 25.4R
                                End If
                                cell.SetCellValue(d)
                            Else
                                cell.SetCellValue(text)
                            End If
                        Else
                            cell.SetCellValue(text)
                        End If
                        cell.CellStyle = styleToUse
                    Next
                    written += 1
                    Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Report(progressChannel, "EXCEL_WRITE", "엑셀 데이터 작성", written, totalRows)
                Next
            End If

            If doAutoFit Then
                ApplyFastColumnWidths(sh, headers, rows)
            End If
        End Sub

        Private Shared Sub ApplyFastColumnWidths(sh As ISheet, headers As List(Of String), rows As List(Of Dictionary(Of String, Object)))
            If sh Is Nothing OrElse headers Is Nothing Then Return

            Const MAX_SAMPLE As Integer = 2000
            Const MIN_CHARS As Integer = 6
            Const MAX_CHARS As Integer = 60
            Const MAX_WIDTH As Integer = 255 * 256

            Dim maxLens(headers.Count - 1) As Integer
            For i = 0 To headers.Count - 1
                maxLens(i) = If(headers(i) Is Nothing, 0, headers(i).Length)
            Next

            Dim sampleCount As Integer = 0
            If rows IsNot Nothing Then sampleCount = Math.Min(rows.Count, MAX_SAMPLE)

            For i = 0 To sampleCount - 1
                Dim row = rows(i)
                If row Is Nothing Then Continue For

                For c = 0 To headers.Count - 1
                    Dim key = headers(c)
                    Dim v As Object = Nothing
                    If row.ContainsKey(key) Then v = row(key)
                    Dim text As String = If(v Is Nothing, String.Empty, v.ToString())
                    If text.Length > maxLens(c) Then maxLens(c) = text.Length
                Next
            Next

            For c = 0 To headers.Count - 1
                Dim chars As Integer = maxLens(c) + 2
                If chars < MIN_CHARS Then chars = MIN_CHARS
                If chars > MAX_CHARS Then chars = MAX_CHARS

                Dim width As Integer = chars * 256
                If width > MAX_WIDTH Then width = MAX_WIDTH

                Try
                    sh.SetColumnWidth(c, width)
                Catch
                End Try
            Next
        End Sub

        Private Shared Function StatusRank(status As String) As Integer
            If String.IsNullOrEmpty(status) Then Return 0
            If String.Equals(status, "Mismatch", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "Shared Parameter 등록 필요", StringComparison.OrdinalIgnoreCase) Then Return 3
            If String.Equals(status, "연결 필요(Proximity)", StringComparison.OrdinalIgnoreCase) Then Return 2
            If String.Equals(status, "연결 대상 객체 없음", StringComparison.OrdinalIgnoreCase) Then Return 2
            If String.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) Then Return 1
            Return 0
        End Function

        Private Shared Function ToDoubleLocal(val As String) As Double
            Try
                If String.IsNullOrWhiteSpace(val) Then Return Double.MaxValue
                Return Convert.ToDouble(val)
            Catch
                Return Double.MaxValue
            End Try
        End Function

        Private Shared Function ToIntLocal(val As String) As Integer
            Try
                If String.IsNullOrEmpty(val) Then Return 0
                Dim s = val.Trim()
                Return Convert.ToInt32(s)
            Catch
                Return 0
            End Try
        End Function

#End Region

#Region "공종검토 엑셀 기본 파일명 유틸(공통)"

        ' 요구사항:
        ' - RVT 파일명에서 prefix 추출:
        '   앞에서부터 3번째 "_"까지 + (그 다음 토큰에서 첫 "-숫자"까지)
        '   예) P4-2_FBELO_FFP_4F-0-김경연_2026-02-24  →  P4-2_FBELO_FFP_4F-0
        ' - suffix 고정: -06_공종검토_0차_
        ' - (n건) 은 "최종 엑셀에 실제로 저장되는 오류 행 수"
        ' - 규칙 불일치 시: "{원본파일명}_공종검토_(n건).xlsx"

        ' totalRows의 "File" 컬럼에서 RVT 파일명을 입력 순서대로 중복 제거하여 수집한다.
        Private Shared Function CollectDistinctRvtFilesInOrder(rows As List(Of Dictionary(Of String, Object))) As List(Of String)
            Dim result As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If rows Is Nothing Then Return result

            For Each r In rows
                Dim f As String = ReadFieldInsensitive(r, "File")
                If String.IsNullOrWhiteSpace(f) Then Continue For
                f = f.Trim()
                If seen.Add(f) Then result.Add(f)
            Next

            Return result
        End Function

        Private Shared Function BuildTradeReviewDefaultExcelName(rvtBaseName As String, issueCount As Integer) As String
            Dim n As Integer = If(issueCount < 0, 0, issueCount)

            Dim prefix As String = ExtractTradePrefix(rvtBaseName)
            Dim baseName As String

            If Not String.IsNullOrWhiteSpace(prefix) Then
                baseName = $"{prefix}-06_공종검토_0차_({n}건)"
            ElseIf Not String.IsNullOrWhiteSpace(rvtBaseName) Then
                baseName = $"{rvtBaseName}_공종검토_({n}건)"
            Else
                Return String.Empty
            End If

            baseName = SanitizeFileName(baseName)
            If String.IsNullOrWhiteSpace(baseName) Then baseName = "Export"

            If baseName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then
                Return baseName
            End If
            Return baseName & ".xlsx"
        End Function

        Private Shared Function ExtractTradePrefix(rvtBaseName As String) As String
            If String.IsNullOrWhiteSpace(rvtBaseName) Then Return String.Empty

            Dim parts As String() = rvtBaseName.Split("_"c)
            If parts Is Nothing OrElse parts.Length < 4 Then Return String.Empty

            Dim token3 As String = If(parts(3), "").Trim()
            If String.IsNullOrWhiteSpace(token3) Then Return String.Empty

            ' token3에서 "첫 -숫자"까지 추출 (예: "4F-0-김경연" → "4F-0", "7F-0-06" → "7F-0")
            Dim cut As Integer = -1
            For i As Integer = 0 To token3.Length - 2
                If token3(i) = "-"c AndAlso Char.IsDigit(token3(i + 1)) Then
                    Dim j As Integer = i + 1
                    While j < token3.Length AndAlso Char.IsDigit(token3(j))
                        j += 1
                    End While
                    cut = j ' end(exclusive)
                    Exit For
                End If
            Next

            Dim token3Prefix As String = If(cut > 0, token3.Substring(0, cut), token3)

            Return $"{parts(0)}_{parts(1)}_{parts(2)}_{token3Prefix}"
        End Function

        Private Shared Function SanitizeFileName(fileName As String) As String
            If String.IsNullOrWhiteSpace(fileName) Then Return String.Empty
            Dim s As String = fileName

            Try
                For Each ch In System.IO.Path.GetInvalidFileNameChars()
                    s = s.Replace(ch, "_"c)
                Next
            Catch
                ' ignore
            End Try

            s = s.Trim()
            ' Windows에서 끝의 점/공백은 문제를 만들 수 있음
            s = s.TrimEnd("."c)

            Return s
        End Function

#End Region



        ' 검토내용(F열) 셀에 대해 연속성 오류(Mismatch/BothEmpty)일 때만 배경색+빨간 글씨 서식 적용
        ' - 다른 상태(등록 필요/연결 대상 없음/ERROR/Match 등)는 서식 변경 없음
        Private Shared Sub ApplyConnectorReviewContentIssueStyles(xlsxPath As String)
            If String.IsNullOrWhiteSpace(xlsxPath) Then Return
            If Not File.Exists(xlsxPath) Then Return

            Dim wb As XSSFWorkbook = Nothing
            Try
                Using fs As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    wb = New XSSFWorkbook(fs)
                End Using

                ' 연한 노랑 배경 + 빨간 글씨
                Dim fillRgb As Byte() = New Byte() {255, 242, 204} ' #FFF2CC
                Dim fontRed As Short = IndexedColors.Red.Index

                For si As Integer = 0 To wb.NumberOfSheets - 1
                    Dim sh As ISheet = wb.GetSheetAt(si)
                    If sh Is Nothing Then Continue For

                    Dim header As IRow = sh.GetRow(0)
                    If header Is Nothing Then Continue For

                    Dim colReview As Integer = -1 ' 검토내용
                    Dim colPc As Integer = -1     ' ParamCompare

                    For c As Integer = 0 To header.LastCellNum - 1
                        Dim hc As ICell = header.GetCell(c)
                        Dim t As String = If(hc Is Nothing, "", hc.ToString()).Trim()

                        If colReview < 0 AndAlso String.Equals(t, "검토내용", StringComparison.OrdinalIgnoreCase) Then colReview = c
                        If colPc < 0 AndAlso String.Equals(t, "ParamCompare", StringComparison.OrdinalIgnoreCase) Then colPc = c
                    Next

                    If colReview < 0 Then colReview = 5 ' F열
                    If colPc < 0 Then colPc = header.LastCellNum - 1

                    ' base style: 첫 데이터행의 검토내용 셀 스타일을 기반으로 clone
                    Dim baseStyle As ICellStyle = Nothing
                    Dim r1 As IRow = sh.GetRow(1)
                    If r1 IsNot Nothing Then
                        Dim cReview As ICell = r1.GetCell(colReview)
                        If cReview IsNot Nothing Then baseStyle = cReview.CellStyle
                    End If
                    If baseStyle Is Nothing Then baseStyle = wb.CreateCellStyle()

                    ' style 1개만 만들어 재사용(스타일 폭증 방지)
                    Dim issueStyle As XSSFCellStyle = CType(wb.CreateCellStyle(), XSSFCellStyle)
                    issueStyle.CloneStyleFrom(baseStyle)
                    issueStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
                    issueStyle.SetFillForegroundColor(New XSSFColor(fillRgb))

                    Dim f As IFont = wb.CreateFont()
                    f.Color = fontRed
                    issueStyle.SetFont(f)

                    For r As Integer = 1 To sh.LastRowNum
                        Dim row As IRow = sh.GetRow(r)
                        If row Is Nothing Then Continue For

                        Dim pcCell As ICell = row.GetCell(colPc)
                        Dim pc As String = If(pcCell Is Nothing, "", pcCell.ToString()).Trim()

                        If pc.Equals("Mismatch", StringComparison.OrdinalIgnoreCase) OrElse pc.Equals("BothEmpty", StringComparison.OrdinalIgnoreCase) Then
                            Dim reviewCell As ICell = row.GetCell(colReview)
                            If reviewCell Is Nothing Then reviewCell = row.CreateCell(colReview)
                            reviewCell.CellStyle = issueStyle
                        End If
                    Next
                Next

                Using fsw As New FileStream(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.Read)
                    wb.Write(fsw)
                End Using
            Catch
                ' ignore
            Finally
                Try
                    If wb IsNot Nothing Then wb.Close()
                Catch
                End Try
            End Try
        End Sub


    End Class
End Namespace
