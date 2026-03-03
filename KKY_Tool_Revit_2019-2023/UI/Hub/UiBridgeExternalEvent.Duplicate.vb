Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Exports
Imports KKY_Tool_Revit.Infrastructure

' ✅ WPF 팝업(별칭 import로 Color 충돌 회피)
Imports WPF = System.Windows
Imports WControls = System.Windows.Controls
Imports WMedia = System.Windows.Media

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

#Region "상태 (세션/결과 보관)"

        ' 삭제 묶음 스택 (직전 삭제 → 되돌리기 대상)
        Private ReadOnly _deleteOps As New Stack(Of List(Of Integer))()

        ' 마지막 스캔 결과(엑셀 내보내기 및 UI 상태 동기화용)
        Private _lastRows As New List(Of DupRowDto)

        ' 마지막 실행 모드 ("duplicate" | "clash")
        Private _lastMode As String = "duplicate"

        ' 중첩(shared) 패밀리 인스턴스(서브컴포넌트) ID 집합
        Private Shared _nestedSharedIds As HashSet(Of Integer) = Nothing

        Private Class DupRowDto
            Public Property ElementId As Integer
            Public Property Category As String
            Public Property Family As String
            Public Property [Type] As String
            Public Property ConnectedCount As Integer
            Public Property ConnectedIds As String
            Public Property Candidate As Boolean
            Public Property Deleted As Boolean

            ' 상위호환: Host에서 그룹을 명시적으로 내려주면 Web/Excel 모두 동일 그룹핑 가능
            Public Property GroupKey As String

            ' 상위호환: mode를 row에도 내려(디버그/호환)
            Public Property Mode As String
        End Class

#End Region

#Region "핸들러"

        ' ====== 중복/자체간섭 스캔 ======
        Private Sub HandleDupRun(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("host:error", New With {.message = "활성 문서가 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document

            ' 중첩 Shared 컴포넌트 목록 캐시
            _nestedSharedIds = New HashSet(Of Integer)()
            Try
                Dim famCol As New FilteredElementCollector(doc)
                famCol.OfClass(GetType(FamilyInstance)).WhereElementIsNotElementType()
                For Each o As Element In famCol
                    Dim fi As FamilyInstance = TryCast(o, FamilyInstance)
                    If fi Is Nothing Then Continue For
                    Try
                        Dim subs = fi.GetSubComponentIds()
                        If subs Is Nothing Then Continue For
                        For Each sid As ElementId In subs
                            _nestedSharedIds.Add(sid.IntegerValue)
                        Next
                    Catch
                    End Try
                Next
            Catch
            End Try

            Dim mode As String = "duplicate"
            Try
                Dim mObj = GetProp(payload, "mode")
                If mObj IsNot Nothing Then
                    Dim mStr = Convert.ToString(mObj)
                    If Not String.IsNullOrWhiteSpace(mStr) Then
                        mStr = mStr.Trim().ToLowerInvariant()
                        If mStr = "duplicate" OrElse mStr = "clash" Then
                            mode = mStr
                        End If
                    End If
                End If
            Catch
            End Try

            Dim tolFeet As Double = 1.0R / 64.0R
            Try
                Dim tolObj = GetProp(payload, "tolFeet")
                If tolObj IsNot Nothing Then
                    tolFeet = Math.Max(0.000001R, Convert.ToDouble(tolObj))
                End If
            Catch
            End Try

            _lastMode = mode

            If mode = "clash" Then
                RunSelfClash(doc, tolFeet)
            Else
                RunDuplicate(doc, tolFeet)
            End If
        End Sub

        ' ====== 중복 스캔(기존) ======
        Private Sub RunDuplicate(doc As Document, tolFeet As Double)
            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0
            Dim groupsCount As Integer = 0
            Dim candidates As Integer = 0

            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()

            Dim q = Function(x As Double) As Long
                        Return CLng(Math.Round(x / tolFeet))
                    End Function

            Dim buckets As New Dictionary(Of String, List(Of ElementId))(StringComparer.Ordinal)

            Dim catCache As New Dictionary(Of Integer, String)()
            Dim famCache As New Dictionary(Of Integer, String)()
            Dim typCache As New Dictionary(Of Integer, String)()

            For Each e As Element In collector
                total += 1

                If ShouldSkipForQuantity(e) Then Continue For
                If e Is Nothing OrElse e.Category Is Nothing Then Continue For

                Dim center As XYZ = TryGetCenter(e)
                If center Is Nothing Then Continue For

                Dim catName As String = SafeCategoryName(e, catCache)
                Dim famName As String = SafeFamilyName(e, famCache)
                Dim typName As String = SafeTypeName(e, typCache)

                Dim lvl As Integer = TryGetLevelId(e)
                Dim oriKey As String = GetOrientationKey(e)

                Dim key As String =
                    String.Concat(catName, "|", famName, "|", typName, "|",
                                  "O", oriKey, "|", "L", lvl.ToString(), "|",
                                  "Q(", q(center.X).ToString(), ",", q(center.Y).ToString(), ",", q(center.Z).ToString(), ")")

                Dim list As List(Of ElementId) = Nothing
                If Not buckets.TryGetValue(key, list) Then
                    list = New List(Of ElementId)()
                    buckets.Add(key, list)
                End If
                list.Add(e.Id)
            Next

            Dim groupIndex As Integer = 0

            For Each kv In buckets
                Dim ids As List(Of ElementId) = kv.Value
                If ids.Count <= 1 Then Continue For

                groupsCount += 1
                groupIndex += 1

                Dim gk As String = "D" & groupIndex.ToString("D4")

                For Each id As ElementId In ids
                    Dim e As Element = doc.GetElement(id)
                    If e Is Nothing Then Continue For

                    Dim catName As String = SafeCategoryName(e, catCache)
                    Dim famName As String = SafeFamilyName(e, famCache)
                    Dim typName As String = SafeTypeName(e, typCache)

                    Dim connIds = ids _
                        .Where(Function(x) x.IntegerValue <> id.IntegerValue) _
                        .Select(Function(x) x.IntegerValue.ToString()) _
                        .ToArray()

                    rows.Add(New DupRowDto With {
                        .ElementId = id.IntegerValue,
                        .Category = catName,
                        .Family = famName,
                        .Type = typName,
                        .ConnectedCount = connIds.Length,
                        .ConnectedIds = String.Join(", ", connIds),
                        .Candidate = True,
                        .Deleted = False,
                        .GroupKey = gk,
                        .Mode = "duplicate"
                    })
                    candidates += 1
                Next
            Next

            _lastRows = rows

            Dim wireRows = rows.Select(Function(r) New With {
                .elementId = r.ElementId,
                .category = r.Category,
                .family = r.Family,
                .type = r.Type,
                .connectedCount = r.ConnectedCount,
                .connectedIds = r.ConnectedIds,
                .candidate = r.Candidate,
                .deleted = r.Deleted,
                .groupKey = r.GroupKey,
                .mode = r.Mode
            }).ToList()

            SendToWeb("dup:list", wireRows)
            SendToWeb("dup:result", New With {.mode = "duplicate", .scan = total, .groups = groupsCount, .candidates = candidates, .tolFeet = tolFeet})
        End Sub

        ' ====== 자체간섭(클래시) 스캔(상위호환) ======
        Private Sub RunSelfClash(doc As Document, tolFeet As Double)
            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0

            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()

            Dim catCache As New Dictionary(Of Integer, String)()
            Dim famCache As New Dictionary(Of Integer, String)()
            Dim typCache As New Dictionary(Of Integer, String)()

            ' 1) 대상 요소 + BoundingBox 수집
            Dim infos As New Dictionary(Of Integer, ClashInfo)()

            For Each e As Element In collector
                total += 1

                If ShouldSkipForQuantity(e) Then Continue For
                If e Is Nothing OrElse e.Category Is Nothing Then Continue For

                Dim bb As BoundingBoxXYZ = GetBoundingBox(e)
                If bb Is Nothing OrElse bb.Min Is Nothing OrElse bb.Max Is Nothing Then Continue For

                Dim minX As Double = bb.Min.X - tolFeet
                Dim minY As Double = bb.Min.Y - tolFeet
                Dim minZ As Double = bb.Min.Z - tolFeet
                Dim maxX As Double = bb.Max.X + tolFeet
                Dim maxY As Double = bb.Max.Y + tolFeet
                Dim maxZ As Double = bb.Max.Z + tolFeet

                ' invalid guard
                If maxX < minX OrElse maxY < minY OrElse maxZ < minZ Then Continue For

                Dim id As Integer = e.Id.IntegerValue
                Dim ci As New ClashInfo With {
                    .Id = id,
                    .MinX = minX, .MinY = minY, .MinZ = minZ,
                    .MaxX = maxX, .MaxY = maxY, .MaxZ = maxZ,
                    .Category = SafeCategoryName(e, catCache),
                    .Family = SafeFamilyName(e, famCache),
                    .TypeName = SafeTypeName(e, typCache)
                }
                infos(id) = ci
            Next

            ' 2) 공간 해시(2D)로 후보쌍 생성 → BB 교차 판정 → 그래프 구축
            Dim adjacency As New Dictionary(Of Integer, HashSet(Of Integer))()
            Dim pairSeen As New HashSet(Of Long)()

            ' cellSize: tol이 너무 작으면 셀 폭발 방지 위해 최소 0.5ft
            Dim cellSize As Double = Math.Max(0.5R, tolFeet * 32.0R)

            Dim cells As New Dictionary(Of Long, List(Of Integer))()

            For Each kv In infos
                Dim ci = kv.Value

                Dim ix0 As Integer = CInt(Math.Floor(ci.MinX / cellSize))
                Dim ix1 As Integer = CInt(Math.Floor(ci.MaxX / cellSize))
                Dim iy0 As Integer = CInt(Math.Floor(ci.MinY / cellSize))
                Dim iy1 As Integer = CInt(Math.Floor(ci.MaxY / cellSize))

                ' 큰 요소(너무 많은 셀 점유)는 center cell만 사용(성능 보호)
                Dim dx As Integer = ix1 - ix0
                Dim dy As Integer = iy1 - iy0
                If dx > 25 OrElse dy > 25 Then
                    Dim cx As Double = (ci.MinX + ci.MaxX) * 0.5R
                    Dim cy As Double = (ci.MinY + ci.MaxY) * 0.5R
                    ix0 = CInt(Math.Floor(cx / cellSize))
                    ix1 = ix0
                    iy0 = CInt(Math.Floor(cy / cellSize))
                    iy1 = iy0
                End If

                For ix As Integer = ix0 To ix1
                    For iy As Integer = iy0 To iy1
                        Dim key As Long = PackCell(ix, iy)
                        Dim lst As List(Of Integer) = Nothing
                        If Not cells.TryGetValue(key, lst) Then
                            lst = New List(Of Integer)()
                            cells.Add(key, lst)
                        End If
                        lst.Add(ci.Id)
                    Next
                Next
            Next

            For Each kv In cells
                Dim lst = kv.Value
                If lst Is Nothing OrElse lst.Count < 2 Then Continue For

                For i As Integer = 0 To lst.Count - 2
                    Dim aId As Integer = lst(i)
                    Dim a As ClashInfo = Nothing
                    If Not infos.TryGetValue(aId, a) Then Continue For

                    For j As Integer = i + 1 To lst.Count - 1
                        Dim bId As Integer = lst(j)
                        If aId = bId Then Continue For

                        Dim pairKey As Long = MakePairKey(aId, bId)
                        If pairSeen.Contains(pairKey) Then Continue For
                        pairSeen.Add(pairKey)

                        Dim b As ClashInfo = Nothing
                        If Not infos.TryGetValue(bId, b) Then Continue For

                        If BBoxIntersects(a, b) Then
                            AddEdge(adjacency, aId, bId)
                        End If
                    Next
                Next
            Next

            ' 3) Connected Components → 그룹
            Dim groups As New List(Of List(Of Integer))()
            Dim visited As New HashSet(Of Integer)()

            For Each id As Integer In adjacency.Keys.ToList()
                If visited.Contains(id) Then Continue For

                Dim comp As New List(Of Integer)()
                Dim q As New Queue(Of Integer)()
                q.Enqueue(id)
                visited.Add(id)

                While q.Count > 0
                    Dim cur As Integer = q.Dequeue()
                    comp.Add(cur)

                    Dim nbSet As HashSet(Of Integer) = Nothing
                    If adjacency.TryGetValue(cur, nbSet) AndAlso nbSet IsNot Nothing Then
                        For Each nb As Integer In nbSet
                            If Not visited.Contains(nb) Then
                                visited.Add(nb)
                                q.Enqueue(nb)
                            End If
                        Next
                    End If
                End While

                If comp.Count >= 2 Then
                    groups.Add(comp)
                End If
            Next

            ' 큰 그룹 우선 정렬
            groups = groups.OrderByDescending(Function(g) g.Count).ToList()

            Dim groupIndex As Integer = 0
            For Each g In groups
                groupIndex += 1
                Dim gk As String = "C" & groupIndex.ToString("D4")

                For Each id As Integer In g
                    Dim ci As ClashInfo = Nothing
                    If Not infos.TryGetValue(id, ci) Then Continue For

                    Dim nbSet As HashSet(Of Integer) = Nothing
                    Dim conn As String = ""
                    Dim connCnt As Integer = 0

                    If adjacency.TryGetValue(id, nbSet) AndAlso nbSet IsNot Nothing AndAlso nbSet.Count > 0 Then
                        Dim connArr = nbSet.OrderBy(Function(x) x).Select(Function(x) x.ToString()).ToArray()
                        connCnt = connArr.Length
                        conn = String.Join(", ", connArr)
                    End If

                    rows.Add(New DupRowDto With {
                        .ElementId = id,
                        .Category = ci.Category,
                        .Family = ci.Family,
                        .Type = ci.TypeName,
                        .ConnectedCount = connCnt,
                        .ConnectedIds = conn,
                        .Candidate = True,
                        .Deleted = False,
                        .GroupKey = gk,
                        .Mode = "clash"
                    })
                Next
            Next

            _lastRows = rows

            Dim wireRows = rows.Select(Function(r) New With {
                .elementId = r.ElementId,
                .category = r.Category,
                .family = r.Family,
                .type = r.Type,
                .connectedCount = r.ConnectedCount,
                .connectedIds = r.ConnectedIds,
                .candidate = r.Candidate,
                .deleted = r.Deleted,
                .groupKey = r.GroupKey,
                .mode = r.Mode
            }).ToList()

            SendToWeb("dup:list", wireRows)
            SendToWeb("dup:result", New With {.mode = "clash", .scan = total, .groups = groups.Count, .candidates = rows.Count, .tolFeet = tolFeet})
        End Sub

        ' ====== 선택/줌 ======
        Private Sub HandleDuplicateSelect(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing Then Return

            Dim idVal As Integer = SafeInt(GetProp(payload, "id"))
            If idVal <= 0 Then Return

            Dim elId As New ElementId(idVal)
            Dim el As Element = uiDoc.Document.GetElement(elId)
            If el Is Nothing Then
                SendToWeb("host:warn", New With {.message = $"요소 {idVal} 을(를) 찾을 수 없습니다."})
                Return
            End If

            Try
                uiDoc.Selection.SetElementIds(New List(Of ElementId) From {elId})
            Catch
            End Try

            Dim bb As BoundingBoxXYZ = GetBoundingBox(el)
            Try
                If bb IsNot Nothing Then
                    Dim views = uiDoc.GetOpenUIViews()
                    Dim target = views.FirstOrDefault(Function(v) v.ViewId.IntegerValue = uiDoc.ActiveView.Id.IntegerValue)
                    If target IsNot Nothing Then
                        target.ZoomAndCenterRectangle(bb.Min, bb.Max)
                    Else
                        uiDoc.ShowElements(elId)
                    End If
                Else
                    uiDoc.ShowElements(elId)
                End If
            Catch
            End Try
        End Sub

        ' ====== 삭제(트랜잭션 1회 커밋) ======
        Private Sub HandleDuplicateDelete(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document

            Dim ids As List(Of Integer) = ExtractIds(payload)
            If ids Is Nothing OrElse ids.Count = 0 Then
                SendToWeb("revit:error", New With {.message = "잘못된 요청입니다(id 누락/형식 오류)."})
                Return
            End If

            Dim eidList As New List(Of ElementId)
            For Each i In ids
                If i > 0 Then
                    Dim eid As New ElementId(i)
                    If doc.GetElement(eid) IsNot Nothing Then
                        eidList.Add(eid)
                    End If
                End If
            Next

            If eidList.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "삭제할 유효한 요소가 없습니다."})
                Return
            End If

            Dim actuallyDeleted As New List(Of Integer)

            Using t As New Transaction(doc, $"KKY Dup Delete ({eidList.Count})")
                t.Start()
                Try
                    doc.Delete(eidList)
                    t.Commit()
                Catch ex As Exception
                    t.RollBack()
                    SendToWeb("revit:error", New With {.message = $"삭제 실패({eidList.Count}개): {ex.Message}"})
                    Return
                End Try
            End Using

            For Each eid In eidList
                If doc.GetElement(eid) Is Nothing Then
                    actuallyDeleted.Add(eid.IntegerValue)
                    Dim row = _lastRows.FirstOrDefault(Function(r) r.ElementId = eid.IntegerValue)
                    If row IsNot Nothing Then
                        row.Deleted = True
                        SendToWeb("dup:deleted", New With {.id = eid.IntegerValue})
                    End If
                End If
            Next

            If actuallyDeleted.Count > 0 Then
                _deleteOps.Push(actuallyDeleted)
            End If
        End Sub

        ' ====== 되돌리기(직전 삭제 묶음 Undo) ======
        Private Sub HandleDuplicateRestore(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            If _deleteOps.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "되돌릴 수 있는 최신 삭제가 없습니다."})
                Return
            End If

            ' 요청으로 들어온 id (현재 UI는 단일 id 기준)
            Dim requestIds As List(Of Integer) = ExtractIds(payload)
            Dim lastPack As List(Of Integer) = _deleteOps.Peek()

            ' 요청 id 집합이 직전 삭제 묶음과 동일한지 확인
            Dim same As Boolean =
                requestIds IsNot Nothing AndAlso
                requestIds.Count = lastPack.Count AndAlso
                Not requestIds.Except(lastPack).Any()

            If Not same Then
                SendToWeb("host:warn", New With {.message = "되돌리기는 직전 삭제 묶음만 가능합니다."})
                Return
            End If

            Try
                ' Revit 공식 Undo 포스터블 커맨드 사용
                Dim cmdId As RevitCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)
                If cmdId Is Nothing Then Throw New InvalidOperationException("Undo 명령을 찾을 수 없습니다.")
                uiDoc.Application.PostCommand(cmdId)
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = $"되돌리기 실패: {ex.Message}"})
                Return
            End Try

            ' 스택에서 제거 후 상태/UI 동기화
            _deleteOps.Pop()

            For Each i In lastPack
                Dim r = _lastRows.FirstOrDefault(Function(x) x.ElementId = i)
                If r IsNot Nothing Then
                    r.Deleted = False
                    SendToWeb("dup:restored", New With {.id = i})
                End If
            Next
        End Sub

        ' ====== 엑셀 내보내기 ======
        Private Sub HandleDuplicateExport(app As UIApplication, Optional payload As Object = Nothing)
            If _lastRows Is Nothing Then
                SendToWeb("host:warn", New With {.message = "내보낼 데이터가 없습니다."})
                Return
            End If

            Dim token As String = TryCast(GetProp(payload, "token"), String)

            Try
                Dim groupsCount As Integer = CountGroups(_lastRows)

                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Dim todayToken As String = Date.Now.ToString("yyMMdd")

                Dim titleKor As String = If(String.Equals(_lastMode, "clash", StringComparison.OrdinalIgnoreCase), "자체간섭", "중복객체")
                Dim defaultFileName As String = $"{todayToken}_{titleKor} 검토결과_{groupsCount}개.xlsx"
                Dim defaultPath As String = Path.Combine(desktop, defaultFileName)

                Dim sfd As New Microsoft.Win32.SaveFileDialog() With {
                    .Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    .FileName = Path.GetFileName(defaultPath),
                    .AddExtension = True,
                    .DefaultExt = "xlsx",
                    .OverwritePrompt = True
                }

                If sfd.ShowDialog() <> True Then Exit Sub

                Dim outPath As String = sfd.FileName

                ' 엑셀 내보내기
                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("dup:progress")

                Dim sheetTitle As String = If(String.Equals(_lastMode, "clash", StringComparison.OrdinalIgnoreCase),
                                              "Self Clash (Simple)",
                                              "Duplicates (Simple)")

                Exports.DuplicateExport.Save(outPath, _lastRows.Cast(Of Object)(), doAutoFit, "dup:progress", sheetTitle)

                SendToWeb("dup:exported", New With {.path = outPath, .ok = True, .token = token})
            Catch ioEx As IOException
                Dim msg As String =
                    "해당 파일이 열려 있어 저장에 실패했습니다." & Environment.NewLine &
                    "엑셀에서 파일을 닫은 뒤 다시 시도해 주세요."
                SendToWeb("dup:exported", New With {.ok = False, .message = msg, .token = token})
            Catch ex As Exception
                SendToWeb("dup:exported", New With {.ok = False, .message = $"엑셀 내보내기에 실패했습니다: {ex.Message}", .token = token})
            End Try
        End Sub

#End Region

#Region "자체간섭 내부 구현"

        Private Structure ClashInfo
            Public Id As Integer
            Public MinX As Double
            Public MinY As Double
            Public MinZ As Double
            Public MaxX As Double
            Public MaxY As Double
            Public MaxZ As Double
            Public Category As String
            Public Family As String
            Public TypeName As String
        End Structure

        Private Shared Function BBoxIntersects(a As ClashInfo, b As ClashInfo) As Boolean
            If a.MaxX < b.MinX OrElse a.MinX > b.MaxX Then Return False
            If a.MaxY < b.MinY OrElse a.MinY > b.MaxY Then Return False
            If a.MaxZ < b.MinZ OrElse a.MinZ > b.MaxZ Then Return False
            Return True
        End Function

        Private Shared Sub AddEdge(adj As Dictionary(Of Integer, HashSet(Of Integer)), a As Integer, b As Integer)
            Dim sa As HashSet(Of Integer) = Nothing
            If Not adj.TryGetValue(a, sa) Then
                sa = New HashSet(Of Integer)()
                adj(a) = sa
            End If
            sa.Add(b)

            Dim sb As HashSet(Of Integer) = Nothing
            If Not adj.TryGetValue(b, sb) Then
                sb = New HashSet(Of Integer)()
                adj(b) = sb
            End If
            sb.Add(a)
        End Sub

        ' (ix,iy) → Long key
        Private Shared Function PackCell(ix As Integer, iy As Integer) As Long
            Dim x As Long = CLng(ix)
            Dim y As Long = CLng(iy) And &HFFFFFFFFL
            Return (x << 32) Or y
        End Function

        ' (minId,maxId) → Long key
        Private Shared Function MakePairKey(a As Integer, b As Integer) As Long
            Dim lo As Long = Math.Min(a, b)
            Dim hi As Long = Math.Max(a, b)
            Return (hi << 32) Or (lo And &HFFFFFFFFL)
        End Function

#End Region

#Region "필터/유틸(물량 필터 강화)"

        ' ⭐ 실제 시공 물량으로 보는 객체만 남기기 위한 필터
        ' ⭐ 모델링된 모든 요소 대상(주석/태그/참조/자동종속 제외)
        Private Shared Function ShouldSkipForQuantity(e As Element) As Boolean
            ' 0) 기본 예외: 널 / 임포트
            If e Is Nothing Then Return True
            If TypeOf e Is ImportInstance Then Return True

            ' 1) 뷰 전용(주석/태그/치수 등 대부분)
            Try
                If e.ViewSpecific Then Return True
            Catch
                ' 일부 요소는 ViewSpecific 접근 예외 가능 → 보수적으로 제외
                Return True
            End Try

            ' 2) 카테고리: Model만, 카테고리 없으면 제외
            Dim cat As Category = Nothing
            Try
                cat = e.Category
            Catch
                Return True
            End Try
            If cat Is Nothing Then Return True

            Try
                If cat.CategoryType <> CategoryType.Model Then Return True
            Catch
                Return True
            End Try

            ' 하위 카테고리(서브카테고리) 제외
            If cat.Parent IsNot Nothing Then Return True

            ' 3) 참조/기준/선류 제외(요구사항)
            If TypeOf e Is CurveElement Then Return True ' ModelLine/ReferenceLine/SketchLine 등
            If TypeOf e Is ReferencePlane Then Return True
            If TypeOf e Is Level Then Return True
            If TypeOf e Is Grid Then Return True
            If TypeOf e Is DatumPlane Then Return True

            ' 4) 합의 4가지:
            ' - Opening / ShaftOpening: 포함(여기서 제외하지 않음)
            ' - Rebar: 제외
            ' - Parts: 제외
            ' - Rooms/Spaces/Areas: 제외
            If TypeOf e Is Part Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Architecture.Room Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.Space Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Area Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.Rebar Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.AreaReinforcement Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.PathReinforcement Then Return True

            ' 5) 자동 종속(호스트 의존)만 제외: Insulation/Lining
            If TypeOf e Is Autodesk.Revit.DB.Plumbing.PipeInsulation Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.DuctInsulation Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.DuctLining Then Return True

            ' 6) 중첩 패밀리(패밀리 안 패밀리) → 상위만 남기고 서브는 제외
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing Then
                Try
                    If fi.SuperComponent IsNot Nothing Then Return True
                    If _nestedSharedIds IsNot Nothing AndAlso _nestedSharedIds.Contains(fi.Id.IntegerValue) Then Return True
                Catch
                End Try
            End If

            ' ✅ Solid(Volume) 검사 제거: 커넥터/솔리드 유무 무관, 모델 요소면 대상
            Return False
        End Function

        Private Shared Function QOri(x As Double) As Long
            Return CLng(Math.Round(x * 1000.0R))
        End Function

        Private Shared Function GetOrientationKey(e As Element) As String
            Try
                Dim fi = TryCast(e, FamilyInstance)
                If fi IsNot Nothing Then
                    Dim mirrored As Boolean = False
                    Dim hand As Boolean = False
                    Dim facing As Boolean = False
                    Try : mirrored = fi.Mirrored : Catch : End Try
                    Try : hand = fi.HandFlipped : Catch : End Try
                    Try : facing = fi.FacingFlipped : Catch : End Try

                    Dim t As Transform = Nothing
                    Try : t = fi.GetTransform() : Catch : End Try

                    Dim keyParts As New List(Of String)()
                    keyParts.Add("M" & If(mirrored, "1", "0"))
                    keyParts.Add("H" & If(hand, "1", "0"))
                    keyParts.Add("F" & If(facing, "1", "0"))

                    If t IsNot Nothing Then
                        Dim ox = t.BasisX
                        Dim oy = t.BasisY
                        Dim oz = t.BasisZ
                        keyParts.Add("OX(" & QOri(ox.X) & "," & QOri(ox.Y) & "," & QOri(ox.Z) & ")")
                        keyParts.Add("OY(" & QOri(oy.X) & "," & QOri(oy.Y) & "," & QOri(oy.Z) & ")")
                        keyParts.Add("OZ(" & QOri(oz.X) & "," & QOri(oz.Y) & "," & QOri(oz.Z) & ")")
                    End If

                    Return String.Join("|", keyParts)
                End If

                Dim loc As Location = Nothing
                Try : loc = e.Location : Catch : End Try

                Dim lc = TryCast(loc, LocationCurve)
                If lc IsNot Nothing AndAlso lc.Curve IsNot Nothing Then
                    Dim c = lc.Curve
                    Dim dir As XYZ = Nothing
                    Try : dir = (c.GetEndPoint(1) - c.GetEndPoint(0)) : Catch : End Try
                    If dir IsNot Nothing Then
                        Dim len As Double = dir.GetLength()
                        If len > 0.000001R Then dir = dir / len
                        Return "LC(" & QOri(dir.X) & "," & QOri(dir.Y) & "," & QOri(dir.Z) & ")"
                    End If
                End If

            Catch
            End Try

            Return String.Empty
        End Function

        Private Shared Function SafeCategoryName(e As Element, cache As Dictionary(Of Integer, String)) As String
            If e Is Nothing OrElse e.Category Is Nothing Then Return ""
            Dim id As Integer = e.Category.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = e.Category.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeFamilyName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi Is Nothing OrElse fi.Symbol Is Nothing OrElse fi.Symbol.Family Is Nothing Then Return ""
            Dim id As Integer = fi.Symbol.Family.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = fi.Symbol.Family.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeTypeName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then
                Dim id As Integer = fi.Symbol.Id.IntegerValue
                Dim s As String = Nothing
                If cache.TryGetValue(id, s) Then Return s
                s = fi.Symbol.Name
                cache(id) = s
                Return s
            End If
            Return e.Name
        End Function

        Private Shared Function TryGetLevelId(e As Element) As Integer
            Try
                Dim p As Parameter = e.Parameter(BuiltInParameter.LEVEL_PARAM)
                If p IsNot Nothing Then
                    Dim lvid As ElementId = p.AsElementId()
                    If lvid IsNot Nothing AndAlso lvid <> ElementId.InvalidElementId Then
                        Return lvid.IntegerValue
                    End If
                End If
            Catch
            End Try

            Try
                Dim pi = e.GetType().GetProperty("LevelId")
                If pi IsNot Nothing Then
                    Dim id = TryCast(pi.GetValue(e, Nothing), ElementId)
                    If id IsNot Nothing AndAlso id <> ElementId.InvalidElementId Then
                        Return id.IntegerValue
                    End If
                End If
            Catch
            End Try

            Return -1
        End Function

        Private Shared Function TryGetCenter(e As Element) As XYZ
            If e Is Nothing Then Return Nothing

            Try
                Dim loc As Location = e.Location
                If TypeOf loc Is LocationPoint Then
                    Return CType(loc, LocationPoint).Point
                ElseIf TypeOf loc Is LocationCurve Then
                    Dim crv = CType(loc, LocationCurve).Curve
                    If crv IsNot Nothing Then Return crv.Evaluate(0.5, True)
                End If
            Catch
            End Try

            Dim bb = GetBoundingBox(e)
            If bb IsNot Nothing Then
                Return (bb.Min + bb.Max) * 0.5R
            End If

            Return Nothing
        End Function

        Private Shared Function GetBoundingBox(e As Element) As BoundingBoxXYZ
            Try
                Dim bb As BoundingBoxXYZ = e.BoundingBox(Nothing)
                If bb IsNot Nothing Then Return bb
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function SafeInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                Return Convert.ToInt32(o)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ExtractIds(payload As Object) As List(Of Integer)
            Dim result As New List(Of Integer)()

            Dim singleObj = GetProp(payload, "id")
            Dim v As Integer = SafeToInt(singleObj)
            If v > 0 Then
                result.Add(v)
                Return result
            End If

            Dim arr = GetProp(payload, "ids")
            If arr Is Nothing Then Return result

            Dim enumerable = TryCast(arr, System.Collections.IEnumerable)
            If enumerable IsNot Nothing Then
                For Each o In enumerable
                    Dim iv = SafeToInt(o)
                    If iv > 0 Then result.Add(iv)
                Next
            End If

            Return result
        End Function

        Private Shared Function SafeToInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                If TypeOf o Is Integer Then Return CInt(o)
                If TypeOf o Is Long Then Return CInt(CLng(o))
                If TypeOf o Is Double Then Return CInt(CDbl(o))
                If TypeOf o Is String Then
                    Dim s As String = CStr(o)
                    Dim iv As Integer
                    If Integer.TryParse(s, iv) Then Return iv
                End If
            Catch
            End Try
            Return 0
        End Function

        ' ✅ groupKey가 있으면 그걸로 그룹 수 산정, 없으면 기존 connectedIds 기반 산정
        Private Function CountGroups(rows As IEnumerable(Of DupRowDto)) As Integer
            If rows Is Nothing Then Return 0

            Dim gkCount As Integer =
                rows.Select(Function(r) If(r.GroupKey, "")) _
                    .Where(Function(s) Not String.IsNullOrWhiteSpace(s)) _
                    .Distinct(StringComparer.Ordinal) _
                    .Count()

            If gkCount > 0 Then Return gkCount

            ' fallback: 기존 로직(connectedIds cluster)
            Dim bucket As New HashSet(Of String)(StringComparer.Ordinal)

            For Each r In rows
                Dim id As String = r.ElementId.ToString()
                Dim cat As String = If(r.Category, "")
                Dim fam As String = If(r.Family, "")
                Dim typ As String = If(r.Type, "")
                Dim conStr As String = If(r.ConnectedIds, "")

                Dim cluster As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(id) Then cluster.Add(id)
                cluster.AddRange(SplitIds(conStr))

                Dim norm = cluster _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Select(Function(x) x.Trim()) _
                    .Distinct() _
                    .OrderBy(Function(x) x) _
                    .ToList()

                Dim clusterKey As String = If(norm.Count > 1, String.Join(",", norm), "")
                Dim famOut As String = If(String.IsNullOrWhiteSpace(fam), If(String.IsNullOrWhiteSpace(cat), "", cat & " Type"), fam)
                Dim key = String.Join("|", {cat, famOut, typ, clusterKey})
                bucket.Add(key)
            Next

            Return bucket.Count
        End Function

        Private Function SplitIds(s As String) As IEnumerable(Of String)
            If String.IsNullOrWhiteSpace(s) Then Return Array.Empty(Of String)()
            Return s.Split(New Char() {","c, " "c, ";"c, "|"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        End Function

#End Region

#Region "WPF 팝업 (실패 시 TaskDialog 폴백)"

        Private Function IsSystemDark() As Boolean
            Try
                Dim key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim v = key.GetValue("AppsUseLightTheme", 1)
                    Return (Convert.ToInt32(v) = 0)
                End If
            Catch
            End Try
            Return False
        End Function

        ''' <summary>
        ''' 엑셀 내보내기 안내를 WPF로 시도하고 실패하면 TaskDialog로 폴백한다.
        ''' 반환값: True = 파일 열기
        ''' </summary>
        Private Function ShowExcelSavedDialog(outPath As String,
                                             groupsCount As Integer,
                                             Optional chipLabel As String = "중복 그룹",
                                             Optional dialogTitle As String = "중복검토 내보내기",
                                             Optional headerText As String = "엑셀로 저장했습니다.",
                                             Optional questionText As String = "지금 파일을 열어보시겠어요?") As Boolean
            Try
                Dim win As New WPF.Window()
                win.Title = "KKY Tool_Revit - " & dialogTitle
                win.SizeToContent = WPF.SizeToContent.WidthAndHeight
                win.WindowStartupLocation = WPF.WindowStartupLocation.CenterOwner
                win.ResizeMode = WPF.ResizeMode.NoResize
                win.Topmost = True

                win.Content = BuildExcelSavedContent(outPath, groupsCount, chipLabel, headerText, questionText, win)

                ' Owner 연결(가능하면 Revit 메인 윈도우에 붙이기)
                Try
                    Dim t = Type.GetType("Autodesk.Windows.ComponentManager, AdWindows")
                    If t IsNot Nothing Then
                        Dim p = t.GetProperty("ApplicationWindow", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static)
                        If p IsNot Nothing Then
                            Dim hwnd = CType(p.GetValue(Nothing, Nothing), IntPtr)
                            Dim helper = New WPF.Interop.WindowInteropHelper(win)
                            helper.Owner = hwnd
                        End If
                    End If
                Catch
                End Try

                Dim res As Boolean? = win.ShowDialog()
                Return If(res.HasValue AndAlso res.Value, True, False)

            Catch
                ' 폴백: TaskDialog
                Dim td As New TaskDialog(dialogTitle)
                td.MainIcon = TaskDialogIcon.TaskDialogIconInformation
                td.MainInstruction = headerText
                td.MainContent = $"{chipLabel}: {groupsCount}개{Environment.NewLine}파일: {outPath}"
                td.CommonButtons = TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No
                td.DefaultButton = TaskDialogResult.Yes
                td.FooterText = questionText

                Dim r = td.Show()
                Return r = TaskDialogResult.Yes
            End Try
        End Function

        Private Function BuildExcelSavedContent(outPath As String,
                                               groupsCount As Integer,
                                               chipLabel As String,
                                               headerText As String,
                                               questionText As String,
                                               host As WPF.Window) As WPF.UIElement
            Dim isDark = IsSystemDark()

            ' 테마별 색상 정의 (Byte 캐스팅으로 Option Strict 대응)
            Dim bgPanel As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H12), CByte(&H16), CByte(&H1C)), WMedia.Color.FromRgb(CByte(&HFF), CByte(&HFF), CByte(&HFF)))
            Dim bgCard As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H18), CByte(&H1C), CByte(&H24)), WMedia.Color.FromRgb(CByte(&HF7), CByte(&HF8), CByte(&HFA)))
            Dim headG1 As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H1F), CByte(&H5A), CByte(&HFF)), WMedia.Color.FromRgb(CByte(&H66), CByte(&H99), CByte(&HFF)))
            Dim headG2 As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H78), CByte(&H9B), CByte(&HFF)), WMedia.Color.FromRgb(CByte(&H9F), CByte(&HBE), CByte(&HFF)))
            Dim fgMain As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&HE8), CByte(&HEA), CByte(&HED)), WMedia.Color.FromRgb(CByte(&H11), CByte(&H11), CByte(&H11)))
            Dim fgSub As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&HC7), CByte(&HC9), CByte(&HCC)), WMedia.Color.FromRgb(CByte(&H55), CByte(&H55), CByte(&H55)))
            Dim chipBg As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H21), CByte(&H26), CByte(&H32)), WMedia.Color.FromRgb(CByte(&HEE), CByte(&HF1), CByte(&HF5)))
            Dim accent As WMedia.Color = If(isDark, WMedia.Color.FromRgb(CByte(&H7A), CByte(&HA2), CByte(&HFF)), WMedia.Color.FromRgb(CByte(&H38), CByte(&H67), CByte(&HFF)))
            Dim bdLine As WMedia.Color = If(isDark, WMedia.Color.FromArgb(CByte(&H33), CByte(&HFF), CByte(&HFF), CByte(&HFF)), WMedia.Color.FromArgb(CByte(&H22), CByte(&H0), CByte(&H0), CByte(&H0)))

            Dim root As New WControls.Border() With {
                .Background = New WMedia.SolidColorBrush(bgPanel),
                .Padding = New WPF.Thickness(16)
            }

            Dim card As New WControls.Border() With {
                .Padding = New WPF.Thickness(0),
                .CornerRadius = New WPF.CornerRadius(14),
                .BorderThickness = New WPF.Thickness(1),
                .BorderBrush = New WMedia.SolidColorBrush(bdLine),
                .Background = New WMedia.SolidColorBrush(bgCard),
                .Effect = New WMedia.Effects.DropShadowEffect() With {.Opacity = 0.25, .BlurRadius = 16, .ShadowDepth = 0}
            }

            Dim wrap As New WControls.StackPanel() With {.Width = 560}

            ' 헤더 (그라데이션 바)
            Dim header As New WControls.Border() With {
                .CornerRadius = New WPF.CornerRadius(14, 14, 0, 0),
                .Background = New WMedia.LinearGradientBrush(headG1, headG2, 0),
                .Padding = New WPF.Thickness(20, 14, 20, 16)
            }

            Dim hTitle As New WControls.TextBlock() With {
                .Text = headerText,
                .FontSize = 18,
                .FontWeight = WPF.FontWeights.SemiBold,
                .Foreground = WMedia.Brushes.White
            }
            header.Child = hTitle

            ' 바디 패딩
            Dim bodyPad As New WControls.Border() With {
                .Padding = New WPF.Thickness(20),
                .Background = New WMedia.SolidColorBrush(bgCard),
                .CornerRadius = New WPF.CornerRadius(0, 0, 14, 14)
            }

            Dim body As New WControls.StackPanel() With {.Orientation = WControls.Orientation.Vertical}

            ' 그룹 수 칩
            Dim chip As New WControls.Border() With {
                .CornerRadius = New WPF.CornerRadius(999),
                .Background = New WMedia.SolidColorBrush(chipBg),
                .Padding = New WPF.Thickness(12, 6, 12, 6),
                .Margin = New WPF.Thickness(0, 8, 0, 10)
            }

            Dim chipText As New WControls.TextBlock() With {
                .Text = $"{chipLabel} {groupsCount}개",
                .Foreground = New WMedia.SolidColorBrush(accent),
                .FontWeight = WPF.FontWeights.SemiBold
            }
            chip.Child = chipText

            ' 파일 경로
            Dim pathTb As New WControls.TextBlock() With {
                .Text = $"파일: {outPath}",
                .TextWrapping = WPF.TextWrapping.Wrap,
                .Foreground = New WMedia.SolidColorBrush(fgSub),
                .Margin = New WPF.Thickness(0, 0, 0, 14)
            }

            ' 질문 텍스트
            Dim question As New WControls.TextBlock() With {
                .Text = questionText,
                .Foreground = New WMedia.SolidColorBrush(fgMain),
                .Margin = New WPF.Thickness(0, 4, 0, 10)
            }

            ' 버튼 바
            Dim btnBar As New WControls.StackPanel() With {
                .Orientation = WControls.Orientation.Horizontal,
                .HorizontalAlignment = WPF.HorizontalAlignment.Right
            }

            Dim yesBtn As New WControls.Button() With {
                .Content = "예(Y)",
                .MinWidth = 88,
                .Padding = New WPF.Thickness(14, 7, 14, 7),
                .Margin = New WPF.Thickness(0, 0, 8, 0),
                .Foreground = WMedia.Brushes.White,
                .Background = New WMedia.SolidColorBrush(accent),
                .BorderBrush = WMedia.Brushes.Transparent
            }

            Dim noBtn As New WControls.Button() With {
                .Content = "아니오(N)",
                .MinWidth = 88,
                .Padding = New WPF.Thickness(14, 7, 14, 7),
                .Foreground = New WMedia.SolidColorBrush(fgMain),
                .Background = New WMedia.SolidColorBrush(chipBg),
                .BorderBrush = WMedia.Brushes.Transparent
            }

            AddHandler yesBtn.Click, Sub(sender As Object, e As WPF.RoutedEventArgs)
                                         host.DialogResult = True
                                         host.Close()
                                     End Sub

            AddHandler noBtn.Click, Sub(sender As Object, e As WPF.RoutedEventArgs)
                                        host.DialogResult = False
                                        host.Close()
                                    End Sub

            btnBar.Children.Add(yesBtn)
            btnBar.Children.Add(noBtn)

            body.Children.Add(chip)
            body.Children.Add(pathTb)
            body.Children.Add(question)
            body.Children.Add(btnBar)

            bodyPad.Child = body
            wrap.Children.Add(header)
            wrap.Children.Add(bodyPad)
            card.Child = wrap
            root.Child = card

            Return root
        End Function

#End Region

    End Class

End Namespace