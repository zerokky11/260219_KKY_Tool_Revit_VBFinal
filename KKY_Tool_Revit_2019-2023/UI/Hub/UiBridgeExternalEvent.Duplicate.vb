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

        Private Const MAX_UI_ROWS As Integer = 5000

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

            ' 상위호환: 그룹키(Host가 만들어 내려줌)
            Public Property GroupKey As String

            ' 상위호환: mode
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

                        ' ✅ subcomponent 목록(있으면) 추가
                        If subs IsNot Nothing Then
                            For Each sid As ElementId In subs
                                _nestedSharedIds.Add(sid.IntegerValue)
                            Next
                        End If

                        ' ✅ 최상위 패밀리만 표시: nested(하위) 패밀리 인스턴스는 제외용 집합에 추가
                        Try
                            If fi.SuperComponent IsNot Nothing Then
                                _nestedSharedIds.Add(fi.Id.IntegerValue)
                            End If
                        Catch
                        End Try
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

            ' ✅ Navisworks Set처럼: 선택집합을 범위/제외로 사용
            Dim scopeMode As String = "all"
            Try
                Dim smObj = GetProp(payload, "scopeMode")
                If smObj IsNot Nothing Then
                    Dim s = Convert.ToString(smObj)
                    If Not String.IsNullOrWhiteSpace(s) Then
                        s = s.Trim().ToLowerInvariant()
                        If s = "all" OrElse s = "scope" OrElse s = "exclude" Then scopeMode = s
                    End If
                End If
            Catch
            End Try

            Dim selectedIds As HashSet(Of Integer) = Nothing
            Try
                Dim sel = uiDoc.Selection.GetElementIds()
                If sel IsNot Nothing AndAlso sel.Count > 0 Then
                    selectedIds = New HashSet(Of Integer)(sel.Select(Function(x) x.IntegerValue))
                End If
            Catch
            End Try

            Dim scopeIds As HashSet(Of Integer) = Nothing
            Dim excludeIds As HashSet(Of Integer) = Nothing
            If selectedIds IsNot Nothing AndAlso selectedIds.Count > 0 Then
                If scopeMode = "scope" Then
                    scopeIds = selectedIds
                ElseIf scopeMode = "exclude" Then
                    excludeIds = selectedIds
                End If
            End If

            ' ✅ 제외 키워드(콤마): Family/Type/Category/Name에 포함되면 결과에서 제외
            Dim excludeKeywords As List(Of String) = Nothing
            Try
                Dim kwObj = GetProp(payload, "excludeKeywords")
                If kwObj IsNot Nothing Then
                    excludeKeywords = New List(Of String)()
                    Dim en = TryCast(kwObj, System.Collections.IEnumerable)
                    If en IsNot Nothing AndAlso Not TypeOf kwObj Is String Then
                        For Each o In en
                            If o Is Nothing Then Continue For
                            Dim s = o.ToString().Trim()
                            If s.Length > 0 Then excludeKeywords.Add(s.ToLowerInvariant())
                        Next
                    Else
                        Dim s = Convert.ToString(kwObj)
                        If Not String.IsNullOrWhiteSpace(s) Then
                            For Each part In s.Split(","c)
                                Dim p = part.Trim()
                                If p.Length > 0 Then excludeKeywords.Add(p.ToLowerInvariant())
                            Next
                        End If
                    End If
                    If excludeKeywords.Count = 0 Then excludeKeywords = Nothing
                End If
            Catch
                excludeKeywords = Nothing
            End Try

            ' ✅ 메타(드롭다운 목록)만 요청하는 경우: 스캔 없이 목록만 전송
            Dim metaOnly As Boolean = False
            Try
                Dim mo = GetProp(payload, "metaOnly")
                If mo IsNot Nothing Then
                    Dim s = Convert.ToString(mo).Trim().ToLowerInvariant()
                    If s = "true" OrElse s = "1" OrElse s = "yes" Then metaOnly = True
                End If
            Catch
            End Try

            Dim ruleConfig As Object = Nothing
            Try
                ruleConfig = GetProp(payload, "ruleConfig")
            Catch
                ruleConfig = Nothing
            End Try

            If metaOnly Then
                Try
                    SendDupMeta(doc)
                Catch
                End Try
                SendToWeb("dup:result", New With {.mode = mode, .scan = 0, .groups = 0, .candidates = 0, .tolFeet = tolFeet, .shown = 0, .total = 0, .truncated = False})
                Return
            End If

            _lastMode = mode

            If mode = "clash" Then
                RunSelfClash(doc, tolFeet, scopeIds, excludeIds, excludeKeywords, ruleConfig)
            Else
                RunDuplicate(doc, tolFeet, scopeIds, excludeIds, excludeKeywords, ruleConfig)
            End If
        End Sub

        ' ====== 중복 스캔(개선: Curve는 EndPoint+Size 기반) ======
        Private Sub RunDuplicate(doc As Document, tolFeet As Double, Optional scopeIds As HashSet(Of Integer) = Nothing, Optional excludeIds As HashSet(Of Integer) = Nothing, Optional excludeKeywords As List(Of String) = Nothing, Optional ruleConfig As Object = Nothing)
            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0
            Dim groupsCount As Integer = 0

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
                ' scope/exclude/keyword 필터
                If scopeIds IsNot Nothing AndAlso scopeIds.Count > 0 Then
                    Try
                        If e Is Nothing OrElse Not scopeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                        Continue For
                    End Try
                End If

                If excludeIds IsNot Nothing AndAlso excludeIds.Count > 0 Then
                    Try
                        If e IsNot Nothing AndAlso excludeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                    End Try
                End If

                If excludeKeywords IsNot Nothing AndAlso excludeKeywords.Count > 0 Then
                    If ShouldExcludeByKeywords(e, excludeKeywords) Then Continue For
                End If

                If excludeIds IsNot Nothing AndAlso excludeIds.Count > 0 Then
                    Try
                        If e IsNot Nothing AndAlso excludeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                    End Try
                End If

                If excludeKeywords IsNot Nothing AndAlso excludeKeywords.Count > 0 Then
                    If ShouldExcludeByKeywords(e, excludeKeywords) Then Continue For
                End If

                If scopeIds IsNot Nothing AndAlso scopeIds.Count > 0 Then
                    Try
                        If e Is Nothing OrElse Not scopeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                        Continue For
                    End Try
                End If
                total += 1

                If ShouldSkipForQuantity(e) Then Continue For
                If e Is Nothing OrElse e.Category Is Nothing Then Continue For

                Dim catName As String = SafeCategoryName(e, catCache)
                Dim famName As String = SafeFamilyName(e, famCache)
                Dim typName As String = SafeTypeName(e, typCache)
                Dim lvl As Integer = TryGetLevelId(e)

                Dim key As String = Nothing

                If TryBuildCurveDuplicateKey(e, tolFeet, catName, famName, typName, lvl, key) Then
                    ' key ok
                Else
                    Dim center As XYZ = TryGetCenter(e)
                    If center Is Nothing Then Continue For

                    Dim oriKey As String = GetOrientationKey(e)
                    key =
                    String.Concat(catName, "|", famName, "|", typName, "|",
                                  "O", oriKey, "|", "L", lvl.ToString(), "|",
                                  "Q(", q(center.X).ToString(), ",", q(center.Y).ToString(), ",", q(center.Z).ToString(), ")")
                End If

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

                    Dim connIds = ids.Where(Function(x) x.IntegerValue <> id.IntegerValue).
                    Select(Function(x) x.IntegerValue.ToString()).
                    ToArray()

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
                Next
            Next

            _lastRows = rows

            ' UI 표시 제한
            Dim truncated As Boolean = False
            Dim shownRows As List(Of DupRowDto) = rows
            If rows.Count > MAX_UI_ROWS Then
                truncated = True
                shownRows = rows.Take(MAX_UI_ROWS).ToList()
                SendToWeb("host:warn", New With {.message = $"결과가 {rows.Count}건으로 매우 많아 상위 {MAX_UI_ROWS}건만 표시합니다. 전체는 엑셀 내보내기에서 확인하세요."})
            End If

            Dim wireRows = shownRows.Select(Function(r) New With {
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
            SendToWeb("dup:result", New With {
            .mode = "duplicate",
            .scan = total,
            .groups = groupsCount,
            .candidates = rows.Count,
            .tolFeet = tolFeet,
            .shown = shownRows.Count,
            .total = rows.Count,
            .truncated = truncated
        })
        End Sub

        ' ====== 자체간섭(클래시) 스캔(후보 누락 방지 + 정밀판정) ======
        Private Sub RunSelfClash(doc As Document, tolFeet As Double, Optional scopeIds As HashSet(Of Integer) = Nothing, Optional excludeIds As HashSet(Of Integer) = Nothing, Optional excludeKeywords As List(Of String) = Nothing, Optional ruleConfig As Object = Nothing)
            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0

            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()

            Dim catCache As New Dictionary(Of Integer, String)()
            Dim famCache As New Dictionary(Of Integer, String)()
            Dim typCache As New Dictionary(Of Integer, String)()
            ' ✅ Rule/Set config (Navis-like)
            Dim cfg As ClashRuleConfig = ParseRuleConfig(ruleConfig)
            Dim setIndex As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim setDefs As List(Of ClashSetDef) = cfg.Sets

            For i As Integer = 0 To setDefs.Count - 1
                If setDefs(i) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(setDefs(i).Id) Then
                    If Not setIndex.ContainsKey(setDefs(i).Id) Then setIndex.Add(setDefs(i).Id, i)
                End If
            Next

            Dim excludeMask As ULong = 0UL
            For Each sid As String In cfg.ExcludeSetIds
                Dim si As Integer = -1
                If Not String.IsNullOrWhiteSpace(sid) AndAlso setIndex.TryGetValue(sid, si) Then
                    If si >= 0 AndAlso si < 64 Then excludeMask = excludeMask Or (1UL << si)
                End If
            Next

            Dim pairRules As New List(Of Tuple(Of Integer, Integer, Boolean, Boolean))() ' (aIdx, bIdx, aAll, bAll)
            For Each pr As ClashPairRule In cfg.Pairs
                If pr Is Nothing OrElse Not pr.Enabled Then Continue For

                Dim aId As String = pr.A
                Dim bId As String = pr.B
                Dim aAll As Boolean = String.Equals(aId, "__ALL__", StringComparison.OrdinalIgnoreCase)
                Dim bAll As Boolean = String.Equals(bId, "__ALL__", StringComparison.OrdinalIgnoreCase)

                Dim ai As Integer = -1
                Dim bi As Integer = -1
                If Not aAll Then
                    If Not setIndex.TryGetValue(aId, ai) Then Continue For
                End If
                If Not bAll Then
                    If Not setIndex.TryGetValue(bId, bi) Then Continue For
                End If
                pairRules.Add(Tuple.Create(ai, bi, aAll, bAll))
            Next

            Dim hasPairRules As Boolean = (pairRules.Count > 0)


            ' 1) 대상 요소 + BoundingBox 수집
            Dim infos As New Dictionary(Of Integer, ClashInfo)()



            ' ✅ shape 기반(솔리드) 교차 판정용
            Dim optGeom As New Options() With {.ComputeReferences = False, .IncludeNonVisibleObjects = False, .DetailLevel = ViewDetailLevel.Fine}
            Dim solidCache As New Dictionary(Of Integer, List(Of Solid))()
            For Each e As Element In collector
                ' scope/exclude/keyword 필터
                If scopeIds IsNot Nothing AndAlso scopeIds.Count > 0 Then
                    Try
                        If e Is Nothing OrElse Not scopeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                        Continue For
                    End Try
                End If

                If excludeIds IsNot Nothing AndAlso excludeIds.Count > 0 Then
                    Try
                        If e IsNot Nothing AndAlso excludeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                    End Try
                End If

                If excludeKeywords IsNot Nothing AndAlso excludeKeywords.Count > 0 Then
                    If ShouldExcludeByKeywords(e, excludeKeywords) Then Continue For
                End If

                If excludeIds IsNot Nothing AndAlso excludeIds.Count > 0 Then
                    Try
                        If e IsNot Nothing AndAlso excludeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                    End Try
                End If

                If excludeKeywords IsNot Nothing AndAlso excludeKeywords.Count > 0 Then
                    If ShouldExcludeByKeywords(e, excludeKeywords) Then Continue For
                End If

                If scopeIds IsNot Nothing AndAlso scopeIds.Count > 0 Then
                    Try
                        If e Is Nothing OrElse Not scopeIds.Contains(e.Id.IntegerValue) Then Continue For
                    Catch
                        Continue For
                    End Try
                End If
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

                If maxX < minX OrElse maxY < minY OrElse maxZ < minZ Then Continue For

                Dim id As Integer = e.Id.IntegerValue

                Dim ci As New ClashInfo With {
                .Id = id,
                .MinX = minX, .MinY = minY, .MinZ = minZ,
                .MaxX = maxX, .MaxY = maxY, .MaxZ = maxZ,
                .Category = SafeCategoryName(e, catCache),
                .Family = SafeFamilyName(e, famCache),
                .TypeName = SafeTypeName(e, typCache),
                .TypeIdInt = TryGetTypeIdInt(e),
                .SizeKey = GetMEPSizeKey(e),
                .RadiusHint = 0R,
                .HasCurve = False,
                .Radius = 0R
            }

                ci.RadiusHint = ComputeRadiusHint(ci.MinX, ci.MinY, ci.MinZ, ci.MaxX, ci.MaxY, ci.MaxZ, tolFeet)

                Dim p0 As XYZ = Nothing, p1 As XYZ = Nothing
                If TryGetCurveEndpoints(e, p0, p1) Then
                    ci.HasCurve = True
                    ci.X0 = p0.X : ci.Y0 = p0.Y : ci.Z0 = p0.Z
                    ci.X1 = p1.X : ci.Y1 = p1.Y : ci.Z1 = p1.Z

                    Dim r As Double = 0R
                    If TryGetCrossSectionRadius(e, r) Then
                        ci.Radius = Math.Max(0R, r)
                    End If
                End If

                ' set membership mask (max 64 sets)
                Dim mask As ULong = 0UL
                If setDefs IsNot Nothing AndAlso setDefs.Count > 0 Then
                    Dim si As Integer = 0
                    For Each sd As ClashSetDef In setDefs
                        If sd IsNot Nothing Then
                            Try
                                If ElementMatchesSet(e, sd) Then
                                    If si < 64 Then mask = mask Or (1UL << si)
                                End If
                            Catch
                            End Try
                        End If
                        si += 1
                        If si >= 64 Then Exit For
                    Next
                End If
                ci.Mask = mask

                infos(id) = ci
            Next

            ' 2) 공간 해시(2D)로 후보쌍 생성
            Dim adjacency As New Dictionary(Of Integer, HashSet(Of Integer))()
            Dim pairSeen As New HashSet(Of Long)()

            Dim cellSize As Double = Math.Max(0.5R, tolFeet * 32.0R)
            Dim cells As New Dictionary(Of Long, List(Of Integer))()

            Const MAX_CELLS_PER_ELEMENT As Integer = 2500
            Dim hugeIds As New List(Of Integer)()

            For Each kv In infos
                Dim ci = kv.Value

                Dim ix0 As Integer = CInt(Math.Floor(ci.MinX / cellSize))
                Dim ix1 As Integer = CInt(Math.Floor(ci.MaxX / cellSize))
                Dim iy0 As Integer = CInt(Math.Floor(ci.MinY / cellSize))
                Dim iy1 As Integer = CInt(Math.Floor(ci.MaxY / cellSize))

                Dim dx As Integer = ix1 - ix0
                Dim dy As Integer = iy1 - iy0

                Dim cellCount As Long = CLng(dx + 1) * CLng(dy + 1)

                ' ✅ 셀 커버리지가 너무 크면 hugeIds로 분리(후보 누락 방지: huge vs all은 별도 처리)
                If cellCount > MAX_CELLS_PER_ELEMENT Then
                    hugeIds.Add(ci.Id)
                    Continue For
                End If

                For ix As Integer = ix0 To ix1
                    For iy As Integer = iy0 To iy1
                        AddCell(cells, ix, iy, ci.Id)
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

                        If Not BBoxIntersects(a, b) Then Continue For

                        If IsRealClash(doc, a, b, tolFeet, optGeom, solidCache) Then
                            AddEdge(adjacency, aId, bId)
                        End If
                    Next
                Next
            Next



            ' ✅ huge 요소는 셀 분해 대신, 모든 요소와 AABB/정밀판정(후보 누락 방지)
            If hugeIds.Count > 0 Then
                Dim allIds = infos.Keys.ToList()

                For Each aId As Integer In hugeIds
                    Dim a As ClashInfo = Nothing
                    If Not infos.TryGetValue(aId, a) Then Continue For

                    For Each bId As Integer In allIds
                        If aId = bId Then Continue For

                        Dim pairKey As Long = MakePairKey(aId, bId)
                        If pairSeen.Contains(pairKey) Then Continue For
                        pairSeen.Add(pairKey)

                        Dim b As ClashInfo = Nothing
                        If Not infos.TryGetValue(bId, b) Then Continue For

                        If Not BBoxIntersects(a, b) Then Continue For

                        If IsRealClash(doc, a, b, tolFeet, optGeom, solidCache) Then
                            AddEdge(adjacency, aId, bId)
                        End If
                    Next
                Next
            End If
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

            Dim truncated As Boolean = False
            Dim shownRows As List(Of DupRowDto) = rows
            If rows.Count > MAX_UI_ROWS Then
                truncated = True
                shownRows = rows.Take(MAX_UI_ROWS).ToList()
                SendToWeb("host:warn", New With {.message = $"결과가 {rows.Count}건으로 매우 많아 상위 {MAX_UI_ROWS}건만 표시합니다. 전체는 엑셀 내보내기에서 확인하세요."})
            End If

            Dim wireRows = shownRows.Select(Function(r) New With {
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
            SendToWeb("dup:result", New With {
            .mode = "clash",
            .scan = total,
            .groups = groups.Count,
            .candidates = rows.Count,
            .tolFeet = tolFeet,
            .shown = shownRows.Count,
            .total = rows.Count,
            .truncated = truncated
        })
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

            Dim requestIds As List(Of Integer) = ExtractIds(payload)
            Dim lastPack As List(Of Integer) = _deleteOps.Peek()

            Dim same As Boolean =
            requestIds IsNot Nothing AndAlso
            requestIds.Count = lastPack.Count AndAlso
            Not requestIds.Except(lastPack).Any()

            If Not same Then
                SendToWeb("host:warn", New With {.message = "되돌리기는 직전 삭제 묶음만 가능합니다."})
                Return
            End If

            Try
                Dim cmdId As RevitCommandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)
                If cmdId Is Nothing Then Throw New InvalidOperationException("Undo 명령을 찾을 수 없습니다.")
                uiDoc.Application.PostCommand(cmdId)
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = $"되돌리기 실패: {ex.Message}"})
                Return
            End Try

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

                Dim doAutoFit As Boolean = ParseExcelMode(payload)
                Global.KKY_Tool_Revit.UI.Hub.ExcelProgressReporter.Reset("dup:progress")

                Dim sheetTitle As String = If(String.Equals(_lastMode, "clash", StringComparison.OrdinalIgnoreCase),
                                          "Self Clash (Refined)",
                                          "Duplicates (Refined)")

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

        ' ===========================
        ' Rule / Set (Navis-like) v1
        ' ===========================
        Private Class ClashClause
            Public Property Field As String   ' category|family|type|name|param
            Public Property [Op] As String    ' contains|equals|startswith|endswith|notcontains|notequals
            Public Property Value As String
            Public Property Param As String   ' for param field
        End Class

        Private Class ClashGroup
            Public Property Clauses As New List(Of ClashClause)() ' AND
        End Class

        Private Class ClashSetDef
            Public Property Id As String
            Public Property Name As String
            Public Property Groups As New List(Of ClashGroup)() ' OR
        End Class

        Private Class ClashPairRule
            Public Property A As String
            Public Property B As String
            Public Property Enabled As Boolean = True
        End Class

        Private Class ClashRuleConfig
            Public Property Sets As New List(Of ClashSetDef)()
            Public Property Pairs As New List(Of ClashPairRule)()
            Public Property ExcludeSetIds As New List(Of String)()
        End Class

        Private Shared Function GetAnyProp(obj As Object, name As String) As Object
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing
            Try
                Dim d = TryCast(obj, System.Collections.IDictionary)
                If d IsNot Nothing AndAlso d.Contains(name) Then Return d(name)
            Catch
            End Try
            Try
                Dim t = obj.GetType()
                Dim pi = t.GetProperty(name)
                If pi IsNot Nothing Then Return pi.GetValue(obj, Nothing)
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function ToLowerSafe(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Return s.Trim().ToLowerInvariant()
        End Function

        Private Shared Function AnyStr(obj As Object) As String
            If obj Is Nothing Then Return ""
            Try
                Return Convert.ToString(obj)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function ParseRuleConfig(ruleConfigObj As Object) As ClashRuleConfig
            Dim cfg As New ClashRuleConfig()
            If ruleConfigObj Is Nothing Then Return cfg

            ' sets
            Try
                Dim setsObj = GetAnyProp(ruleConfigObj, "sets")
                Dim en = TryCast(setsObj, System.Collections.IEnumerable)
                If en IsNot Nothing AndAlso Not TypeOf setsObj Is String Then
                    For Each so In en
                        If so Is Nothing Then Continue For
                        Dim sd As New ClashSetDef()
                        sd.Id = AnyStr(GetAnyProp(so, "id"))
                        sd.Name = AnyStr(GetAnyProp(so, "name"))
                        If String.IsNullOrWhiteSpace(sd.Id) Then sd.Id = Guid.NewGuid().ToString("N")

                        Dim groupsObj = GetAnyProp(so, "groups")
                        Dim gen = TryCast(groupsObj, System.Collections.IEnumerable)
                        If gen IsNot Nothing AndAlso Not TypeOf groupsObj Is String Then
                            For Each go In gen
                                If go Is Nothing Then Continue For
                                Dim g As New ClashGroup()

                                Dim clausesObj = GetAnyProp(go, "clauses")
                                Dim cen = TryCast(clausesObj, System.Collections.IEnumerable)
                                If cen IsNot Nothing AndAlso Not TypeOf clausesObj Is String Then
                                    For Each co In cen
                                        If co Is Nothing Then Continue For
                                        Dim c As New ClashClause()
                                        c.Field = ToLowerSafe(AnyStr(GetAnyProp(co, "field")))
                                        c.Op = ToLowerSafe(AnyStr(GetAnyProp(co, "op")))
                                        c.Value = AnyStr(GetAnyProp(co, "value"))
                                        c.Param = AnyStr(GetAnyProp(co, "param"))
                                        g.Clauses.Add(c)
                                    Next
                                End If

                                If g.Clauses.Count > 0 Then sd.Groups.Add(g)
                            Next
                        End If

                        If sd.Groups.Count > 0 Then cfg.Sets.Add(sd)
                    Next
                End If
            Catch
            End Try

            ' pairs
            Try
                Dim pairsObj = GetAnyProp(ruleConfigObj, "pairs")
                Dim en = TryCast(pairsObj, System.Collections.IEnumerable)
                If en IsNot Nothing AndAlso Not TypeOf pairsObj Is String Then
                    For Each po In en
                        If po Is Nothing Then Continue For
                        Dim pr As New ClashPairRule()
                        pr.A = AnyStr(GetAnyProp(po, "a"))
                        pr.B = AnyStr(GetAnyProp(po, "b"))
                        Dim enb = GetAnyProp(po, "enabled")
                        If enb IsNot Nothing Then
                            Dim s = AnyStr(enb).Trim().ToLowerInvariant()
                            pr.Enabled = Not (s = "false" OrElse s = "0" OrElse s = "no")
                        End If
                        If Not String.IsNullOrWhiteSpace(pr.A) AndAlso Not String.IsNullOrWhiteSpace(pr.B) Then
                            cfg.Pairs.Add(pr)
                        End If
                    Next
                End If
            Catch
            End Try

            ' exclude sets
            Try
                Dim exObj = GetAnyProp(ruleConfigObj, "excludeSetIds")
                Dim en = TryCast(exObj, System.Collections.IEnumerable)
                If en IsNot Nothing AndAlso Not TypeOf exObj Is String Then
                    For Each xo In en
                        Dim s = AnyStr(xo)
                        If Not String.IsNullOrWhiteSpace(s) Then cfg.ExcludeSetIds.Add(s)
                    Next
                End If
            Catch
            End Try

            Return cfg
        End Function

        Private Shared Function GetElementStringForField(e As Element, field As String, paramName As String) As String
            If e Is Nothing Then Return ""
            field = ToLowerSafe(field)

            Try
                If field = "category" Then
                    If e.Category IsNot Nothing Then Return AnyStr(e.Category.Name)
                    Return ""
                End If

                If field = "family" Then
                    Dim fi As FamilyInstance = TryCast(e, FamilyInstance)
                    If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing AndAlso fi.Symbol.Family IsNot Nothing Then
                        Return AnyStr(fi.Symbol.Family.Name)
                    End If
                    Return ""
                End If

                If field = "type" Then
                    Dim fi As FamilyInstance = TryCast(e, FamilyInstance)
                    If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then
                        Return AnyStr(fi.Symbol.Name)
                    End If
                    Return AnyStr(e.Name)
                End If

                If field = "name" Then
                    Return AnyStr(e.Name)
                End If

                If field = "param" OrElse field = "parameter" Then
                    If String.IsNullOrWhiteSpace(paramName) Then Return ""
                    Dim p As Parameter = Nothing
                    Try : p = e.LookupParameter(paramName) : Catch : p = Nothing : End Try
                    If p Is Nothing Then Return ""

                    Try
                        Select Case p.StorageType
                            Case StorageType.String
                                Dim s = p.AsString()
                                If s IsNot Nothing Then Return s
                            Case StorageType.Double
                                Dim vs = p.AsValueString()
                                If Not String.IsNullOrWhiteSpace(vs) Then Return vs
                                Return p.AsDouble().ToString()
                            Case StorageType.Integer
                                Dim vs = p.AsValueString()
                                If Not String.IsNullOrWhiteSpace(vs) Then Return vs
                                Return p.AsInteger().ToString()
                            Case StorageType.ElementId
                                Dim id = p.AsElementId()
                                If id IsNot Nothing Then Return id.IntegerValue.ToString()
                        End Select
                    Catch
                    End Try

                    Return ""
                End If
            Catch
            End Try

            Return ""
        End Function

        Private Shared Function EvalClause(val As String, c As ClashClause) As Boolean
            Dim left = ToLowerSafe(val)
            Dim right = ToLowerSafe(If(c Is Nothing, "", c.Value))
            Dim op = ToLowerSafe(If(c Is Nothing, "", c.Op))

            If op = "" OrElse op = "contains" Then
                Return (right = "" OrElse left.Contains(right))
            ElseIf op = "equals" OrElse op = "equal" Then
                Return left = right
            ElseIf op = "startswith" Then
                Return (right = "" OrElse left.StartsWith(right))
            ElseIf op = "endswith" Then
                Return (right = "" OrElse left.EndsWith(right))
            ElseIf op = "notcontains" Then
                If right = "" Then Return True
                Return Not left.Contains(right)
            ElseIf op = "notequals" OrElse op = "notequal" Then
                Return left <> right
            End If

            Return (right = "" OrElse left.Contains(right))
        End Function

        Private Shared Function ElementMatchesGroup(e As Element, g As ClashGroup) As Boolean
            If g Is Nothing OrElse g.Clauses Is Nothing OrElse g.Clauses.Count = 0 Then Return False
            For Each c In g.Clauses
                Dim v = GetElementStringForField(e, c.Field, c.Param)
                If Not EvalClause(v, c) Then Return False
            Next
            Return True
        End Function

        Private Shared Function ElementMatchesSet(e As Element, sdef As ClashSetDef) As Boolean
            If sdef Is Nothing OrElse sdef.Groups Is Nothing OrElse sdef.Groups.Count = 0 Then Return False
            For Each g In sdef.Groups
                If ElementMatchesGroup(e, g) Then Return True
            Next
            Return False
        End Function

        Private Sub SendDupMeta(doc As Document)
            If doc Is Nothing Then Return

            Dim cats As New List(Of String)()
            Try
                For Each c As Category In doc.Settings.Categories
                    If c Is Nothing Then Continue For
                    Try
                        If c.Parent IsNot Nothing Then Continue For
                        If c.CategoryType <> CategoryType.Model Then Continue For
                        cats.Add(c.Name)
                    Catch
                    End Try
                Next
            Catch
            End Try
            cats = cats.Distinct().OrderBy(Function(x) x).ToList()

            Dim fams As New List(Of String)()
            Try
                Dim fc As New FilteredElementCollector(doc)
                fc.OfClass(GetType(Family))
                For Each f As Family In fc
                    If f Is Nothing Then Continue For
                    Try
                        If Not String.IsNullOrWhiteSpace(f.Name) Then fams.Add(f.Name)
                    Catch
                    End Try
                Next
            Catch
            End Try
            fams = fams.Distinct().OrderBy(Function(x) x).ToList()

            Dim types As New List(Of String)()
            Try
                Dim tc As New FilteredElementCollector(doc)
                tc.WhereElementIsElementType()
                For Each e As Element In tc
                    If e Is Nothing Then Continue For
                    Try
                        If Not String.IsNullOrWhiteSpace(e.Name) Then types.Add(e.Name)
                    Catch
                    End Try
                Next
            Catch
            End Try
            types = types.Distinct().OrderBy(Function(x) x).ToList()

            Dim pars As New List(Of String)()
            Try
                Dim bm As BindingMap = doc.ParameterBindings
                Dim it As DefinitionBindingMapIterator = bm.ForwardIterator()
                it.Reset()
                While it.MoveNext()
                    Dim def As Definition = it.Key
                    If def Is Nothing Then Continue While
                    If Not String.IsNullOrWhiteSpace(def.Name) Then pars.Add(def.Name)
                End While
            Catch
            End Try
            pars = pars.Distinct().OrderBy(Function(x) x).ToList()

            SendToWeb("dup:meta", New With {.categories = cats, .families = fams, .types = types, .parameters = pars})
        End Sub


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

            Public TypeIdInt As Integer
            Public SizeKey As String
            Public Mask As ULong
            Public RadiusHint As Double

            Public HasCurve As Boolean
            Public X0 As Double
            Public Y0 As Double
            Public Z0 As Double
            Public X1 As Double
            Public Y1 As Double
            Public Z1 As Double
            Public Radius As Double
        End Structure

        Private Shared Sub AddCell(cells As Dictionary(Of Long, List(Of Integer)), ix As Integer, iy As Integer, id As Integer)
            Dim key As Long = PackCell(ix, iy)
            Dim lst As List(Of Integer) = Nothing
            If Not cells.TryGetValue(key, lst) Then
                lst = New List(Of Integer)()
                cells.Add(key, lst)
            End If
            lst.Add(id)
        End Sub

        Private Shared Sub AddCellByPoint(cells As Dictionary(Of Long, List(Of Integer)), id As Integer, x As Double, y As Double, cellSize As Double)
            Dim ix As Integer = CInt(Math.Floor(x / cellSize))
            Dim iy As Integer = CInt(Math.Floor(y / cellSize))
            AddCell(cells, ix, iy, id)
        End Sub

        Private Shared Function BBoxIntersects(a As ClashInfo, b As ClashInfo) As Boolean
            If a.MaxX < b.MinX OrElse a.MinX > b.MaxX Then Return False
            If a.MaxY < b.MinY OrElse a.MinY > b.MaxY Then Return False
            If a.MaxZ < b.MinZ OrElse a.MinZ > b.MaxZ Then Return False
            Return True
        End Function



        ' ✅ Revit 버전/VB 바인딩 차이 대응: get_Geometry / GetGeometry 둘 다 리플렉션으로 호출
        Private Shared Function GetGeometryCompat(e As Element, opt As Options) As GeometryElement
            If e Is Nothing Then Return Nothing
            If opt Is Nothing Then opt = New Options()

            Try
                Dim t = e.GetType()
                Dim mi = t.GetMethod("get_Geometry",
                             Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public,
                             Nothing,
                             New Type() {GetType(Options)},
                             Nothing)
                If mi Is Nothing Then
                    mi = t.GetMethod("GetGeometry",
                             Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public,
                             Nothing,
                             New Type() {GetType(Options)},
                             Nothing)
                End If
                If mi Is Nothing Then Return Nothing

                Dim ge = mi.Invoke(e, New Object() {opt})
                Return TryCast(ge, GeometryElement)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub CollectSolidsFromGeom(geom As GeometryElement, xform As Transform, acc As List(Of Solid))
            If geom Is Nothing OrElse acc Is Nothing Then Return

            For Each go As GeometryObject In geom
                If go Is Nothing Then Continue For

                Dim s As Solid = TryCast(go, Solid)
                If s IsNot Nothing Then
                    Try
                        If s.Volume > 0.000000001 Then
                            If xform IsNot Nothing AndAlso Not xform.IsIdentity Then
                                Try
                                    Dim ts = SolidUtils.CreateTransformed(s, xform)
                                    If ts IsNot Nothing AndAlso ts.Volume > 0.000000001 Then
                                        acc.Add(ts)
                                    Else
                                        acc.Add(s)
                                    End If
                                Catch
                                    acc.Add(s)
                                End Try
                            Else
                                acc.Add(s)
                            End If
                        End If
                    Catch
                    End Try
                    Continue For
                End If

                Dim gi As GeometryInstance = TryCast(go, GeometryInstance)
                If gi IsNot Nothing Then
                    Try
                        Dim instXf As Transform = gi.Transform
                        Dim nextXf As Transform = xform
                        If nextXf Is Nothing Then
                            nextXf = instXf
                        Else
                            nextXf = nextXf.Multiply(instXf)
                        End If

                        Dim instGeom As GeometryElement = Nothing
                        Try
                            instGeom = gi.GetInstanceGeometry()
                        Catch
                        End Try
                        If instGeom IsNot Nothing Then
                            CollectSolidsFromGeom(instGeom, nextXf, acc)
                        End If
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Shared Function GetSolidsCached(doc As Document, id As Integer, opt As Options, cache As Dictionary(Of Integer, List(Of Solid))) As List(Of Solid)
            If cache Is Nothing Then cache = New Dictionary(Of Integer, List(Of Solid))()

            Dim lst As List(Of Solid) = Nothing
            If cache.TryGetValue(id, lst) Then Return lst

            lst = New List(Of Solid)()
            Try
                Dim e As Element = doc.GetElement(New ElementId(id))
                If e IsNot Nothing Then
                    Dim ge As GeometryElement = GetGeometryCompat(e, opt)
                    CollectSolidsFromGeom(ge, Nothing, lst)
                End If
            Catch
            End Try

            cache(id) = lst
            Return lst
        End Function

        Private Shared Function SolidsIntersect(doc As Document,
                                       aId As Integer,
                                       bId As Integer,
                                       opt As Options,
                                       cache As Dictionary(Of Integer, List(Of Solid)),
                                       tolFeet As Double) As Boolean
            Dim sa = GetSolidsCached(doc, aId, opt, cache)
            Dim sb = GetSolidsCached(doc, bId, opt, cache)
            If sa Is Nothing OrElse sb Is Nothing OrElse sa.Count = 0 OrElse sb.Count = 0 Then Return False

            Dim tolVol As Double = Math.Max(0.0000000001, tolFeet * tolFeet * tolFeet)

            For Each s1 In sa
                If s1 Is Nothing Then Continue For
                For Each s2 In sb
                    If s2 Is Nothing Then Continue For
                    Try
                        Dim inter As Solid = BooleanOperationsUtils.ExecuteBooleanOperation(s1, s2, BooleanOperationsType.Intersect)
                        If inter IsNot Nothing Then
                            Dim v As Double = 0R
                            Try : v = inter.Volume : Catch : v = 0R : End Try
                            If v > tolVol Then Return True
                        End If
                    Catch
                        ' boolean 실패 시 false로 처리(오탐 감소). 필요시 필터 결과만으로 결정.
                    End Try
                Next
            Next

            Return False
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


        Private Shared Function TryGetTypeIdInt(e As Element) As Integer
            If e Is Nothing Then Return -1
            Try
                Dim tid As ElementId = e.GetTypeId()
                If tid IsNot Nothing AndAlso tid <> ElementId.InvalidElementId Then
                    Return tid.IntegerValue
                End If
            Catch
            End Try
            Return -1
        End Function

        Private Shared Function ComputeRadiusHint(minX As Double, minY As Double, minZ As Double,
                                         maxX As Double, maxY As Double, maxZ As Double,
                                         tolFeet As Double) As Double
            ' bbox는 이미 tolFeet만큼 확장된 값이 들어오므로, 원래 치수에 가깝게 2*tolFeet 보정
            Dim dx As Double = Math.Max(0R, (maxX - minX) - (2.0R * tolFeet))
            Dim dy As Double = Math.Max(0R, (maxY - minY) - (2.0R * tolFeet))
            Dim dz As Double = Math.Max(0R, (maxZ - minZ) - (2.0R * tolFeet))

            Dim arr As Double() = New Double() {dx, dy, dz}
            Array.Sort(arr)

            ' 두 개의 작은 변을 단면(가로/세로)로 보고, 대각선 반지름(=0.5*sqrt(w^2+h^2))로 근사
            Dim s1 As Double = arr(0)
            Dim s2 As Double = arr(1)

            If s1 <= 0R AndAlso s2 <= 0R Then Return 0R
            Return 0.5R * Math.Sqrt((s1 * s1) + (s2 * s2))
        End Function

        Private Shared Function DistXYZ(a As XYZ, b As XYZ) As Double
            If a Is Nothing OrElse b Is Nothing Then Return Double.MaxValue
            Dim dx As Double = a.X - b.X
            Dim dy As Double = a.Y - b.Y
            Dim dz As Double = a.Z - b.Z
            Return Math.Sqrt(dx * dx + dy * dy + dz * dz)
        End Function

        ' ✅ 완전 중복(동일 기하/동일 타입/동일 사이즈/동일 끝점)은 간섭에서 제외
        Private Shared Function IsExactDuplicateCurvePair(a As ClashInfo, b As ClashInfo, tolFeet As Double) As Boolean
            If Not a.HasCurve OrElse Not b.HasCurve Then Return False

            ' 타입이 다르면 동일 중복으로 보지 않음(사이즈 다르거나 타입 다른 경우는 간섭으로 남겨야 함)
            If a.TypeIdInt <= 0 OrElse b.TypeIdInt <= 0 OrElse a.TypeIdInt <> b.TypeIdInt Then Return False

            ' 사이즈키가 있으면 우선 사용(정확)
            Dim sizeOk As Boolean = False
            If Not String.IsNullOrWhiteSpace(a.SizeKey) AndAlso Not String.IsNullOrWhiteSpace(b.SizeKey) Then
                sizeOk = String.Equals(a.SizeKey, b.SizeKey, StringComparison.Ordinal)
            Else
                ' 사이즈키가 없으면 파라미터 기반 Radius가 둘 다 있어야만 비교
                If a.Radius > 0R AndAlso b.Radius > 0R Then
                    Dim radTol As Double = Math.Max(tolFeet, 0.0005R) ' 최소 약 0.15mm
                    sizeOk = Math.Abs(a.Radius - b.Radius) <= radTol
                End If
            End If
            If Not sizeOk Then Return False

            Dim endTol As Double = Math.Max(tolFeet, 0.0005R)

            Dim a0 As New XYZ(a.X0, a.Y0, a.Z0)
            Dim a1 As New XYZ(a.X1, a.Y1, a.Z1)
            Dim b0 As New XYZ(b.X0, b.Y0, b.Z0)
            Dim b1 As New XYZ(b.X1, b.Y1, b.Z1)

            Dim sameOrder As Boolean = (DistXYZ(a0, b0) <= endTol AndAlso DistXYZ(a1, b1) <= endTol)
            Dim swapped As Boolean = (DistXYZ(a0, b1) <= endTol AndAlso DistXYZ(a1, b0) <= endTol)

            Return sameOrder OrElse swapped
        End Function

        ' ✅ Non-curve 완전 중복(동일 타입 + bbox 동일)은 간섭에서 제외
        Private Shared Function IsExactBBoxDuplicate(a As ClashInfo, b As ClashInfo, tolFeet As Double) As Boolean
            If a.TypeIdInt <= 0 OrElse b.TypeIdInt <= 0 OrElse a.TypeIdInt <> b.TypeIdInt Then Return False

            Dim t As Double = Math.Max(tolFeet, 0.0005R)

            If Math.Abs(a.MinX - b.MinX) > t Then Return False
            If Math.Abs(a.MinY - b.MinY) > t Then Return False
            If Math.Abs(a.MinZ - b.MinZ) > t Then Return False
            If Math.Abs(a.MaxX - b.MaxX) > t Then Return False
            If Math.Abs(a.MaxY - b.MaxY) > t Then Return False
            If Math.Abs(a.MaxZ - b.MaxZ) > t Then Return False

            Return True
        End Function

        Private Shared Function IsRealClash(doc As Document, a As ClashInfo, b As ClashInfo, tolFeet As Double, opt As Options, solidCache As Dictionary(Of Integer, List(Of Solid))) As Boolean
            ' ✅ 완전 중복은 간섭에서 제외
            If IsExactBBoxDuplicate(a, b, tolFeet) Then Return False

            ' 1) curve-curve: centerline distance + overlap for near-parallel
            If a.HasCurve AndAlso b.HasCurve Then
                Dim p0 As New XYZ(a.X0, a.Y0, a.Z0)
                Dim p1 As New XYZ(a.X1, a.Y1, a.Z1)
                Dim q0 As New XYZ(b.X0, b.Y0, b.Z0)
                Dim q1 As New XYZ(b.X1, b.Y1, b.Z1)

                Dim rA As Double = If(a.Radius > 0R, a.Radius, Math.Max(0R, a.RadiusHint))
                Dim rB As Double = If(b.Radius > 0R, b.Radius, Math.Max(0R, b.RadiusHint))
                Dim radSum As Double = rA + rB + tolFeet

                Dim dist As Double = SegmentDistance(p0, p1, q0, q1)
                If dist > radSum Then Return False

                Dim vA As XYZ = p1 - p0
                Dim vB As XYZ = q1 - q0
                Dim lenA As Double = vA.GetLength()
                Dim lenB As Double = vB.GetLength()

                If lenA > 0.000001R AndAlso lenB > 0.000001R Then
                    Dim uA As XYZ = vA / lenA
                    Dim uB As XYZ = vB / lenB
                    Dim dot As Double = Math.Abs(uA.DotProduct(uB))

                    If dot >= 0.996R Then
                        Dim overlap As Double = SegmentOverlapLengthAlongAxis(p0, p1, q0, q1, uA, lenA)
                        Dim overlapTol As Double = Math.Max(tolFeet * 2.0R, 0.001R) ' 0.001ft ≈ 0.3mm
                        If overlap < overlapTol Then Return False
                    End If
                End If

                ' ✅ 완전 중복(동일 기하/동일 사이즈/동일 타입)은 간섭에서 제외
                If IsExactDuplicateCurvePair(a, b, tolFeet) Then
                    Return False
                End If

                Return True
            End If

            ' 2) mixed/solid: Revit native 교차 판정 우선

            Dim ea As Element = Nothing
            Dim eb As Element = Nothing
            Try
                ea = doc.GetElement(New ElementId(a.Id))
                eb = doc.GetElement(New ElementId(b.Id))
            Catch
            End Try

            If ea Is Nothing OrElse eb Is Nothing Then Return False

            ' ✅ 2) mixed/solid: Revit 교차 필터(빠른 후보) → 솔리드 교차(볼륨>0)로 확정
            Dim maybe As Boolean = False
            Try
                Dim f As New ElementIntersectsElementFilter(eb)
                maybe = f.PassesFilter(doc, ea.Id)
            Catch
                maybe = False
            End Try

            If Not maybe Then
                ' 장비 패밀리 오탐 방지: 필터가 false면 간섭 아님
                Return False
            End If

            ' 솔리드 볼륨 교차로 확정 (Navis Hard 성향)
            Try
                If SolidsIntersect(doc, a.Id, b.Id, opt, solidCache, tolFeet) Then Return True
            Catch
            End Try

            Return False
        End Function

        Private Shared Function SegmentDistance(p0 As XYZ, p1 As XYZ, q0 As XYZ, q1 As XYZ) As Double
            Dim d2 As Double = SegmentDistanceSquared(p0, p1, q0, q1)
            If d2 <= 0R Then Return 0R
            Return Math.Sqrt(d2)
        End Function

        ' robust segment-segment distance squared (VB는 대소문자 구분X → d/D 변수 충돌 방지)
        Private Shared Function SegmentDistanceSquared(p1 As XYZ, p2 As XYZ, q1 As XYZ, q2 As XYZ) As Double
            Dim u As XYZ = p2 - p1
            Dim v As XYZ = q2 - q1
            Dim w As XYZ = p1 - q1

            Dim a0 As Double = u.DotProduct(u)
            Dim b0 As Double = u.DotProduct(v)
            Dim c0 As Double = v.DotProduct(v)
            Dim d0 As Double = u.DotProduct(w)
            Dim e0 As Double = v.DotProduct(w)

            Dim denom As Double = a0 * c0 - b0 * b0
            Dim sc As Double, sN As Double, sD As Double = denom
            Dim tc As Double, tN As Double, tD As Double = denom

            Const EPS As Double = 0.000000000001

            If denom < EPS Then
                sN = 0.0R
                sD = 1.0R
                tN = e0
                tD = c0
            Else
                sN = (b0 * e0 - c0 * d0)
                tN = (a0 * e0 - b0 * d0)
                If sN < 0.0R Then
                    sN = 0.0R
                    tN = e0
                    tD = c0
                ElseIf sN > sD Then
                    sN = sD
                    tN = e0 + b0
                    tD = c0
                End If
            End If

            If tN < 0.0R Then
                tN = 0.0R
                If -d0 < 0.0R Then
                    sN = 0.0R
                ElseIf -d0 > a0 Then
                    sN = sD
                Else
                    sN = -d0
                    sD = a0
                End If
            ElseIf tN > tD Then
                tN = tD
                If (-d0 + b0) < 0.0R Then
                    sN = 0.0R
                ElseIf (-d0 + b0) > a0 Then
                    sN = sD
                Else
                    sN = (-d0 + b0)
                    sD = a0
                End If
            End If

            sc = If(Math.Abs(sN) < EPS, 0.0R, sN / sD)
            tc = If(Math.Abs(tN) < EPS, 0.0R, tN / tD)

            Dim dP As XYZ = w + (u * sc) - (v * tc)
            Return dP.DotProduct(dP)
        End Function

        Private Shared Function SegmentOverlapLengthAlongAxis(a0 As XYZ, a1 As XYZ, b0 As XYZ, b1 As XYZ, uA As XYZ, lenA As Double) As Double
            Dim t0 As Double = (b0 - a0).DotProduct(uA)
            Dim t1 As Double = (b1 - a0).DotProduct(uA)
            Dim bMin As Double = Math.Min(t0, t1)
            Dim bMax As Double = Math.Max(t0, t1)

            Dim i0 As Double = Math.Max(0.0R, bMin)
            Dim i1 As Double = Math.Min(lenA, bMax)
            Dim ov As Double = i1 - i0
            If ov < 0.0R Then Return 0.0R
            Return ov
        End Function

        Private Shared Function PackCell(ix As Integer, iy As Integer) As Long
            Dim x As Long = CLng(ix)
            Dim y As Long = CLng(iy) And &HFFFFFFFFL
            Return (x << 32) Or y
        End Function

        Private Shared Function MakePairKey(a As Integer, b As Integer) As Long
            Dim lo As Long = Math.Min(a, b)
            Dim hi As Long = Math.Max(a, b)
            Return (hi << 32) Or (lo And &HFFFFFFFFL)
        End Function

#End Region

#Region "중복 키/MEP 치수"

        Private Shared Function TryBuildCurveDuplicateKey(e As Element,
                                                     tolFeet As Double,
                                                     catName As String,
                                                     famName As String,
                                                     typName As String,
                                                     lvl As Integer,
                                                     ByRef key As String) As Boolean
            key = Nothing

            Dim p0 As XYZ = Nothing, p1 As XYZ = Nothing
            If Not TryGetCurveEndpoints(e, p0, p1) Then Return False

            Dim q = Function(x As Double) As Long
                        Return CLng(Math.Round(x / tolFeet))
                    End Function

            Dim a As String = $"P({q(p0.X)},{q(p0.Y)},{q(p0.Z)})"
            Dim b As String = $"P({q(p1.X)},{q(p1.Y)},{q(p1.Z)})"

            Dim pA As String = a
            Dim pB As String = b
            If String.CompareOrdinal(pA, pB) > 0 Then
                Dim tmp = pA : pA = pB : pB = tmp
            End If

            Dim sizeKey As String = GetMEPSizeKey(e)

            key = String.Concat(catName, "|", famName, "|", typName, "|",
                            "L", lvl.ToString(), "|",
                            "LC|", pA, "|", pB, "|S(", sizeKey, ")")
            Return True
        End Function

        Private Shared Function TryGetCurveEndpoints(e As Element, ByRef p0 As XYZ, ByRef p1 As XYZ) As Boolean
            p0 = Nothing : p1 = Nothing
            If e Is Nothing Then Return False

            Try
                Dim lc As LocationCurve = TryCast(e.Location, LocationCurve)
                If lc Is Nothing OrElse lc.Curve Is Nothing Then Return False
                Dim c As Curve = lc.Curve
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                If p0 Is Nothing OrElse p1 Is Nothing Then Return False
                Return True
            Catch
            End Try

            Return False
        End Function

        Private Shared Function GetMEPSizeKey(e As Element) As String
            Const sizeTol As Double = (1.0R / 256.0R) ' ~1.19mm
            Dim qS = Function(x As Double) As Long
                         Return CLng(Math.Round(x / sizeTol))
                     End Function

            Try
                If TypeOf e Is Autodesk.Revit.DB.Plumbing.Pipe Then
                    Dim d As Double = GetParamDouble(e, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                    If d > 0 Then Return "D" & qS(d).ToString()
                End If

                If TypeOf e Is Autodesk.Revit.DB.Mechanical.Duct Then
                    Dim w As Double = GetParamDouble(e, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                    Dim h As Double = GetParamDouble(e, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                    If w > 0 AndAlso h > 0 Then Return "W" & qS(w) & "H" & qS(h)
                    If w > 0 Then Return "W" & qS(w)
                End If

                If TypeOf e Is Autodesk.Revit.DB.Electrical.Conduit Then
                    Dim d As Double = GetParamDouble(e, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                    If d > 0 Then Return "D" & qS(d).ToString()
                End If

                If TypeOf e Is Autodesk.Revit.DB.Electrical.CableTray Then
                    Dim w As Double = GetParamDouble(e, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                    Dim h As Double = GetParamDouble(e, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                    If w > 0 AndAlso h > 0 Then Return "W" & qS(w) & "H" & qS(h)
                    If w > 0 Then Return "W" & qS(w)
                End If
            Catch
            End Try

            Return ""
        End Function

        Private Shared Function GetParamDouble(e As Element, bip As BuiltInParameter) As Double
            Try
                Dim p As Parameter = e.Parameter(bip)
                If p IsNot Nothing AndAlso p.StorageType = StorageType.Double Then
                    Dim v As Double = p.AsDouble()
                    If v > 0 Then Return v
                End If
            Catch
            End Try
            Return 0R
        End Function

        Private Shared Function TryGetCrossSectionRadius(e As Element, ByRef radius As Double) As Boolean
            radius = 0R
            If e Is Nothing Then Return False

            Try
                If TypeOf e Is Autodesk.Revit.DB.Plumbing.Pipe Then
                    Dim d As Double = GetParamDouble(e, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                    If d > 0 Then
                        radius = d * 0.5R
                        Return True
                    End If
                End If

                If TypeOf e Is Autodesk.Revit.DB.Electrical.Conduit Then
                    Dim d As Double = GetParamDouble(e, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                    If d > 0 Then
                        radius = d * 0.5R
                        Return True
                    End If
                End If

                If TypeOf e Is Autodesk.Revit.DB.Mechanical.Duct Then
                    Dim w As Double = GetParamDouble(e, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                    Dim h As Double = GetParamDouble(e, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                    If w > 0 AndAlso h > 0 Then
                        radius = Math.Sqrt((w * 0.5R) * (w * 0.5R) + (h * 0.5R) * (h * 0.5R))
                        Return True
                    ElseIf w > 0 Then
                        radius = w * 0.5R
                        Return True
                    End If
                End If

                If TypeOf e Is Autodesk.Revit.DB.Electrical.CableTray Then
                    Dim w As Double = GetParamDouble(e, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                    Dim h As Double = GetParamDouble(e, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                    If w > 0 AndAlso h > 0 Then
                        radius = Math.Sqrt((w * 0.5R) * (w * 0.5R) + (h * 0.5R) * (h * 0.5R))
                        Return True
                    ElseIf w > 0 Then
                        radius = w * 0.5R
                        Return True
                    End If
                End If
            Catch
            End Try

            Return False
        End Function

#End Region

#Region "필터/유틸(물량 필터 강화)"

        Private Shared Function ShouldSkipForQuantity(e As Element) As Boolean
            If e Is Nothing Then Return True
            If TypeOf e Is ImportInstance Then Return True

            ' 뷰/시트/카메라 등 표시용 요소 제외
            If TypeOf e Is View Then Return True
            If TypeOf e Is Viewport Then Return True

            Try
                If e.ViewSpecific Then Return True
            Catch
                Return True
            End Try

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

            ' ✅ 복합(중첩) 패밀리: 하위 요소는 결과에서 제외 (최상위만)
            Try
                If _nestedSharedIds IsNot Nothing AndAlso _nestedSharedIds.Contains(e.Id.IntegerValue) Then Return True
            Catch
            End Try

            ' ✅ Insulation/Lining 계열은 중복/간섭에서 제외
            Try
                Dim bic As BuiltInCategory = CType(cat.Id.IntegerValue, BuiltInCategory)
                Select Case bic
                    Case BuiltInCategory.OST_PipeInsulations,
             BuiltInCategory.OST_DuctInsulations,
             BuiltInCategory.OST_DuctLinings
                        Return True
                End Select
            Catch
            End Try

            ' 타입으로도 한 번 더 방어(버전별 클래스 차이 대비)
            Try
                If TypeOf e Is Autodesk.Revit.DB.Plumbing.PipeInsulation Then Return True
            Catch
            End Try
            ' 카메라/래스터 이미지 등 비물리 요소 제외
            Try
                Dim cid As Integer = cat.Id.IntegerValue
                If cid = CInt(BuiltInCategory.OST_Cameras) Then Return True
                If cid = CInt(BuiltInCategory.OST_RasterImages) Then Return True
            Catch
            End Try

            ' 참조/기준/선류 제외
            If TypeOf e Is CurveElement Then Return True
            If TypeOf e Is ReferencePlane Then Return True
            If TypeOf e Is Level Then Return True
            If TypeOf e Is Grid Then Return True
            If TypeOf e Is DatumPlane Then Return True

            ' 합의: Part/Room/Space/Area/Rebar 제외
            If TypeOf e Is Part Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Architecture.Room Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.Space Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Area Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.Rebar Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.AreaReinforcement Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Structure.PathReinforcement Then Return True

            ' 자동 종속 제외
            If TypeOf e Is Autodesk.Revit.DB.Plumbing.PipeInsulation Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.DuctInsulation Then Return True
            If TypeOf e Is Autodesk.Revit.DB.Mechanical.DuctLining Then Return True

            ' 중첩 패밀리 제외
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing Then
                Try
                    If fi.SuperComponent IsNot Nothing Then Return True
                    If _nestedSharedIds IsNot Nothing AndAlso _nestedSharedIds.Contains(fi.Id.IntegerValue) Then Return True
                Catch
                End Try
            End If

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

        ' ✅ 키워드(콤마) 제외: Category/Family/Type/Name에 포함되면 대상에서 제외
        Private Shared Function ShouldExcludeByKeywords(e As Element, kws As List(Of String)) As Boolean
            If e Is Nothing OrElse kws Is Nothing OrElse kws.Count = 0 Then Return False

            Dim parts As New List(Of String)()

            Try
                If e.Category IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Category.Name) Then parts.Add(e.Category.Name)
            Catch
            End Try

            Try
                Dim fi As FamilyInstance = TryCast(e, FamilyInstance)
                If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then
                    Try
                        If fi.Symbol.Family IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(fi.Symbol.Family.Name) Then parts.Add(fi.Symbol.Family.Name)
                    Catch
                    End Try
                    Try
                        If Not String.IsNullOrWhiteSpace(fi.Symbol.Name) Then parts.Add(fi.Symbol.Name)
                    Catch
                    End Try
                End If
            Catch
            End Try

            Try
                If Not String.IsNullOrWhiteSpace(e.Name) Then parts.Add(e.Name)
            Catch
            End Try

            Dim hay As String = String.Join(" | ", parts).ToLowerInvariant()
            If hay.Length = 0 Then Return False

            For Each kw As String In kws
                If String.IsNullOrWhiteSpace(kw) Then Continue For
                If hay.Contains(kw.ToLowerInvariant()) Then Return True
            Next

            Return False
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

        Private Function CountGroups(rows As IEnumerable(Of DupRowDto)) As Integer
            If rows Is Nothing Then Return 0

            Dim gkCount As Integer =
            rows.Select(Function(r) If(r.GroupKey, "")) _
                .Where(Function(s) Not String.IsNullOrWhiteSpace(s)) _
                .Distinct(StringComparer.Ordinal) _
                .Count()

            If gkCount > 0 Then Return gkCount

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

            Dim bodyPad As New WControls.Border() With {
            .Padding = New WPF.Thickness(20),
            .Background = New WMedia.SolidColorBrush(bgCard),
            .CornerRadius = New WPF.CornerRadius(0, 0, 14, 14)
        }

            Dim body As New WControls.StackPanel() With {.Orientation = WControls.Orientation.Vertical}

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

            Dim pathTb As New WControls.TextBlock() With {
            .Text = $"파일: {outPath}",
            .TextWrapping = WPF.TextWrapping.Wrap,
            .Foreground = New WMedia.SolidColorBrush(fgSub),
            .Margin = New WPF.Thickness(0, 0, 0, 14)
        }

            Dim question As New WControls.TextBlock() With {
            .Text = questionText,
            .Foreground = New WMedia.SolidColorBrush(fgMain),
            .Margin = New WPF.Thickness(0, 4, 0, 10)
        }

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
