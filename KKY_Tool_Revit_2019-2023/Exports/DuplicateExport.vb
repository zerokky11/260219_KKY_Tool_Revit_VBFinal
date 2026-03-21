Imports System.Data
Imports System.Linq
Imports KKY_Tool_Revit.Infrastructure

Namespace Exports

    Public Class DupRowDto
        Public Property FileName As String

        Public Property Id As String
        Public Property Category As String
        Public Property Family As String
        Public Property Type As String
        Public Property ConnectedIds As System.Collections.Generic.List(Of String)
        Public Property ExtraParams As System.Collections.Generic.Dictionary(Of String, String)

        ' 상위호환: 그룹키(Host가 만들어 내려줌)
        Public Property GroupKey As String
    End Class

Public Class PairRowDto
    Public Property FileName As String
    Public Property GroupKey As String

    Public Property AId As String
    Public Property ACategory As String
    Public Property AFamily As String
    Public Property AType As String

    Public Property BId As String
    Public Property BCategory As String
    Public Property BFamily As String
    Public Property BType As String
    Public Property Comment As String
    Public Property AExtraParams As System.Collections.Generic.Dictionary(Of String, String)
    Public Property BExtraParams As System.Collections.Generic.Dictionary(Of String, String)
End Class


    Public Module DuplicateExport

        ' ✅ 기존 호출 호환: reportTitle은 마지막 Optional로만 추가
        Public Function Save(rows As System.Collections.IEnumerable,
                             Optional doAutoFit As Boolean = False,
                             Optional progressChannel As String = Nothing,
                             Optional reportTitle As String = Nothing,
                             Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing) As String
            Dim mapped = MapRows(rows)
            Dim dt = BuildSimpleTable(mapped, extraParamNames)
            Dim title = If(String.IsNullOrWhiteSpace(reportTitle), "Duplicates (Simple)", reportTitle)
            Return ExcelCore.PickAndSaveXlsx(title, dt, "Duplicates.xlsx", doAutoFit, progressChannel)
        End Function

        ' ✅ 기존 호출 호환: reportTitle은 마지막 Optional로만 추가
        Public Sub Save(outPath As String,
                        rows As System.Collections.IEnumerable,
                        Optional doAutoFit As Boolean = False,
                        Optional progressChannel As String = Nothing,
                        Optional reportTitle As String = Nothing,
                        Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing)
            Export(outPath, rows, doAutoFit, progressChannel, reportTitle, extraParamNames)
        End Sub

        Public Sub Export(outPath As String,
                          rows As System.Collections.IEnumerable,
                          Optional doAutoFit As Boolean = False,
                          Optional progressChannel As String = Nothing,
                          Optional reportTitle As String = Nothing,
                          Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing)
            Dim mapped = MapRows(rows)
            Dim dt = BuildSimpleTable(mapped, extraParamNames)
            Dim title = If(String.IsNullOrWhiteSpace(reportTitle), "Duplicates (Simple)", reportTitle)
            ExcelCore.SaveStyledSimple(outPath, title, dt, "Group", doAutoFit, progressChannel)
        End Sub

        

' ===== Pair Export (Self Clash pairs) =====
Public Function SavePairs(pairs As System.Collections.IEnumerable,
                          Optional doAutoFit As Boolean = False,
                          Optional progressChannel As String = Nothing,
                          Optional reportTitle As String = Nothing,
                          Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing) As String
    Dim mapped = MapPairs(pairs)
    Dim dt = BuildPairTable(mapped, extraParamNames)
    Dim title = If(String.IsNullOrWhiteSpace(reportTitle), "Self Clash Pairs", reportTitle)
    Return ExcelCore.PickAndSaveXlsx(title, dt, "ClashPairs.xlsx", doAutoFit, progressChannel)
End Function

Public Sub ExportPairs(outPath As String,
                       pairs As System.Collections.IEnumerable,
                       Optional doAutoFit As Boolean = False,
                       Optional progressChannel As String = Nothing,
                       Optional reportTitle As String = Nothing,
                       Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing)
    Dim mapped = MapPairs(pairs)
    Dim dt = BuildPairTable(mapped, extraParamNames)
    Dim title = If(String.IsNullOrWhiteSpace(reportTitle), "Self Clash Pairs", reportTitle)
    ExcelCore.SaveStyledSimple(outPath, title, dt, "Group", doAutoFit, progressChannel)
End Sub

Private Function MapPairs(pairs As System.Collections.IEnumerable) As System.Collections.Generic.List(Of PairRowDto)
    Dim list As New System.Collections.Generic.List(Of PairRowDto)
    If pairs Is Nothing Then Return list

    For Each o In pairs
        Dim it As New PairRowDto()
        it.FileName = ReadProp(o, "FileName", "File", "Rvt", "Doc")
        it.GroupKey = ReadProp(o, "GroupKey", "groupKey", "Group", "group")

        it.AId = ReadProp(o, "AId", "aId", "IdA", "A_ID")
        it.ACategory = ReadProp(o, "ACategory", "aCategory", "CategoryA")
        it.AFamily = ReadProp(o, "AFamily", "aFamily", "FamilyA")
        it.AType = ReadProp(o, "AType", "aType", "TypeA")

        it.BId = ReadProp(o, "BId", "bId", "IdB", "B_ID")
        it.BCategory = ReadProp(o, "BCategory", "bCategory", "CategoryB")
        it.BFamily = ReadProp(o, "BFamily", "bFamily", "FamilyB")
        it.BType = ReadProp(o, "BType", "bType", "TypeB")
        it.Comment = ReadProp(o, "Comment", "comment", "Note", "note", "Reason", "reason")
        it.AExtraParams = ReadStringMap(o, "AExtraParams", "aExtraParams")
        it.BExtraParams = ReadStringMap(o, "BExtraParams", "bExtraParams")

        list.Add(it)
    Next

    Return list
End Function

Private Function BuildPairTable(pairs As System.Collections.Generic.List(Of PairRowDto),
                                Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing) As DataTable
    Dim dt As New DataTable("pairs")
    Dim orderedExtraNames = BuildOrderedExtraParamNames(extraParamNames, pairs.Select(Function(p) p.AExtraParams), pairs.Select(Function(p) p.BExtraParams))
    dt.Columns.Add("File")
    dt.Columns.Add("Group")

    dt.Columns.Add("A_ID")
    dt.Columns.Add("A_Category")
    dt.Columns.Add("A_Family")
    dt.Columns.Add("A_Type")
    For Each paramName In orderedExtraNames
        dt.Columns.Add("A_" & paramName)
    Next

    dt.Columns.Add("B_ID")
    dt.Columns.Add("B_Category")
    dt.Columns.Add("B_Family")
    dt.Columns.Add("B_Type")
    For Each paramName In orderedExtraNames
        dt.Columns.Add("B_" & paramName)
    Next
    dt.Columns.Add("Comment")

    If pairs IsNot Nothing AndAlso pairs.Count > 0 Then
        For Each p In pairs
            Dim dr = dt.NewRow()
            dr("File") = Nz(p.FileName)
            dr("Group") = Nz(p.GroupKey)

            dr("A_ID") = Nz(p.AId)
            dr("A_Category") = Nz(p.ACategory)
            dr("A_Family") = Nz(p.AFamily)
            dr("A_Type") = Nz(p.AType)
            For Each paramName In orderedExtraNames
                dr("A_" & paramName) = GetMapValue(p.AExtraParams, paramName)
            Next

            dr("B_ID") = Nz(p.BId)
            dr("B_Category") = Nz(p.BCategory)
            dr("B_Family") = Nz(p.BFamily)
            dr("B_Type") = Nz(p.BType)
            For Each paramName In orderedExtraNames
                dr("B_" & paramName) = GetMapValue(p.BExtraParams, paramName)
            Next
            dr("Comment") = Nz(p.Comment)

            dt.Rows.Add(dr)
        Next
    Else
        Dim dr = dt.NewRow()
        dr(0) = "오류가 없습니다."
        dt.Rows.Add(dr)
    End If

    Return dt
End Function

Private Function MapRows(rows As System.Collections.IEnumerable) As System.Collections.Generic.List(Of DupRowDto)
            Dim list As New System.Collections.Generic.List(Of DupRowDto)
            If rows Is Nothing Then Return list

            For Each o In rows
                Dim it As New DupRowDto()
                it.FileName = ReadProp(o, "FileName", "File", "Rvt", "Doc")
                it.Id = ReadProp(o, "Id", "ID", "ElementId", "ElementID", "elementId")
                it.Category = ReadProp(o, "Category", "category")
                it.Family = ReadProp(o, "Family", "family")
                it.Type = ReadProp(o, "Type", "type")
                it.ConnectedIds = ReadList(o, "ConnectedIds", "connectedIds", "Links", "links", "connected", "Connected", "ConnectedElements")
                it.ExtraParams = ReadStringMap(o, "ExtraParams", "extraParams")

                it.GroupKey = ReadProp(o, "GroupKey", "groupKey", "Group", "group")

                list.Add(it)
            Next

            Return list
        End Function

        Private Function BuildSimpleTable(rows As System.Collections.Generic.List(Of DupRowDto),
                                          Optional extraParamNames As System.Collections.Generic.IList(Of String) = Nothing) As DataTable
            Dim dt As New DataTable("simple")
            Dim orderedExtraNames = BuildOrderedExtraParamNames(extraParamNames, rows.Select(Function(r) r.ExtraParams))
            dt.Columns.Add("File")
            dt.Columns.Add("Group")
            dt.Columns.Add("ID")
            dt.Columns.Add("Category")
            dt.Columns.Add("Family")
            dt.Columns.Add("Type")
            For Each paramName In orderedExtraNames
                dt.Columns.Add(paramName)
            Next

            Dim groupList = GroupByLogic(rows)

            For i = 0 To groupList.Count - 1
                Dim gName As String = $"Group{i + 1}"

                ' groupKey가 있으면 우선 사용(엑셀에서도 JS와 동일 그룹 유지)
                Dim gk As String = ""
                If groupList(i).Count > 0 Then gk = Nz(groupList(i)(0).GroupKey)
                If Not String.IsNullOrWhiteSpace(gk) Then gName = gk

                For Each r In groupList(i)
                    Dim famOut As String = If(String.IsNullOrWhiteSpace(r.Family), If(String.IsNullOrWhiteSpace(r.Category), "", r.Category & " Type"), r.Family)
                    Dim dr = dt.NewRow()
                    dr("File") = Nz(r.FileName)
                    dr("Group") = gName
                    dr("ID") = Nz(r.Id)
                    dr("Category") = Nz(r.Category)
                    dr("Family") = Nz(famOut)
                    dr("Type") = Nz(r.Type)
                    For Each paramName In orderedExtraNames
                        dr(paramName) = GetMapValue(r.ExtraParams, paramName)
                    Next
                    dt.Rows.Add(dr)
                Next
            Next

            If dt.Rows.Count = 0 Then
                Dim dr = dt.NewRow()
                dr(0) = "오류가 없습니다."
                dt.Rows.Add(dr)
            End If

            Return dt
        End Function

        Private Function GroupByLogic(items As System.Collections.Generic.List(Of DupRowDto)) As System.Collections.Generic.List(Of System.Collections.Generic.List(Of DupRowDto))
            Dim buckets As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.List(Of DupRowDto))()

            For Each r In items
                Dim gk As String = Nz(r.GroupKey)
                If Not String.IsNullOrWhiteSpace(gk) Then
                    If Not buckets.ContainsKey(gk) Then buckets(gk) = New System.Collections.Generic.List(Of DupRowDto)()
                    buckets(gk).Add(r)
                    Continue For
                End If

                ' fallback: 기존 로직(cat/fam/type/cluster)
                Dim fam As String = If(String.IsNullOrWhiteSpace(r.Family), If(String.IsNullOrWhiteSpace(r.Category), "", r.Category & " Type"), r.Family)
                Dim typ As String = If(String.IsNullOrWhiteSpace(r.Type), "", r.Type)
                Dim cat As String = If(String.IsNullOrWhiteSpace(r.Category), "", r.Category)

                Dim clusterSrc As New System.Collections.Generic.List(Of String)
                If Not String.IsNullOrWhiteSpace(r.Id) Then clusterSrc.Add(r.Id)
                If r.ConnectedIds IsNot Nothing Then clusterSrc.AddRange(r.ConnectedIds)

                Dim cluster =
                    clusterSrc _
                    .SelectMany(Function(s) SplitIds(s)) _
                    .Where(Function(x) Not String.IsNullOrWhiteSpace(x)) _
                    .Select(Function(x) x.Trim()) _
                    .Distinct() _
                    .OrderBy(Function(x) PadNum(x)) _
                    .ToList()

                Dim clusterKey As String = If(cluster.Count > 1, String.Join(",", cluster), "")
                Dim key = String.Join("|", {cat, fam, typ, clusterKey})

                If Not buckets.ContainsKey(key) Then buckets(key) = New System.Collections.Generic.List(Of DupRowDto)()
                buckets(key).Add(r)
            Next

            Return buckets.Values.ToList()
        End Function

        Private Function SplitIds(s As String) As System.Collections.Generic.IEnumerable(Of String)
            If String.IsNullOrWhiteSpace(s) Then Return Array.Empty(Of String)()
            Return s.Split(New Char() {","c, " "c, ";"c, "|"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        End Function

        Private Function PadNum(s As String) As String
            Dim n As Integer
            If Integer.TryParse(s, n) Then Return n.ToString("D10")
            Return s
        End Function

        Private Function Nz(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Return s
        End Function

        Private Function ReadProp(obj As Object, ParamArray names() As String) As String
            If obj Is Nothing Then Return ""
            For Each nm In names
                If String.IsNullOrEmpty(nm) Then Continue For
                Dim p = obj.GetType().GetProperty(nm)
                If p IsNot Nothing Then
                    Dim v = p.GetValue(obj, Nothing)
                    If v IsNot Nothing Then Return v.ToString()
                End If
            Next
            Return ""
        End Function

        Private Function ReadList(obj As Object, ParamArray names() As String) As System.Collections.Generic.List(Of String)
            Dim res As New System.Collections.Generic.List(Of String)
            If obj Is Nothing Then Return res

            For Each nm In names
                Dim p = obj.GetType().GetProperty(nm)
                If p Is Nothing Then Continue For

                Dim v = p.GetValue(obj, Nothing)
                If v Is Nothing Then Continue For

                If TypeOf v Is String Then
                    res.AddRange(SplitIds(DirectCast(v, String)))
                    Exit For
                End If

                If TypeOf v Is System.Collections.IEnumerable AndAlso Not TypeOf v Is String Then
                    For Each x In DirectCast(v, System.Collections.IEnumerable)
                        If x IsNot Nothing Then res.Add(x.ToString())
                    Next
                    Exit For
                End If
            Next

            Return res
        End Function

        Private Function ReadStringMap(obj As Object, ParamArray names() As String) As System.Collections.Generic.Dictionary(Of String, String)
            Dim res As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If obj Is Nothing Then Return res

            For Each nm In names
                If String.IsNullOrWhiteSpace(nm) Then Continue For
                Dim p = obj.GetType().GetProperty(nm)
                If p Is Nothing Then Continue For

                Dim v = p.GetValue(obj, Nothing)
                If v Is Nothing Then Continue For

                Dim dict = TryCast(v, System.Collections.IDictionary)
                If dict Is Nothing Then Continue For

                For Each de As System.Collections.DictionaryEntry In dict
                    Dim key As String = If(de.Key, "").ToString().Trim()
                    If String.IsNullOrWhiteSpace(key) Then Continue For
                    res(key) = If(de.Value, "").ToString()
                Next
                Exit For
            Next

            Return res
        End Function

        Private Function BuildOrderedExtraParamNames(requested As System.Collections.Generic.IList(Of String),
                                                     ParamArray maps() As IEnumerable(Of System.Collections.Generic.Dictionary(Of String, String))) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            Dim seen As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If requested IsNot Nothing Then
                For Each name In requested
                    Dim clean = Nz(name).Trim()
                    If String.IsNullOrWhiteSpace(clean) Then Continue For
                    If seen.Add(clean) Then result.Add(clean)
                Next
            End If

            If maps IsNot Nothing Then
                For Each mapGroup In maps
                    If mapGroup Is Nothing Then Continue For
                    For Each map In mapGroup
                        If map Is Nothing Then Continue For
                        For Each key In map.Keys
                            Dim clean = Nz(key).Trim()
                            If String.IsNullOrWhiteSpace(clean) Then Continue For
                            If seen.Add(clean) Then result.Add(clean)
                        Next
                    Next
                Next
            End If

            Return result
        End Function

        Private Function GetMapValue(map As System.Collections.Generic.Dictionary(Of String, String), key As String) As String
            If map Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""
            Dim value As String = Nothing
            If map.TryGetValue(key, value) Then Return Nz(value)
            Return ""
        End Function

    End Module

End Namespace
