Option Explicit On
Option Strict On

Imports Autodesk.Revit.DB
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.CompilerServices

' Revit 2019~2023에서 사용하던 BuiltInParameterGroup 을
' Revit 2025 (GroupTypeId 기반)에서만 호환용으로 제공하는 쉼 레이어
' ▶ Global.Autodesk.Revit.DB 로 선언해서
'    실제 타입 이름이 정확히 Autodesk.Revit.DB.BuiltInParameterGroup 이 되도록 한다.
Namespace Global.Autodesk.Revit.DB

    ''' <summary>
    ''' 최소 호환용 BuiltInParameterGroup enum.
    ''' 실제 값은 중요하지 않고, GroupTypeId 매핑에만 사용된다.
    ''' ParamPropagateService.vb 에서 사용하는 그룹들만 넣어두고,
    ''' 나머지는 Else 분기로 Data 그룹으로 처리한다.
    ''' </summary>
    Public Enum BuiltInParameterGroup
        PG_DATA
        PG_TEXT
        PG_GEOMETRY
        PG_CONSTRAINTS
        PG_IDENTITY_DATA
        PG_MATERIALS
        PG_GRAPHICS
        PG_ANALYSIS
        PG_GENERAL
    End Enum

    ''' <summary>
    ''' FamilyManager.AddParameter/ReplaceParameter 에
    ''' BuiltInParameterGroup 을 그대로 넘길 수 있도록 하는 확장 메서드.
    ''' 내부에서는 Revit 2025 API가 요구하는 GroupTypeId로 변환한다.
    ''' </summary>
    Public Module BuiltInParameterGroupCompat

        Public NotInheritable Class GroupOptionInfo
            Public Property Id As Integer
            Public Property Key As String = String.Empty
            Public Property Label As String = String.Empty
            Public Property GroupTypeIdValue As ForgeTypeId
        End Class

        Private ReadOnly _groupOptionLock As New Object()
        Private _groupOptions As List(Of GroupOptionInfo)

        Public Function GetGroupOptions() As List(Of GroupOptionInfo)
            SyncLock _groupOptionLock
                If _groupOptions Is Nothing Then
                    _groupOptions = BuildGroupOptions()
                End If

                Return _groupOptions.
                    Select(Function(x) New GroupOptionInfo With {
                        .Id = x.Id,
                        .Key = x.Key,
                        .Label = x.Label,
                        .GroupTypeIdValue = x.GroupTypeIdValue
                    }).
                    ToList()
            End SyncLock
        End Function

        <Extension>
        Public Function ToGroupTypeId(group As BuiltInParameterGroup) As ForgeTypeId
            Dim dynamicGroup As ForgeTypeId = ResolveCatalogGroupTypeId(CInt(group))
            If dynamicGroup IsNot Nothing Then Return dynamicGroup

            Return MapGroup(group)
        End Function

        Public Function FromGroupTypeId(groupId As ForgeTypeId) As BuiltInParameterGroup
            If groupId IsNot Nothing Then
                Dim catalogMatch = GetGroupOptions().
                    FirstOrDefault(Function(x) String.Equals(GetTypeIdKey(x.GroupTypeIdValue), GetTypeIdKey(groupId), StringComparison.OrdinalIgnoreCase))
                If catalogMatch IsNot Nothing Then
                    Return CType(catalogMatch.Id, BuiltInParameterGroup)
                End If
            End If

            Return ToBuiltInGroup(groupId)
        End Function

        Public Function FromSerializedValue(value As Object,
                                            Optional defaultGroup As BuiltInParameterGroup = BuiltInParameterGroup.PG_DATA) As BuiltInParameterGroup
            If value Is Nothing Then Return defaultGroup

            Dim raw As String = ""
            Try
                raw = Convert.ToString(value).Trim()
            Catch
                raw = ""
            End Try

            If raw <> "" Then
                Dim iv As Integer
                If Integer.TryParse(raw, iv) Then
                    Return CType(iv, BuiltInParameterGroup)
                End If

                Dim catalogMatch = GetGroupOptions().
                    FirstOrDefault(Function(x) String.Equals(Convert.ToString(x.Id), raw, StringComparison.OrdinalIgnoreCase) OrElse
                                              String.Equals(x.Key, raw, StringComparison.OrdinalIgnoreCase))
                If catalogMatch IsNot Nothing Then
                    Return CType(catalogMatch.Id, BuiltInParameterGroup)
                End If

                Dim parsed As BuiltInParameterGroup
                If [Enum].TryParse(raw, True, parsed) Then
                    Return parsed
                End If
            End If

            Try
                Dim iv As Integer = Convert.ToInt32(value)
                Return CType(iv, BuiltInParameterGroup)
            Catch
            End Try

            Return defaultGroup
        End Function

        <Extension>
        Public Function IsInGroup(def As Definition, group As BuiltInParameterGroup) As Boolean
            If def Is Nothing Then Return False

            Return ToBuiltInGroup(def.GetGroupTypeId()) = group
        End Function

        Private Function MapGroup(group As BuiltInParameterGroup) As ForgeTypeId
            Select Case group
                Case BuiltInParameterGroup.PG_TEXT
                    Return GroupTypeId.Text
                Case BuiltInParameterGroup.PG_GEOMETRY
                    Return GroupTypeId.Geometry
                Case BuiltInParameterGroup.PG_CONSTRAINTS
                    Return GroupTypeId.Constraints
                Case BuiltInParameterGroup.PG_MATERIALS
                    Return GroupTypeId.Materials
                Case BuiltInParameterGroup.PG_GRAPHICS
                    Return GroupTypeId.Graphics

                Case BuiltInParameterGroup.PG_IDENTITY_DATA
                    Return GroupTypeId.IdentityData

                ' 나머지 그룹들은 전부 Data 그룹으로 통일
                Case BuiltInParameterGroup.PG_DATA,
                     BuiltInParameterGroup.PG_ANALYSIS,
                     BuiltInParameterGroup.PG_GENERAL
                    Return GroupTypeId.Data

                Case Else
                    Return GroupTypeId.Data
            End Select
        End Function

        Private Function ResolveCatalogGroupTypeId(rawId As Integer) As ForgeTypeId
            If rawId < 1000 Then Return Nothing

            Dim match = GetGroupOptions().FirstOrDefault(Function(x) x.Id = rawId)
            If match Is Nothing Then Return Nothing
            Return match.GroupTypeIdValue
        End Function

        Private Function BuildGroupOptions() As List(Of GroupOptionInfo)
            Dim items As New List(Of GroupOptionInfo)()
            Dim seenTypeIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim preferredNames As String() = {"Text", "IdentityData", "Data", "Constraints"}

            Dim props As IEnumerable(Of PropertyInfo) =
                GetType(GroupTypeId).GetProperties(BindingFlags.Public Or BindingFlags.Static).
                Where(Function(p) p IsNot Nothing AndAlso p.CanRead AndAlso p.PropertyType Is GetType(ForgeTypeId))

            For Each prop As PropertyInfo In props
                Dim gtid As ForgeTypeId = Nothing
                Try
                    gtid = TryCast(prop.GetValue(Nothing, Nothing), ForgeTypeId)
                Catch
                    gtid = Nothing
                End Try
                If gtid Is Nothing Then Continue For

                Dim key As String = GetTypeIdKey(gtid)
                If String.IsNullOrWhiteSpace(key) Then key = prop.Name
                If String.IsNullOrWhiteSpace(key) OrElse seenTypeIds.Contains(key) Then Continue For

                Dim label As String = ""
                Try
                    label = LabelUtils.GetLabelForGroup(gtid)
                Catch
                    label = ""
                End Try
                If String.IsNullOrWhiteSpace(label) Then
                    label = HumanizeName(prop.Name)
                End If

                seenTypeIds.Add(key)
                items.Add(New GroupOptionInfo With {
                    .Key = key,
                    .Label = label,
                    .GroupTypeIdValue = gtid
                })
            Next

            Dim ordered As List(Of GroupOptionInfo) =
                items.
                OrderBy(Function(x)
                            Dim idx As Integer = Array.IndexOf(preferredNames, HumanizeLookupName(x))
                            Return If(idx >= 0, idx, Integer.MaxValue)
                        End Function).
                ThenBy(Function(x) x.Label, StringComparer.CurrentCultureIgnoreCase).
                ToList()

            For i As Integer = 0 To ordered.Count - 1
                ordered(i).Id = 1000 + i
            Next

            Return ordered
        End Function

        Private Function HumanizeLookupName(x As GroupOptionInfo) As String
            If x Is Nothing Then Return ""

            Dim key As String = x.Key
            If String.IsNullOrWhiteSpace(key) Then Return ""

            Dim tail As String = key
            Dim lastColon As Integer = tail.LastIndexOf(":"c)
            If lastColon >= 0 AndAlso lastColon < tail.Length - 1 Then
                tail = tail.Substring(lastColon + 1)
            End If

            Dim lastDash As Integer = tail.LastIndexOf("-"c)
            If lastDash >= 0 AndAlso lastDash < tail.Length - 1 Then
                tail = tail.Substring(lastDash + 1)
            End If

            Return HumanizeName(tail).Replace(" ", "")
        End Function

        Private Function GetTypeIdKey(groupId As ForgeTypeId) As String
            If groupId Is Nothing Then Return ""

            Try
                Return If(groupId.TypeId, "")
            Catch
                Return ""
            End Try
        End Function

        Private Function HumanizeName(raw As String) As String
            Dim source As String = If(raw, "").Trim()
            If source = "" Then Return ""

            Dim chars As New List(Of Char)()
            For i As Integer = 0 To source.Length - 1
                Dim ch As Char = source(i)
                If i > 0 AndAlso Char.IsUpper(ch) AndAlso (Char.IsLower(source(i - 1)) OrElse Char.IsDigit(source(i - 1))) Then
                    chars.Add(" "c)
                End If
                chars.Add(ch)
            Next

            Return New String(chars.ToArray()).Trim()
        End Function

        Public Function ToBuiltInGroup(groupId As ForgeTypeId) As BuiltInParameterGroup
            If groupId Is Nothing Then
                Return BuiltInParameterGroup.PG_DATA
            End If

            If groupId.Equals(GroupTypeId.Text) Then
                Return BuiltInParameterGroup.PG_TEXT
            End If
            If groupId.Equals(GroupTypeId.Geometry) Then
                Return BuiltInParameterGroup.PG_GEOMETRY
            End If
            If groupId.Equals(GroupTypeId.Constraints) Then
                Return BuiltInParameterGroup.PG_CONSTRAINTS
            End If
            If groupId.Equals(GroupTypeId.Materials) Then
                Return BuiltInParameterGroup.PG_MATERIALS
            End If
            If groupId.Equals(GroupTypeId.Graphics) Then
                Return BuiltInParameterGroup.PG_GRAPHICS
            End If
            If groupId.Equals(GroupTypeId.IdentityData) Then
                Return BuiltInParameterGroup.PG_IDENTITY_DATA
            End If

            ' 나머지는 전부 Data로 처리
            Return BuiltInParameterGroup.PG_DATA
        End Function
    End Module

    Public Module FamilyManagerCompatExtensions

        <Extension>
        Public Function AddParameter(fm As FamilyManager,
                                     def As ExternalDefinition,
                                     group As BuiltInParameterGroup,
                                     isInstance As Boolean) As FamilyParameter

            Dim gId As ForgeTypeId = BuiltInParameterGroupCompat.ToGroupTypeId(group)
            Return fm.AddParameter(def, gId, isInstance)
        End Function

        <Extension>
        Public Function ReplaceParameter(fm As FamilyManager,
                                         param As FamilyParameter,
                                         def As ExternalDefinition,
                                         group As BuiltInParameterGroup,
                                         isInstance As Boolean) As FamilyParameter

            Dim gId As ForgeTypeId = BuiltInParameterGroupCompat.ToGroupTypeId(group)
            Return fm.ReplaceParameter(param, def, gId, isInstance)
        End Function

    End Module

End Namespace
