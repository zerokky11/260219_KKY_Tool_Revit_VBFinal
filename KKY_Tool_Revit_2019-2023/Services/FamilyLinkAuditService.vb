Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Namespace Services

    Public Class FamilyLinkTargetParam
        Public Property Name As String = ""
        Public Property Guid As Guid
        Public Property GroupName As String = ""
        Public Property DataTypeToken As String = ""
    End Class

    Public Class FamilyLinkAuditRow
        Public Property FileName As String = ""
        Public Property HostFamilyName As String = ""
        Public Property HostFamilyCategory As String = ""
        Public Property NestedFamilyName As String = ""
        Public Property NestedTypeName As String = ""
        Public Property NestedInstanceId As String = ""
        Public Property NestedPath As String = ""
        Public Property NestingLevel As String = ""
        Public Property NestedCategory As String = ""
        Public Property NestedParamName As String = ""
        Public Property TargetParamName As String = ""
        Public Property ExpectedGuid As String = ""
        Public Property FoundScope As String = ""
        Public Property NestedParamGuid As String = ""
        Public Property NestedParamDataType As String = ""
        Public Property AssocHostParamName As String = ""
        Public Property HostParamGuid As String = ""
        Public Property HostParamIsShared As String = ""
        Public Property Issue As String = ""
        Public Property Notes As String = ""
    End Class

    Friend Enum FamilyLinkAuditIssue
        OK
        MissingAssociation
        GuidMismatch
        HostParamNotShared
        ParamNotFound
        DescendantNestedUnsupported
        [Error]
    End Enum

    Friend Enum FoundScope
        InstanceParam
        TypeParam
    End Enum

    Friend Class FoundParam
        Public Property P As Parameter
        Public Property Scope As FoundScope
    End Class

    Friend Class NestedFamilyStructureInfo
        Public Property HasNestedFamilies As Boolean
        Public Property NestedFamilyCount As Integer
    End Class

    Public NotInheritable Class FamilyLinkAuditService

        Private Sub New()
        End Sub

        Public Shared Function Run(app As UIApplication,
                                   rvtPaths As IList(Of String),
                                   targets As IList(Of FamilyLinkTargetParam),
                                   progress As Action(Of Integer, String)) As List(Of FamilyLinkAuditRow)

            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))

            Dim rows As New List(Of FamilyLinkAuditRow)()
            Dim targetMap As Dictionary(Of String, List(Of FamilyLinkTargetParam)) = BuildTargetMap(targets)
            If targetMap.Count = 0 Then Return rows

            Dim cleanedPaths As List(Of String) = NormalizePaths(rvtPaths)
            Dim total As Integer = cleanedPaths.Count
            If total = 0 Then Return rows

            For i As Integer = 0 To total - 1
                Dim rvtPath As String = cleanedPaths(i)
                Dim doc As Document = Nothing
                Dim fileName As String = SafeFileName(rvtPath)

                Try
                    ReportProgress(progress, total, i + 1, 0.02R, $"프로젝트 스캔 시작... {i + 1}/{total} {fileName}")

                    If String.IsNullOrWhiteSpace(rvtPath) Then Throw New ArgumentException("RVT 경로가 비어 있습니다.")
                    If Not File.Exists(rvtPath) Then Throw New FileNotFoundException("RVT 파일을 찾을 수 없습니다.", rvtPath)

                    doc = OpenProjectDocument(app.Application, rvtPath)
                    If doc Is Nothing Then Throw New InvalidOperationException("프로젝트 문서를 열 수 없습니다.")

                    Dim hostFamilies As List(Of Family) =
                        New FilteredElementCollector(doc).
                            OfClass(GetType(Family)).
                            Cast(Of Family)().
                            Where(Function(f) f IsNot Nothing AndAlso f.IsEditable AndAlso Not f.IsInPlace).
                            ToList()

                    Dim famTotal As Integer = hostFamilies.Count
                    Dim rvtName As String = fileName

                    If famTotal = 0 Then
                        ReportProgress(progress, total, i + 1, 1.0R, $"{rvtName}: 편집 가능한 패밀리가 없습니다.")
                        Continue For
                    End If

                    For fi As Integer = 0 To famTotal - 1
                        Dim fam As Family = hostFamilies(fi)
                        Dim frac As Double = 0.05R + 0.9R * SafeRatio(fi + 1, famTotal)
                        ReportProgress(progress, total, i + 1, frac, $"[{rvtName}] 패밀리 검사 중({fi + 1}/{famTotal})")
                        AuditFamilyAsHost(doc, fam, fileName, targetMap, rows)
                    Next

                    ReportProgress(progress, total, i + 1, 1.0R, $"완료: {rvtName}")

                Catch ex As Exception
                    rows.Add(New FamilyLinkAuditRow With {
                        .FileName = fileName,
                        .Issue = FamilyLinkAuditIssue.[Error].ToString(),
                        .Notes = $"Project open/scan error: {ex.Message}"
                    })
                Finally
                    If doc IsNot Nothing Then
                        Try
                            doc.Close(False)
                        Catch
                        End Try
                    End If
                End Try
            Next

            Return rows
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             rvtPath As String,
                                             targets As IList(Of FamilyLinkTargetParam),
                                             progress As Action(Of Integer, String)) As List(Of FamilyLinkAuditRow)
            Dim rows As New List(Of FamilyLinkAuditRow)()
            If doc Is Nothing Then Return rows

            Dim targetMap As Dictionary(Of String, List(Of FamilyLinkTargetParam)) = BuildTargetMap(targets)
            If targetMap.Count = 0 Then Return rows

            Dim fileName As String = SafeFileName(rvtPath)
            Try
                Dim hostFamilies As List(Of Family) =
                    New FilteredElementCollector(doc).
                        OfClass(GetType(Family)).
                        Cast(Of Family)().
                        Where(Function(f) f IsNot Nothing AndAlso f.IsEditable AndAlso Not f.IsInPlace).
                        ToList()

                Dim famTotal As Integer = hostFamilies.Count
                If famTotal = 0 Then
                    ReportProgress(progress, 1, 1, 1.0R, $"{fileName}: 편집 가능한 패밀리가 없습니다.")
                    Return rows
                End If

                For fi As Integer = 0 To famTotal - 1
                    Dim fam As Family = hostFamilies(fi)
                    Dim frac As Double = 0.05R + 0.9R * SafeRatio(fi + 1, famTotal)
                    ReportProgress(progress, 1, 1, frac, $"[{fileName}] 패밀리 검사 중({fi + 1}/{famTotal})")
                    AuditFamilyAsHost(doc, fam, fileName, targetMap, rows)
                Next

                ReportProgress(progress, 1, 1, 1.0R, $"완료: {fileName}")
            Catch ex As Exception
                rows.Add(New FamilyLinkAuditRow With {
                    .FileName = fileName,
                    .Issue = FamilyLinkAuditIssue.[Error].ToString(),
                    .Notes = $"Project scan error: {ex.Message}"
                })
            End Try

            Return rows
        End Function

        Private Shared Sub AuditFamilyAsHost(hostDoc As Document,
                                             hostFamily As Family,
                                             fileName As String,
                                             expectedByName As Dictionary(Of String, List(Of FamilyLinkTargetParam)),
                                             rows As List(Of FamilyLinkAuditRow))

            Dim famDoc As Document = Nothing
            Try
                famDoc = hostDoc.EditFamily(hostFamily)
                If famDoc Is Nothing OrElse Not famDoc.IsFamilyDocument Then Return

                Dim nestedInstances As List(Of FamilyInstance) =
                    New FilteredElementCollector(famDoc).
                        OfClass(GetType(FamilyInstance)).
                        WhereElementIsNotElementType().
                        Cast(Of FamilyInstance)().
                        Where(Function(x) x IsNot Nothing AndAlso x.Symbol IsNot Nothing AndAlso x.Symbol.Family IsNot Nothing).
                        ToList()

                If nestedInstances.Count = 0 Then Return

                Dim hostCat As String = ""
                Try
                    If hostFamily.FamilyCategory IsNot Nothing Then hostCat = hostFamily.FamilyCategory.Name
                Catch
                End Try

                AuditNestedInstancesRecursive(famDoc, hostFamily, hostCat, fileName, expectedByName, rows, "", 1)

            Finally
                If famDoc IsNot Nothing Then
                    Try
                        famDoc.Close(False)
                    Catch
                    End Try
                End If
            End Try
        End Sub

        Private Shared Sub AuditNestedInstancesRecursive(parentFamilyDoc As Document,
                                                         hostFamily As Family,
                                                         hostCat As String,
                                                         fileName As String,
                                                         expectedByName As Dictionary(Of String, List(Of FamilyLinkTargetParam)),
                                                         rows As List(Of FamilyLinkAuditRow),
                                                         pathPrefix As String,
                                                         nestingLevel As Integer)

            If parentFamilyDoc Is Nothing Then Return

            Dim nestedInstances As List(Of FamilyInstance) =
                New FilteredElementCollector(parentFamilyDoc).
                    OfClass(GetType(FamilyInstance)).
                    WhereElementIsNotElementType().
                    Cast(Of FamilyInstance)().
                    Where(Function(x) x IsNot Nothing AndAlso x.Symbol IsNot Nothing AndAlso x.Symbol.Family IsNot Nothing).
                    ToList()

            If nestedInstances.Count = 0 Then Return

            For Each fi As FamilyInstance In nestedInstances
                Dim nestedFam As Family = fi.Symbol.Family
                Dim nestedCat As String = ""
                Try
                    If fi.Category IsNot Nothing Then nestedCat = fi.Category.Name
                Catch
                End Try

                Dim nestedPath As String = BuildNestedPath(pathPrefix, nestedFam, fi)
                Dim nestedStructure As NestedFamilyStructureInfo = Nothing
                Dim childDoc As Document = Nothing

                Try
                    If nestedFam IsNot Nothing AndAlso nestedFam.IsEditable AndAlso Not nestedFam.IsInPlace Then
                        childDoc = parentFamilyDoc.EditFamily(nestedFam)
                        If childDoc IsNot Nothing AndAlso childDoc.IsFamilyDocument Then
                            nestedStructure = New NestedFamilyStructureInfo()
                            nestedStructure.NestedFamilyCount =
                                New FilteredElementCollector(childDoc).
                                    OfClass(GetType(FamilyInstance)).
                                    WhereElementIsNotElementType().
                                    GetElementCount()
                            nestedStructure.HasNestedFamilies = (nestedStructure.NestedFamilyCount > 0)
                        End If
                    End If
                Catch
                    childDoc = Nothing
                End Try

                If nestedStructure Is Nothing Then
                    nestedStructure = New NestedFamilyStructureInfo()
                End If

                If nestingLevel <= 1 Then
                    AuditDirectNestedInstance(parentFamilyDoc, fileName, hostFamily, hostCat, nestedFam, fi, nestedCat, nestedStructure, expectedByName, rows, nestedPath, nestingLevel)
                Else
                    AuditDescendantNestedInstance(fileName, hostFamily, hostCat, nestedFam, fi, nestedCat, nestedStructure, expectedByName, rows, nestedPath, nestingLevel)
                End If

                If childDoc IsNot Nothing AndAlso nestedStructure.HasNestedFamilies Then
                    AuditNestedInstancesRecursive(childDoc, hostFamily, hostCat, fileName, expectedByName, rows, nestedPath, nestingLevel + 1)
                End If

                If childDoc IsNot Nothing Then
                    Try
                        childDoc.Close(False)
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Shared Sub AuditDirectNestedInstance(parentFamilyDoc As Document,
                                                     fileName As String,
                                                     hostFamily As Family,
                                                     hostCat As String,
                                                     nestedFam As Family,
                                                     fi As FamilyInstance,
                                                     nestedCat As String,
                                                     nestedStructure As NestedFamilyStructureInfo,
                                                     expectedByName As Dictionary(Of String, List(Of FamilyLinkTargetParam)),
                                                     rows As List(Of FamilyLinkAuditRow),
                                                     nestedPath As String,
                                                     nestingLevel As Integer)

            Dim map As Dictionary(Of String, List(Of FoundParam)) = CollectParamMap(fi)

            For Each kv As KeyValuePair(Of String, List(Of FamilyLinkTargetParam)) In expectedByName
                Dim targetName As String = kv.Key
                Dim expectedItems As IEnumerable(Of FamilyLinkTargetParam) =
                    If(kv.Value, Enumerable.Empty(Of FamilyLinkTargetParam)()).
                        Where(Function(x) x IsNot Nothing)

                Dim resolvedFound As List(Of FoundParam) = Nothing
                If Not map.TryGetValue(targetName, resolvedFound) OrElse resolvedFound Is Nothing OrElse resolvedFound.Count = 0 Then
                    For Each expectedItem As FamilyLinkTargetParam In expectedItems
                        rows.Add(BuildParamNotFoundRow(fileName, hostFamily, hostCat, nestedFam, fi, nestedCat, targetName, expectedItem, nestedStructure, nestedPath, nestingLevel))
                    Next
                    Continue For
                End If

                For Each expectedItem As FamilyLinkTargetParam In expectedItems
                    Dim candidateRows As New List(Of FamilyLinkAuditRow)()

                    For Each fp As FoundParam In resolvedFound
                        Dim candidate As FamilyLinkAuditRow =
                            EvaluateFoundParam(fileName, hostFamily, hostCat, nestedFam, fi, nestedCat, targetName, expectedItem, fp, parentFamilyDoc, nestedStructure, nestedPath, nestingLevel)
                        If candidate IsNot Nothing Then candidateRows.Add(candidate)
                    Next

                    If candidateRows.Count = 0 Then
                        rows.Add(BuildParamNotFoundRow(fileName, hostFamily, hostCat, nestedFam, fi, nestedCat, targetName, expectedItem, nestedStructure, nestedPath, nestingLevel))
                        Continue For
                    End If

                    If candidateRows.Any(Function(x) IsOkIssue(x)) Then
                        Continue For
                    End If

                    Dim preferred As FamilyLinkAuditRow = SelectPreferredAuditRow(candidateRows)
                    If preferred IsNot Nothing Then rows.Add(preferred)
                Next
            Next
        End Sub

        Private Shared Sub AuditDescendantNestedInstance(fileName As String,
                                                         hostFamily As Family,
                                                         hostCat As String,
                                                         nestedFam As Family,
                                                         fi As FamilyInstance,
                                                         nestedCat As String,
                                                         nestedStructure As NestedFamilyStructureInfo,
                                                         expectedByName As Dictionary(Of String, List(Of FamilyLinkTargetParam)),
                                                         rows As List(Of FamilyLinkAuditRow),
                                                         nestedPath As String,
                                                         nestingLevel As Integer)

            Dim map As Dictionary(Of String, List(Of FoundParam)) = CollectParamMap(fi)

            For Each kv As KeyValuePair(Of String, List(Of FamilyLinkTargetParam)) In expectedByName
                Dim targetName As String = kv.Key
                Dim expectedItems As IEnumerable(Of FamilyLinkTargetParam) =
                    If(kv.Value, Enumerable.Empty(Of FamilyLinkTargetParam)()).
                        Where(Function(x) x IsNot Nothing)

                Dim found As List(Of FoundParam) = Nothing
                map.TryGetValue(targetName, found)
                Dim representative As FoundParam = If(found IsNot Nothing AndAlso found.Count > 0, found(0), Nothing)
                Dim p As Parameter = If(representative IsNot Nothing, representative.P, Nothing)

                Dim nestedParamName As String = ""
                Dim nestedScope As String = ""
                Dim nestedGuidStr As String = ""
                Dim nestedDataType As String = ""

                If p IsNot Nothing AndAlso p.Definition IsNot Nothing Then
                    nestedParamName = SafeStr(p.Definition.Name)
                    nestedScope = If(representative IsNot Nothing, representative.Scope.ToString(), "")
                    Dim nestedGuid As Guid
                    If TryGetParameterGuid(p, nestedGuid) Then
                        nestedGuidStr = nestedGuid.ToString("D")
                    End If
                    nestedDataType = SafeDefTypeToken(p.Definition)
                End If

                For Each expectedItem As FamilyLinkTargetParam In expectedItems
                    Dim notes As String =
                        $"직접 하위가 아닌 하위의 하위 경로(Level {nestingLevel})에서 발견된 패밀리입니다. 경로: {nestedPath}. 직접 하위만 호스트 연동 검토 대상이며 이 행은 동일 타입 중첩 구조를 구분하기 위한 표시입니다"

                    If String.IsNullOrWhiteSpace(nestedParamName) Then
                        notes &= " / 현재 인스턴스/타입에서 해당 이름의 파라미터를 찾지 못함"
                    End If

                    Dim compositeNote As String = BuildCompositeNestedNote(nestedStructure)
                    If compositeNote <> "" Then notes &= " / " & compositeNote

                    rows.Add(New FamilyLinkAuditRow With {
                        .FileName = fileName,
                        .HostFamilyName = If(hostFamily IsNot Nothing, hostFamily.Name, ""),
                        .HostFamilyCategory = hostCat,
                        .NestedFamilyName = If(nestedFam IsNot Nothing, nestedFam.Name, ""),
                        .NestedTypeName = SafeStr(If(fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing, fi.Symbol.Name, "")),
                        .NestedInstanceId = SafeElementIdText(fi),
                        .NestedPath = nestedPath,
                        .NestingLevel = nestingLevel.ToString(),
                        .NestedCategory = nestedCat,
                        .NestedParamName = nestedParamName,
                        .TargetParamName = targetName,
                        .ExpectedGuid = If(expectedItem IsNot Nothing, expectedItem.Guid.ToString("D"), ""),
                        .FoundScope = nestedScope,
                        .NestedParamGuid = nestedGuidStr,
                        .NestedParamDataType = nestedDataType,
                        .Issue = FamilyLinkAuditIssue.DescendantNestedUnsupported.ToString(),
                        .Notes = notes
                    })
                Next
            Next
        End Sub

        Private Shared Function BuildParamNotFoundRow(fileName As String,
                                                      hostFamily As Family,
                                                      hostCat As String,
                                                      nestedFam As Family,
                                                      fi As FamilyInstance,
                                                      nestedCat As String,
                                                      targetName As String,
                                                      expected As FamilyLinkTargetParam,
                                                      nestedStructure As NestedFamilyStructureInfo,
                                                      nestedPath As String,
                                                      nestingLevel As Integer) As FamilyLinkAuditRow
            Dim notes As String = "중첩 패밀리(하위) 또는 패밀리 인스턴스/타입에서 해당 이름의 파라미터를 찾지 못함"
            Dim compositeNote As String = BuildCompositeNestedNote(nestedStructure)
            If compositeNote <> "" Then notes &= " / " & compositeNote

            Return New FamilyLinkAuditRow With {
                .FileName = fileName,
                .HostFamilyName = If(hostFamily IsNot Nothing, hostFamily.Name, ""),
                .HostFamilyCategory = hostCat,
                .NestedFamilyName = If(nestedFam IsNot Nothing, nestedFam.Name, ""),
                .NestedTypeName = SafeStr(If(fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing, fi.Symbol.Name, "")),
                .NestedInstanceId = SafeElementIdText(fi),
                .NestedPath = nestedPath,
                .NestingLevel = nestingLevel.ToString(),
                .NestedCategory = nestedCat,
                .NestedParamName = "",
                .TargetParamName = targetName,
                .ExpectedGuid = If(expected IsNot Nothing, expected.Guid.ToString("D"), ""),
                .Issue = FamilyLinkAuditIssue.ParamNotFound.ToString(),
                .Notes = notes
            }
        End Function

        Private Shared Function EvaluateFoundParam(fileName As String,
                                                   hostFamily As Family,
                                                   hostCat As String,
                                                   nestedFam As Family,
                                                   fi As FamilyInstance,
                                                   nestedCat As String,
                                                   targetName As String,
                                                   expected As FamilyLinkTargetParam,
                                                   fp As FoundParam,
                                                   famDoc As Document,
                                                   nestedStructure As NestedFamilyStructureInfo,
                                                   nestedPath As String,
                                                   nestingLevel As Integer) As FamilyLinkAuditRow
            If expected Is Nothing OrElse fp Is Nothing Then Return Nothing

            Dim p As Parameter = fp.P
            If p Is Nothing OrElse p.Definition Is Nothing Then Return Nothing

            Dim nestedGuid As Guid
            Dim nestedGuidOk As Boolean = TryGetParameterGuid(p, nestedGuid)
            Dim nestedGuidStr As String = If(nestedGuidOk, nestedGuid.ToString("D"), "")

            Dim nestedIsShared As Boolean = False
            Dim nestedIsSharedKnown As Boolean = TryGetParameterIsShared(p, nestedIsShared)

            Dim assoc As FamilyParameter = Nothing
            Try
                assoc = famDoc.FamilyManager.GetAssociatedFamilyParameter(p)
            Catch ex As Exception
                Dim errorNotes As String = "GetAssociatedFamilyParameter 실패: " & ex.Message
                Dim compositeNote As String = BuildCompositeNestedNote(nestedStructure)
                If compositeNote <> "" Then errorNotes = compositeNote & " / " & errorNotes

                Return New FamilyLinkAuditRow With {
                    .FileName = fileName,
                    .HostFamilyName = If(hostFamily IsNot Nothing, hostFamily.Name, ""),
                    .HostFamilyCategory = hostCat,
                    .NestedFamilyName = If(nestedFam IsNot Nothing, nestedFam.Name, ""),
                    .NestedTypeName = SafeStr(If(fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing, fi.Symbol.Name, "")),
                    .NestedInstanceId = SafeElementIdText(fi),
                    .NestedPath = nestedPath,
                    .NestingLevel = nestingLevel.ToString(),
                    .NestedCategory = nestedCat,
                    .NestedParamName = SafeStr(p.Definition.Name),
                    .TargetParamName = targetName,
                    .ExpectedGuid = expected.Guid.ToString("D"),
                    .FoundScope = fp.Scope.ToString(),
                    .NestedParamGuid = nestedGuidStr,
                    .NestedParamDataType = SafeDefTypeToken(p.Definition),
                    .Issue = FamilyLinkAuditIssue.[Error].ToString(),
                    .Notes = errorNotes
                }
            End Try

            Dim issue As FamilyLinkAuditIssue = FamilyLinkAuditIssue.OK
            Dim notes As String = ""

            If assoc Is Nothing Then
                issue = FamilyLinkAuditIssue.MissingAssociation
                notes = "중첩 패밀리 파라미터에 호스트 연결(Associate)이 없습니다"
            Else
                If nestedGuidOk Then
                    If nestedGuid <> expected.Guid Then
                        issue = FamilyLinkAuditIssue.GuidMismatch
                        notes = $"중첩 패밀리 파라미터 GUID 불일치(Expected {expected.Guid:D}, Nested {nestedGuid:D})"
                    End If
                Else
                    If nestedIsSharedKnown Then
                        If nestedIsShared Then
                            notes = "중첩 패밀리 파라미터 IsShared=True 이나 GUID 추출 실패(특이 케이스)"
                        Else
                            notes = "중첩 패밀리 파라미터 IsShared=False (Shared 아님, 이름만 일치)"
                        End If
                    Else
                        notes = "중첩 패밀리 파라미터 Shared 여부 확인 실패(이름만 일치)"
                    End If
                End If

                If assoc.IsShared = False Then
                    If issue = FamilyLinkAuditIssue.OK Then issue = FamilyLinkAuditIssue.HostParamNotShared
                    If notes <> "" Then notes &= " / "
                    notes &= "연결된 호스트 FamilyParameter가 Shared가 아님"
                End If

                Dim hostGuid As Guid
                If TryGetDefinitionGuid(assoc.Definition, hostGuid) Then
                    If hostGuid <> expected.Guid Then
                        If issue = FamilyLinkAuditIssue.OK Then issue = FamilyLinkAuditIssue.GuidMismatch
                        If notes <> "" Then notes &= " / "
                        notes &= $"호스트 파라미터 GUID 불일치(Expected {expected.Guid:D}, Host {hostGuid:D})"
                    End If
                End If
            End If

            Dim compositeSummaryNote As String = BuildCompositeNestedNote(nestedStructure)
            If compositeSummaryNote <> "" Then
                If notes <> "" Then
                    notes &= " / "
                End If
                notes &= compositeSummaryNote
            End If

            Dim row As New FamilyLinkAuditRow With {
                .FileName = fileName,
                .HostFamilyName = If(hostFamily IsNot Nothing, hostFamily.Name, ""),
                .HostFamilyCategory = hostCat,
                .NestedFamilyName = If(nestedFam IsNot Nothing, nestedFam.Name, ""),
                .NestedTypeName = SafeStr(If(fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing, fi.Symbol.Name, "")),
                .NestedInstanceId = SafeElementIdText(fi),
                .NestedPath = nestedPath,
                .NestingLevel = nestingLevel.ToString(),
                .NestedCategory = nestedCat,
                .NestedParamName = SafeStr(p.Definition.Name),
                .TargetParamName = targetName,
                .ExpectedGuid = expected.Guid.ToString("D"),
                .FoundScope = fp.Scope.ToString(),
                .NestedParamGuid = nestedGuidStr,
                .NestedParamDataType = SafeDefTypeToken(p.Definition),
                .Issue = issue.ToString(),
                .Notes = notes
            }

            If assoc IsNot Nothing Then
                row.AssocHostParamName = SafeStr(assoc.Definition.Name)
                row.HostParamIsShared = assoc.IsShared.ToString()

                Dim hostGuid2 As Guid
                If TryGetDefinitionGuid(assoc.Definition, hostGuid2) Then
                    row.HostParamGuid = hostGuid2.ToString("D")
                End If
            End If

            Return row
        End Function

        Private Shared Function BuildCompositeNestedNote(info As NestedFamilyStructureInfo) As String
            If info Is Nothing OrElse Not info.HasNestedFamilies Then Return ""

            If info.NestedFamilyCount > 0 Then
                Return $"직접 중첩 패밀리가 또 다른 하위 패밀리 {info.NestedFamilyCount}개를 포함하는 복합 패밀리입니다. 하위 패밀리의 하위 패밀리 파라미터는 현재 단계에서 직접 연결 검토할 수 없습니다"
            End If

            Return "직접 중첩 패밀리가 또 다른 하위 패밀리를 포함하는 복합 패밀리입니다. 하위 패밀리의 하위 패밀리 파라미터는 현재 단계에서 직접 연결 검토할 수 없습니다"
        End Function

        Private Shared Function BuildNestedPath(pathPrefix As String,
                                                nestedFam As Family,
                                                fi As FamilyInstance) As String
            Dim segment As String = If(nestedFam IsNot Nothing, nestedFam.Name, "")
            Dim typeName As String = SafeStr(If(fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing, fi.Symbol.Name, ""))
            If typeName <> "" Then
                segment &= $" [{typeName}]"
            End If

            If String.IsNullOrWhiteSpace(pathPrefix) Then Return segment
            If String.IsNullOrWhiteSpace(segment) Then Return pathPrefix
            Return pathPrefix & " > " & segment
        End Function

        Private Shared Function IsOkIssue(row As FamilyLinkAuditRow) As Boolean
            If row Is Nothing Then Return False
            Return String.Equals(SafeStr(row.Issue).Trim(), FamilyLinkAuditIssue.OK.ToString(), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function SelectPreferredAuditRow(candidates As IEnumerable(Of FamilyLinkAuditRow)) As FamilyLinkAuditRow
            If candidates Is Nothing Then Return Nothing

            Return candidates.
                Where(Function(x) x IsNot Nothing).
                OrderBy(Function(x) GetIssuePriority(x)).
                ThenBy(Function(x) GetScopePriority(x)).
                ThenBy(Function(x) SafeStr(x.NestedParamName), StringComparer.OrdinalIgnoreCase).
                FirstOrDefault()
        End Function

        Private Shared Function GetIssuePriority(row As FamilyLinkAuditRow) As Integer
            If row Is Nothing Then Return Integer.MaxValue

            Select Case SafeStr(row.Issue).Trim()
                Case FamilyLinkAuditIssue.GuidMismatch.ToString()
                    Return 0
                Case FamilyLinkAuditIssue.HostParamNotShared.ToString()
                    Return 1
                Case FamilyLinkAuditIssue.MissingAssociation.ToString()
                    Return 2
                Case FamilyLinkAuditIssue.DescendantNestedUnsupported.ToString()
                    Return 3
                Case FamilyLinkAuditIssue.ParamNotFound.ToString()
                    Return 4
                Case FamilyLinkAuditIssue.[Error].ToString()
                    Return 5
                Case FamilyLinkAuditIssue.OK.ToString()
                    Return 6
                Case Else
                    Return 7
            End Select
        End Function

        Private Shared Function GetScopePriority(row As FamilyLinkAuditRow) As Integer
            If row Is Nothing Then Return Integer.MaxValue

            Select Case SafeStr(row.FoundScope).Trim()
                Case FoundScope.InstanceParam.ToString()
                    Return 0
                Case FoundScope.TypeParam.ToString()
                    Return 1
                Case Else
                    Return 2
            End Select
        End Function

        Private Shared Function CollectParamMap(fi As FamilyInstance) As Dictionary(Of String, List(Of FoundParam))
            Dim map As New Dictionary(Of String, List(Of FoundParam))(StringComparer.OrdinalIgnoreCase)

            Try
                For Each p As Parameter In fi.Parameters
                    If p Is Nothing OrElse p.Definition Is Nothing Then Continue For
                    Dim name As String = p.Definition.Name
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    If Not map.ContainsKey(name) Then map(name) = New List(Of FoundParam)()
                    map(name).Add(New FoundParam With {.P = p, .Scope = FoundScope.InstanceParam})
                Next
            Catch
            End Try

            Try
                If fi.Symbol IsNot Nothing Then
                    For Each p As Parameter In fi.Symbol.Parameters
                        If p Is Nothing OrElse p.Definition Is Nothing Then Continue For
                        Dim name As String = p.Definition.Name
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        If Not map.ContainsKey(name) Then map(name) = New List(Of FoundParam)()
                        map(name).Add(New FoundParam With {.P = p, .Scope = FoundScope.TypeParam})
                    Next
                End If
            Catch
            End Try

            Return map
        End Function

        Private Shared Function TryGetDefinitionGuid(defn As Definition, ByRef guid As Guid) As Boolean
            guid = Guid.Empty
            Try
                Dim ext As ExternalDefinition = TryCast(defn, ExternalDefinition)
                If ext IsNot Nothing Then
                    guid = ext.GUID
                    If guid <> Guid.Empty Then Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function TryGetParameterIsShared(p As Parameter, ByRef isShared As Boolean) As Boolean
            isShared = False
            If p Is Nothing Then Return False
            Try
                Dim t As Type = p.GetType()
                Dim prop As Reflection.PropertyInfo = t.GetProperty("IsShared")
                If prop Is Nothing Then Return False
                Dim v As Object = prop.GetValue(p, Nothing)
                If TypeOf v Is Boolean Then
                    isShared = CBool(v)
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Shared Function TryGetParameterGuid(p As Parameter, ByRef guid As Guid) As Boolean
            guid = Guid.Empty
            If p Is Nothing Then Return False

            Try
                Dim isShared As Boolean = False
                Dim isSharedKnown As Boolean = TryGetParameterIsShared(p, isShared)

                If isSharedKnown AndAlso isShared Then
                    Dim t As Type = p.GetType()

                    Dim propGuid As System.Reflection.PropertyInfo = t.GetProperty("GUID")
                    If propGuid Is Nothing Then propGuid = t.GetProperty("Guid")

                    If propGuid IsNot Nothing Then
                        Dim v As Object = propGuid.GetValue(p, Nothing)
                        If TypeOf v Is Guid Then
                            guid = CType(v, Guid)
                            If guid <> Guid.Empty Then Return True
                        End If
                    End If
                End If

            Catch
                ' Reflection/특이 케이스는 무시하고 다음 경로로 진행
            End Try

            Return TryGetDefinitionGuid(p.Definition, guid)
        End Function

        Private Shared Function SafeDefTypeToken(defn As Definition) As String
            If defn Is Nothing Then Return ""
            Try
#If REVIT2023 = 1 Or REVIT2025 = 1 Then
                Dim dt As ForgeTypeId = defn.GetDataType()
                If dt IsNot Nothing Then Return SafeStr(dt.TypeId)
                Return ""
#Else
                Return SafeStr(defn.ParameterType.ToString())
#End If
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeElementIdText(fi As FamilyInstance) As String
            If fi Is Nothing Then Return ""
            Try
                If fi.Id Is Nothing Then Return ""
                Return fi.Id.IntegerValue.ToString()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function OpenProjectDocument(app As Autodesk.Revit.ApplicationServices.Application, rvtPath As String) As Document
            If String.IsNullOrWhiteSpace(rvtPath) Then Throw New ArgumentException("path is empty.")
            Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtPath)

            Dim opts As New OpenOptions()
            opts.Audit = False

            Try
                opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            Catch
            End Try

            Try
                Dim ws As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                opts.SetOpenWorksetsConfiguration(ws)
            Catch
            End Try

            Try
                Return app.OpenDocumentFile(mp, opts)
            Catch
                Dim opts2 As New OpenOptions()
                opts2.Audit = False
                Try
                    Dim ws2 As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                    opts2.SetOpenWorksetsConfiguration(ws2)
                Catch
                End Try
                Return app.OpenDocumentFile(mp, opts2)
            End Try
        End Function

        Private Shared Function BuildTargetMap(targets As IList(Of FamilyLinkTargetParam)) As Dictionary(Of String, List(Of FamilyLinkTargetParam))
            Dim map As New Dictionary(Of String, List(Of FamilyLinkTargetParam))(StringComparer.OrdinalIgnoreCase)
            If targets Is Nothing Then Return map
            For Each t In targets
                If t Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(t.Name) Then Continue For
                Dim key As String = t.Name.Trim()
                If Not map.ContainsKey(key) Then map(key) = New List(Of FamilyLinkTargetParam)()
                If map(key).Any(Function(existing) existing IsNot Nothing AndAlso existing.Guid = t.Guid) Then Continue For
                map(key).Add(t)
            Next
            Return map
        End Function

        Private Shared Function NormalizePaths(paths As IList(Of String)) As List(Of String)
            Dim list As New List(Of String)()
            Dim dedup As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If paths Is Nothing Then Return list

            For Each p In paths
                If String.IsNullOrWhiteSpace(p) Then Continue For
                Dim full As String = p
                Try
                    full = Path.GetFullPath(p)
                Catch
                    full = p
                End Try
                If dedup.Add(full) Then list.Add(full)
            Next
            Return list
        End Function

        Private Shared Function SafeFileName(rvtPath As String) As String
            If String.IsNullOrWhiteSpace(rvtPath) Then Return "(Unknown)"
            Try
                Return Path.GetFileName(rvtPath)
            Catch
                Return rvtPath
            End Try
        End Function

        Private Shared Sub ReportProgress(progress As Action(Of Integer, String), total As Integer, index As Integer, fileProgress As Double, message As String)
            If progress Is Nothing Then Return
            Try
                Dim ratio As Double = 1.0R
                If total > 0 Then
                    Dim clamped As Double = Math.Max(0.0R, Math.Min(1.0R, fileProgress))
                    ratio = (index - 1 + clamped) / total
                End If
                Dim pct As Integer = CInt(Math.Max(0, Math.Min(100, Math.Round(ratio * 100))))
                progress(pct, message)
            Catch
            End Try
        End Sub

        Private Shared Function SafeRatio(current As Integer, total As Integer) As Double
            If total <= 0 Then Return 1.0R
            Return Math.Max(0.0R, Math.Min(1.0R, CDbl(current) / CDbl(total)))
        End Function

        Private Shared Function SafeStr(s As String) As String
            Return If(s, "")
        End Function

    End Class

End Namespace
