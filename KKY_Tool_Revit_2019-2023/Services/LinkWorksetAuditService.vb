Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB

Namespace Services

    Public Class LinkWorksetAuditRow
        Public Property HostFileName As String = ""
        Public Property HostFilePath As String = ""
        Public Property LinkName As String = ""
        Public Property LinkPath As String = ""
        Public Property AttachmentType As String = ""
        Public Property IsNestedLink As Boolean
        Public Property WasLoadedBefore As Boolean
        Public Property IsLoadedAfter As Boolean
        Public Property IsWorkshared As Boolean
        Public Property DefaultWorksetName As String = ""
        Public Property TotalUserWorksets As Integer
        Public Property OpenUserWorksetNamesBefore As String = ""
        Public Property OpenUserWorksetCountBefore As Integer
        Public Property DefaultOnlyOpenBefore As Boolean?
        Public Property OpenUserWorksetNamesAfter As String = ""
        Public Property OpenUserWorksetCountAfter As Integer
        Public Property DefaultOnlyOpenAfter As Boolean?
        Public Property ApplyRequested As Boolean
        Public Property Applied As Boolean
        Public Property Status As String = ""
        Public Property Message As String = ""
        Public Property DiagnosticLog As String = ""
    End Class

    Friend Class LinkWorksetState
        Public Property IsKnown As Boolean
        Public Property IsLoaded As Boolean
        Public Property IsWorkshared As Boolean
        Public Property DefaultWorksetId As WorksetId = WorksetId.InvalidWorksetId
        Public Property DefaultWorksetName As String = ""
        Public Property TotalUserWorksets As Integer
        Public Property OpenUserWorksetNames As List(Of String) = New List(Of String)()
        Public Property OpenUserWorksetCount As Integer
        Public Property DefaultOnlyOpen As Boolean?
    End Class

    Public NotInheritable Class LinkWorksetAuditService

        Private Sub New()
        End Sub

        Public Shared Function RunOnDocument(hostDoc As Document,
                                             hostPath As String,
                                             applyDefaultWorksetOnly As Boolean,
                                             progress As Action(Of Integer, String)) As List(Of LinkWorksetAuditRow)

            Dim rows As New List(Of LinkWorksetAuditRow)()
            If hostDoc Is Nothing Then Return rows

            Dim hostFileName As String = SafeHostFileName(hostDoc, hostPath)
            Dim linkTypes As List(Of RevitLinkType) =
                New FilteredElementCollector(hostDoc).
                    OfClass(GetType(RevitLinkType)).
                    Cast(Of RevitLinkType)().
                    Where(Function(x) x IsNot Nothing AndAlso Not x.IsNestedLink).
                    OrderBy(Function(x) SafeStr(x.Name), StringComparer.OrdinalIgnoreCase).
                    ToList()

            If linkTypes.Count = 0 Then
                ReportProgress(progress, 1, 1, $"[{hostFileName}] 링크 파일이 없습니다.")
                Return rows
            End If

            For i As Integer = 0 To linkTypes.Count - 1
                Dim linkType = linkTypes(i)
                Dim linkName As String = SafeStr(linkType.Name)
                ReportProgress(progress, linkTypes.Count, i + 1, $"[{hostFileName}] 링크 점검 중 ({i + 1}/{linkTypes.Count}) · {linkName}")
                rows.Add(AuditLinkType(hostDoc, hostPath, hostFileName, linkType, applyDefaultWorksetOnly))
            Next

            ReportProgress(progress, linkTypes.Count, linkTypes.Count, $"[{hostFileName}] 링크 점검 완료")
            Return rows
        End Function

        Private Shared Function AuditLinkType(hostDoc As Document,
                                              hostPath As String,
                                              hostFileName As String,
                                              linkType As RevitLinkType,
                                              applyDefaultWorksetOnly As Boolean) As LinkWorksetAuditRow

            Dim row As New LinkWorksetAuditRow With {
                .HostFileName = hostFileName,
                .HostFilePath = SafeStr(hostPath),
                .LinkName = SafeStr(linkType.Name),
                .ApplyRequested = applyDefaultWorksetOnly
            }
            Dim diag As New List(Of String)()
            diag.Add("link=" & row.LinkName)

            Try
                row.IsNestedLink = linkType.IsNestedLink
            Catch
                row.IsNestedLink = False
            End Try

            Try
                row.AttachmentType = linkType.AttachmentType.ToString()
            Catch
                row.AttachmentType = ""
            End Try

            Dim modelPath As ModelPath = Nothing
            Dim resourceRef As ExternalResourceReference = Nothing
            Dim storedLinkPath As String = ""
            Try
                Dim fileRef = linkType.GetExternalFileReference()
                If fileRef IsNot Nothing Then
                    Try
                        storedLinkPath = SafeModelPathToUserVisiblePath(fileRef.GetPath())
                    Catch
                        storedLinkPath = ""
                    End Try

                    ' Revit stores many local links as relative paths; use the API's
                    ' absolute resolution first so LoadFrom targets the real file.
                    Try
                        modelPath = fileRef.GetAbsolutePath()
                    Catch
                        modelPath = Nothing
                    End Try

                    If modelPath Is Nothing Then
                        modelPath = fileRef.GetPath()
                    End If

                    row.LinkPath = SafeModelPathToUserVisiblePath(modelPath)
                End If
            Catch
                modelPath = Nothing
            End Try
            Try
                Dim resourceRefs = linkType.GetExternalResourceReferences()
                If resourceRefs IsNot Nothing Then
                    For Each pair In resourceRefs
                        If pair.Value IsNot Nothing Then
                            resourceRef = pair.Value
                            Exit For
                        End If
                    Next
                End If
            Catch
                resourceRef = Nothing
            End Try
            Dim resolvedLinkPath As String = ResolveLinkPath(hostPath, row.LinkPath)
            If Not String.IsNullOrWhiteSpace(resolvedLinkPath) AndAlso Not String.Equals(resolvedLinkPath, row.LinkPath, StringComparison.OrdinalIgnoreCase) Then
                row.LinkPath = resolvedLinkPath
                Try
                    modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(resolvedLinkPath)
                Catch
                End Try
            End If
            If Not String.IsNullOrWhiteSpace(storedLinkPath) AndAlso Not String.Equals(storedLinkPath, row.LinkPath, StringComparison.OrdinalIgnoreCase) Then
                diag.Add("storedPath=" & storedLinkPath)
            End If
            diag.Add("path=" & If(String.IsNullOrWhiteSpace(row.LinkPath), "(empty)", row.LinkPath))

            Dim beforeState = InspectCurrentLinkState(hostDoc, linkType.Id)
            ApplyStateBefore(row, beforeState)
            diag.Add("before=" & DescribeState(beforeState))

            Dim previewState = PreviewLinkState(modelPath)
            MergeWorksetMetadata(row, beforeState, previewState)
            diag.Add("preview=" & DescribeState(previewState))

            If Not row.IsWorkshared Then
                row.Status = "n/a"
                row.IsLoadedAfter = row.WasLoadedBefore
                row.Message = "Workshared 링크가 아니어서 Manage Worksets 대상이 아닙니다."
                Return FinalizeAuditRow(row, diag)
            End If

            Dim stateForApply = beforeState
            If applyDefaultWorksetOnly AndAlso (Not stateForApply.IsKnown OrElse Not stateForApply.IsLoaded) Then
                stateForApply = PrimeLinkStateForApply(hostDoc, linkType, stateForApply, diag)
                MergeWorksetMetadata(row, stateForApply, previewState)
                diag.Add("applyState=" & DescribeState(stateForApply))
            End If

            Dim needsApply As Boolean =
                applyDefaultWorksetOnly AndAlso
                (Not stateForApply.IsKnown OrElse Not stateForApply.DefaultOnlyOpen.HasValue OrElse stateForApply.DefaultOnlyOpen.Value = False)
            diag.Add("needsApply=" & BoolToken(needsApply))
            If needsApply Then
                If modelPath Is Nothing Then
                    row.Status = "error"
                    row.IsLoadedAfter = row.WasLoadedBefore
                    row.Message = "링크 경로를 확인하지 못해 기본 웍셋 적용을 수행할 수 없습니다."
                    Return FinalizeAuditRow(row, diag)
                End If

                Dim defaultWorksetId As WorksetId = ResolveDefaultWorksetId(stateForApply, previewState)
                diag.Add("selectedDefaultId=" & WorksetIdText(defaultWorksetId))
                diag.Add("selectedDefaultName=" & row.DefaultWorksetName)
                If defaultWorksetId Is Nothing OrElse defaultWorksetId = WorksetId.InvalidWorksetId Then
                    row.Status = "error"
                    row.IsLoadedAfter = row.WasLoadedBefore
                    row.Message = "기본 user workset을 찾지 못해 적용할 수 없습니다."
                    Return FinalizeAuditRow(row, diag)
                End If

                Dim config As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                config.Open(New List(Of WorksetId) From {defaultWorksetId})

                Try
                    Dim checkoutError As String = ""
                    Dim hostWorksetNames As String = ""
                    If Not EnsureEditableHostWorksets(hostDoc, linkType.Id, checkoutError, hostWorksetNames) Then
                        diag.Add("hostCheckout=N")
                        diag.Add("hostWorksets=" & hostWorksetNames)
                        diag.Add("hostCheckoutError=" & checkoutError)
                        row.Status = "error"
                        row.IsLoadedAfter = row.WasLoadedBefore
                        row.Message = If(String.IsNullOrWhiteSpace(checkoutError),
                                         "링크가 들어있는 호스트 workset을 editable 상태로 확보하지 못했습니다.",
                                         checkoutError)
                        Return FinalizeAuditRow(row, diag)
                    End If
                    diag.Add("hostCheckout=Y")
                    diag.Add("hostWorksets=" & hostWorksetNames)

                    Dim loadResult As LinkLoadResult = Nothing
                    If resourceRef IsNot Nothing Then
                        diag.Add("loadSource=ExternalResourceReference")
                        loadResult = linkType.LoadFrom(resourceRef, config)
                    Else
                        diag.Add("loadSource=ModelPath")
                        loadResult = linkType.LoadFrom(modelPath, config)
                    End If
                    Dim loadSucceeded As Boolean = True
                    Dim loadResultText As String = ""
                    Try
                        loadResultText = SafeStr(loadResult.LoadResult.ToString())
                        loadSucceeded = (loadResult.LoadResult = LinkLoadResultType.LinkLoaded)
                    Catch
                        loadSucceeded = True
                    End Try
                    diag.Add("loadResult=" & If(String.IsNullOrWhiteSpace(loadResultText), "(unknown)", loadResultText))
                    If Not loadSucceeded Then
                        row.Status = "error"
                        row.IsLoadedAfter = False
                        row.Message = If(String.IsNullOrWhiteSpace(loadResultText),
                                         "기본 웍셋 적용 후 링크를 다시 로드하지 못했습니다.",
                                         $"기본 웍셋 적용 후 링크를 다시 로드하지 못했습니다. ({loadResultText})")
                        Return FinalizeAuditRow(row, diag)
                    End If
                    row.Applied = True
                Catch ex As Exception
                    diag.Add("exception=" & SafeStr(ex.Message))
                    row.Status = "error"
                    row.IsLoadedAfter = row.WasLoadedBefore
                    row.Message = $"기본 웍셋 적용 실패: {ex.Message}"
                    Return FinalizeAuditRow(row, diag)
                End Try
            End If

            If row.Applied Then
                row.IsLoadedAfter = True
                row.OpenUserWorksetCountAfter = 1
                If String.IsNullOrWhiteSpace(row.DefaultWorksetName) AndAlso previewState IsNot Nothing Then
                    row.DefaultWorksetName = SafeStr(previewState.DefaultWorksetName)
                End If
                row.OpenUserWorksetNamesAfter = SafeStr(row.DefaultWorksetName)
                row.DefaultOnlyOpenAfter = True
                diag.Add("after=post-load-probe-skipped")
                row.Status = "changed"
                row.Message = "기본 user workset만 열도록 링크를 다시 로드했습니다."
                Return FinalizeAuditRow(row, diag)
            End If

            Dim afterState = InspectCurrentLinkState(hostDoc, linkType.Id)
            ApplyStateAfter(row, afterState)
            MergeWorksetMetadata(row, afterState, previewState)
            diag.Add("after=" & DescribeState(afterState))
            If afterState.IsKnown Then
                If afterState.DefaultOnlyOpen = True Then
                    row.Status = If(row.Applied, "changed", "ok")
                    row.Message = If(row.Applied, "기본 workset1 만 열리도록 재로드했습니다.", "이미 기본 workset1 만 열려 있습니다.")
                Else
                    row.Status = "warning"
                    row.Message = "기본 user workset 외 다른 workset도 열려 있습니다."
                End If
            ElseIf row.Applied Then
                row.Status = "warning"
                row.Message = "적용은 수행했지만 재로드 후 링크 상태를 다시 확인하지 못했습니다."
            ElseIf row.WasLoadedBefore Then
                row.Status = "warning"
                row.Message = "링크가 로드되어 있었지만 workset 상태를 확인하지 못했습니다."
            Else
                row.Status = "skipped"
                row.Message = "현재 링크가 로드되어 있지 않아 open workset 상태를 확인하지 못했습니다."
            End If

            Return FinalizeAuditRow(row, diag)
        End Function

        Private Shared Function EnsureEditableHostWorksets(hostDoc As Document,
                                                           linkTypeId As ElementId,
                                                           ByRef errorMessage As String,
                                                           ByRef hostWorksetNames As String) As Boolean
            errorMessage = ""
            hostWorksetNames = ""
            If hostDoc Is Nothing Then Return True

            Try
                If Not hostDoc.IsWorkshared Then Return True
            Catch
                Return True
            End Try

            Dim requestedIds = CollectLinkHostWorksetIds(hostDoc, linkTypeId)
            If requestedIds.Count = 0 Then Return True
            hostWorksetNames = String.Join(", ", ResolveWorksetNames(hostDoc, requestedIds))

            Try
                Dim checkedOut = WorksharingUtils.CheckoutWorksets(hostDoc, requestedIds)
                If checkedOut Is Nothing Then
                    errorMessage = "링크가 배치된 호스트 workset checkout 결과를 확인하지 못했습니다."
                    Return False
                End If

                Dim missingIds As New List(Of WorksetId)()
                For Each worksetId In requestedIds
                    If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Continue For
                    If Not checkedOut.Contains(worksetId) Then
                        missingIds.Add(worksetId)
                    End If
                Next

                If missingIds.Count > 0 Then
                    errorMessage = "링크가 들어있는 호스트 workset을 checkout하지 못했습니다: " & String.Join(", ", ResolveWorksetNames(hostDoc, missingIds))
                    Return False
                End If

                Return True
            Catch ex As Exception
                errorMessage = $"호스트 workset checkout 실패: {ex.Message}"
                Return False
            End Try
        End Function

        Private Shared Function PrimeLinkStateForApply(hostDoc As Document,
                                                       linkType As RevitLinkType,
                                                       fallbackState As LinkWorksetState,
                                                       diag As IList(Of String)) As LinkWorksetState
            If linkType Is Nothing Then Return fallbackState

            Dim currentState = fallbackState
            If currentState IsNot Nothing AndAlso currentState.IsKnown AndAlso currentState.IsLoaded Then
                If diag IsNot Nothing Then diag.Add("preload=skipped-already-loaded")
                Return currentState
            End If

            Try
                If diag IsNot Nothing Then diag.Add("preload=attempt")
                Dim loadResult = linkType.Load()
                Dim loadResultText As String = ""
                Dim loadSucceeded As Boolean = False
                Try
                    loadResultText = SafeStr(loadResult.LoadResult.ToString())
                    loadSucceeded = (loadResult.LoadResult = LinkLoadResultType.LinkLoaded)
                Catch
                    loadSucceeded = False
                End Try

                If diag IsNot Nothing Then
                    diag.Add("preloadResult=" & If(String.IsNullOrWhiteSpace(loadResultText), "(unknown)", loadResultText))
                End If

                currentState = InspectCurrentLinkState(hostDoc, linkType.Id)
                If diag IsNot Nothing Then diag.Add("preloadState=" & DescribeState(currentState))
                If loadSucceeded OrElse (currentState IsNot Nothing AndAlso currentState.IsKnown AndAlso currentState.IsLoaded) Then
                    Return currentState
                End If
            Catch ex As Exception
                If diag IsNot Nothing Then diag.Add("preloadException=" & SafeStr(ex.Message))
                currentState = InspectCurrentLinkState(hostDoc, linkType.Id)
                If diag IsNot Nothing Then diag.Add("preloadState=" & DescribeState(currentState))
                If currentState IsNot Nothing AndAlso currentState.IsKnown AndAlso currentState.IsLoaded Then
                    Return currentState
                End If
            End Try

            Return fallbackState
        End Function

        Private Shared Function CollectLinkHostWorksetIds(hostDoc As Document, linkTypeId As ElementId) As ICollection(Of WorksetId)
            Dim ids As New HashSet(Of WorksetId)()
            If hostDoc Is Nothing OrElse linkTypeId Is Nothing Then Return ids

            Try
                Dim linkType = TryCast(hostDoc.GetElement(linkTypeId), RevitLinkType)
                If linkType IsNot Nothing AndAlso linkType.WorksetId IsNot Nothing AndAlso linkType.WorksetId <> WorksetId.InvalidWorksetId Then
                    ids.Add(linkType.WorksetId)
                End If
            Catch
            End Try

            Try
                Dim instances =
                    New FilteredElementCollector(hostDoc).
                        OfClass(GetType(RevitLinkInstance)).
                        WhereElementIsNotElementType().
                        Cast(Of RevitLinkInstance)().
                        Where(Function(inst) inst IsNot Nothing AndAlso inst.GetTypeId() = linkTypeId)

                For Each inst In instances
                    If inst Is Nothing Then Continue For
                    Try
                        Dim wsId = inst.WorksetId
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

        Private Shared Function ResolveWorksetNames(hostDoc As Document, worksetIds As IEnumerable(Of WorksetId)) As IEnumerable(Of String)
            Dim names As New List(Of String)()
            If hostDoc Is Nothing OrElse worksetIds Is Nothing Then Return names

            Dim table As WorksetTable = Nothing
            Try
                table = hostDoc.GetWorksetTable()
            Catch
                table = Nothing
            End Try

            For Each worksetId In worksetIds
                If worksetId Is Nothing OrElse worksetId = WorksetId.InvalidWorksetId Then Continue For
                Dim name As String = worksetId.IntegerValue.ToString()
                Try
                    If table IsNot Nothing Then
                        Dim ws = table.GetWorkset(worksetId)
                        If ws IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ws.Name) Then
                            name = ws.Name
                        End If
                    End If
                Catch
                End Try
                names.Add(name)
            Next

            Return names
        End Function

        Private Shared Function FinalizeAuditRow(row As LinkWorksetAuditRow, diag As IEnumerable(Of String)) As LinkWorksetAuditRow
            If row Is Nothing Then Return Nothing
            row.DiagnosticLog = String.Join(" || ", If(diag, Enumerable.Empty(Of String)()))
            Return row
        End Function

        Private Shared Function DescribeState(state As LinkWorksetState) As String
            If state Is Nothing Then Return "null"
            Return String.Join(";",
                               "known=" & BoolToken(state.IsKnown),
                               "loaded=" & BoolToken(state.IsLoaded),
                               "workshared=" & BoolToken(state.IsWorkshared),
                               "defaultId=" & WorksetIdText(state.DefaultWorksetId),
                               "defaultName=" & SafeStr(state.DefaultWorksetName),
                               "openCount=" & state.OpenUserWorksetCount.ToString(),
                               "defaultOnly=" & NullableBoolToken(state.DefaultOnlyOpen),
                               "openNames=" & SafeStr(String.Join(", ", state.OpenUserWorksetNames)))
        End Function

        Private Shared Function BoolToken(value As Boolean) As String
            Return If(value, "Y", "N")
        End Function

        Private Shared Function NullableBoolToken(value As Boolean?) As String
            If Not value.HasValue Then Return "NA"
            Return BoolToken(value.Value)
        End Function

        Private Shared Function WorksetIdText(worksetId As WorksetId) As String
            If worksetId Is Nothing Then Return "null"
            If worksetId = WorksetId.InvalidWorksetId Then Return "invalid"
            Try
                Return worksetId.IntegerValue.ToString()
            Catch
                Return "unknown"
            End Try
        End Function

        Private Shared Sub ApplyInferredDefaultOnlyAfter(row As LinkWorksetAuditRow,
                                                         primary As LinkWorksetState,
                                                         fallback As LinkWorksetState)
            If row Is Nothing Then Return

            Dim defaultWorksetName As String = row.DefaultWorksetName
            If String.IsNullOrWhiteSpace(defaultWorksetName) AndAlso primary IsNot Nothing Then
                defaultWorksetName = SafeStr(primary.DefaultWorksetName)
            End If
            If String.IsNullOrWhiteSpace(defaultWorksetName) AndAlso fallback IsNot Nothing Then
                defaultWorksetName = SafeStr(fallback.DefaultWorksetName)
            End If

            row.IsLoadedAfter = True
            row.OpenUserWorksetCountAfter = 1
            row.OpenUserWorksetNamesAfter = defaultWorksetName
            row.DefaultOnlyOpenAfter = True
            If String.IsNullOrWhiteSpace(row.DefaultWorksetName) Then
                row.DefaultWorksetName = defaultWorksetName
            End If
        End Sub

        Private Shared Function InspectCurrentLinkState(hostDoc As Document, linkTypeId As ElementId) As LinkWorksetState
            Dim state As New LinkWorksetState()

            Try
                state.IsLoaded = RevitLinkType.IsLoaded(hostDoc, linkTypeId)
            Catch
                state.IsLoaded = False
            End Try

            If Not state.IsLoaded Then Return state

            Dim linkDoc As Document = GetLoadedLinkDocument(hostDoc, linkTypeId)
            If linkDoc Is Nothing Then Return state

            Dim userWorksets As List(Of Workset) =
                New FilteredWorksetCollector(linkDoc).
                    OfKind(WorksetKind.UserWorkset).
                    ToWorksets().
                    ToList()

            state.IsKnown = True
            state.IsWorkshared = (userWorksets.Count > 0)
            state.TotalUserWorksets = userWorksets.Count

            Dim defaultWorkset = userWorksets.FirstOrDefault(Function(ws) ws IsNot Nothing AndAlso ws.IsDefaultWorkset)
            If defaultWorkset IsNot Nothing Then
                state.DefaultWorksetId = defaultWorkset.Id
                state.DefaultWorksetName = SafeStr(defaultWorkset.Name)
            End If

            state.OpenUserWorksetNames = userWorksets.
                Where(Function(ws) ws IsNot Nothing AndAlso ws.IsOpen).
                OrderBy(Function(ws) SafeStr(ws.Name), StringComparer.OrdinalIgnoreCase).
                Select(Function(ws) SafeStr(ws.Name)).
                ToList()

            state.OpenUserWorksetCount = state.OpenUserWorksetNames.Count

            If userWorksets.Count > 0 AndAlso defaultWorkset IsNot Nothing Then
                state.DefaultOnlyOpen = (defaultWorkset.IsOpen AndAlso state.OpenUserWorksetCount = 1)
            ElseIf userWorksets.Count > 0 Then
                state.DefaultOnlyOpen = False
            Else
                state.DefaultOnlyOpen = Nothing
            End If

            Return state
        End Function

        Private Shared Function PreviewLinkState(modelPath As ModelPath) As LinkWorksetState
            Dim state As New LinkWorksetState()
            If modelPath Is Nothing Then Return state

            Try
                Dim previews = WorksharingUtils.GetUserWorksetInfo(modelPath)
                If previews Is Nothing OrElse previews.Count = 0 Then Return state

                state.IsKnown = True
                state.IsWorkshared = True
                state.TotalUserWorksets = previews.Count

                Dim defaultPreview = previews.FirstOrDefault(Function(ws) ws IsNot Nothing AndAlso ws.IsDefaultWorkset)
                If defaultPreview IsNot Nothing Then
                    state.DefaultWorksetId = defaultPreview.Id
                    state.DefaultWorksetName = SafeStr(defaultPreview.Name)
                End If
            Catch
            End Try

            Return state
        End Function

        Private Shared Function ResolveDefaultWorksetId(primary As LinkWorksetState, fallback As LinkWorksetState) As WorksetId
            If fallback IsNot Nothing AndAlso fallback.DefaultWorksetId IsNot Nothing AndAlso fallback.DefaultWorksetId <> WorksetId.InvalidWorksetId Then
                Return fallback.DefaultWorksetId
            End If
            If primary IsNot Nothing AndAlso primary.DefaultWorksetId IsNot Nothing AndAlso primary.DefaultWorksetId <> WorksetId.InvalidWorksetId Then
                Return primary.DefaultWorksetId
            End If
            Return WorksetId.InvalidWorksetId
        End Function

        Private Shared Sub MergeWorksetMetadata(row As LinkWorksetAuditRow, primary As LinkWorksetState, fallback As LinkWorksetState)
            If row Is Nothing Then Return
            Dim target = If(primary, fallback)
            If primary IsNot Nothing AndAlso primary.IsWorkshared Then
                target = primary
            ElseIf fallback IsNot Nothing AndAlso fallback.IsWorkshared Then
                target = fallback
            End If
            If target Is Nothing Then Return

            row.IsWorkshared = target.IsWorkshared
            row.TotalUserWorksets = Math.Max(row.TotalUserWorksets, target.TotalUserWorksets)
            If String.IsNullOrWhiteSpace(row.DefaultWorksetName) AndAlso Not String.IsNullOrWhiteSpace(target.DefaultWorksetName) Then
                row.DefaultWorksetName = target.DefaultWorksetName
            End If
        End Sub

        Private Shared Sub ApplyStateBefore(row As LinkWorksetAuditRow, state As LinkWorksetState)
            If row Is Nothing OrElse state Is Nothing Then Return
            row.WasLoadedBefore = state.IsLoaded
            row.OpenUserWorksetCountBefore = state.OpenUserWorksetCount
            row.OpenUserWorksetNamesBefore = String.Join(", ", state.OpenUserWorksetNames)
            row.DefaultOnlyOpenBefore = state.DefaultOnlyOpen
        End Sub

        Private Shared Sub ApplyStateAfter(row As LinkWorksetAuditRow, state As LinkWorksetState)
            If row Is Nothing OrElse state Is Nothing Then Return
            row.IsLoadedAfter = state.IsLoaded
            row.OpenUserWorksetCountAfter = state.OpenUserWorksetCount
            row.OpenUserWorksetNamesAfter = String.Join(", ", state.OpenUserWorksetNames)
            row.DefaultOnlyOpenAfter = state.DefaultOnlyOpen
        End Sub

        Private Shared Function GetLoadedLinkDocument(hostDoc As Document, linkTypeId As ElementId) As Document
            Dim instances As IEnumerable(Of RevitLinkInstance) =
                New FilteredElementCollector(hostDoc).
                    OfClass(GetType(RevitLinkInstance)).
                    WhereElementIsNotElementType().
                    Cast(Of RevitLinkInstance)()

            For Each inst In instances
                If inst Is Nothing Then Continue For
                If inst.GetTypeId() <> linkTypeId Then Continue For
                Try
                    Dim linkDoc = inst.GetLinkDocument()
                    If linkDoc IsNot Nothing Then Return linkDoc
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Private Shared Sub ReportProgress(progress As Action(Of Integer, String), total As Integer, index As Integer, message As String)
            If progress Is Nothing Then Return
            Dim percent As Integer = 0
            If total > 0 Then
                percent = CInt(Math.Round((CDbl(index) / CDbl(total)) * 100.0R))
            End If
            progress(Math.Max(0, Math.Min(100, percent)), message)
        End Sub

        Private Shared Function SafeHostFileName(doc As Document, hostPath As String) As String
            Try
                If Not String.IsNullOrWhiteSpace(hostPath) Then
                    Dim name = Path.GetFileName(hostPath)
                    If Not String.IsNullOrWhiteSpace(name) Then Return name
                End If
            Catch
            End Try

            Try
                Return SafeStr(doc.Title)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function ResolveLinkPath(hostPath As String, linkPath As String) As String
            If String.IsNullOrWhiteSpace(linkPath) Then Return ""
            Try
                If Path.IsPathRooted(linkPath) Then Return linkPath
            Catch
            End Try

            Try
                If String.IsNullOrWhiteSpace(hostPath) Then Return linkPath
                Dim hostDir = Path.GetDirectoryName(hostPath)
                If String.IsNullOrWhiteSpace(hostDir) Then Return linkPath
                Dim candidate = Path.GetFullPath(Path.Combine(hostDir, linkPath))
                If File.Exists(candidate) Then Return candidate
            Catch
            End Try

            Return linkPath
        End Function

        Private Shared Function SafeModelPathToUserVisiblePath(modelPath As ModelPath) As String
            If modelPath Is Nothing Then Return ""
            Try
                Return ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath)
            Catch
                Try
                    Return SafeStr(modelPath.ToString())
                Catch
                    Return ""
                End Try
            End Try
        End Function

        Private Shared Function SafeStr(value As Object) As String
            If value Is Nothing Then Return ""
            Try
                Return Convert.ToString(value)
            Catch
                Return ""
            End Try
        End Function

    End Class

End Namespace
