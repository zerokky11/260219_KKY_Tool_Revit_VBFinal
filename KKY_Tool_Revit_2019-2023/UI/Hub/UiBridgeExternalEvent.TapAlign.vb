Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.Linq
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private _tapAlignRows As List(Of Dictionary(Of String, Object)) = Nothing
        Private _tapAlignLastUnit As String = "mm"
        Private _tapAlignLastDomain As String = "all"
        Private _tapAlignLastTol As Double = 0.5R
        Private _tapAlignLastExportLocale As String = "ko"
        Private _tapAlignLastTargetCount As Integer = 0
        Private _tapAlignExtraHeaders As List(Of String) = New List(Of String)()
        Private _tapAlignTargetFilter As String = String.Empty
        Private _tapAlignFeatureTargetFilter As String = String.Empty
        Private _tapAlignExcludeTargetFilter As String = String.Empty
        Private Const TapAlignProgressChannel As String = "tapalign:progress"

        Private Sub HandleTapAlignRun(app As UIApplication, payload As Object)
            Try
                ReportTapAlignProgress(0.05R, "탭/분기 축 틀어짐 검토를 준비하는 중...")

                Dim uidoc = app.ActiveUIDocument
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    SendToWeb("revit:error", New With {.message = "활성 문서가 없습니다."})
                    SendToWeb("tapalign:done", New With {
                        .ok = False,
                        .message = "활성 문서가 없습니다.",
                        .rows = New List(Of Dictionary(Of String, Object))(),
                        .unit = _tapAlignLastUnit,
                        .domain = _tapAlignLastDomain,
                        .extraHeaders = New List(Of String)()
                    })
                    Return
                End If

                Dim tol As Double = 0.5R
                Dim unit As String = "mm"
                Dim domain As String = "all"

                Try
                    Dim rawTol = GetProp(payload, "tol")
                    If rawTol IsNot Nothing Then tol = Convert.ToDouble(rawTol, CultureInfo.InvariantCulture)
                Catch
                    tol = 0.5R
                End Try

                Try
                    Dim rawUnit = TryCast(GetProp(payload, "unit"), String)
                    If Not String.IsNullOrWhiteSpace(rawUnit) Then unit = rawUnit
                Catch
                    unit = "mm"
                End Try

                Try
                    Dim rawDomain = TryCast(GetProp(payload, "domain"), String)
                    If Not String.IsNullOrWhiteSpace(rawDomain) Then domain = rawDomain
                Catch
                    domain = "all"
                End Try

                unit = NormalizeTapAlignUnit(unit)
                domain = NormalizeTapAlignDomain(domain)

                Dim commonOptions = ReadTapAlignCommonOptions(payload)
                _tapAlignFeatureTargetFilter = SafeTapAlignString(GetProp(payload, "featureTargetFilter"), String.Empty).Trim()
                _tapAlignExtraHeaders = BuildConnectorExtraParams(commonOptions.ExtraParamsText,
                                                                 commonOptions.IncludePointXY,
                                                                 commonOptions.IncludeLinearMetrics)
                _tapAlignTargetFilter = Services.TapAlignmentReviewService.CombineTargetFilterText(commonOptions.TargetFilterText,
                                                                                                   _tapAlignFeatureTargetFilter)
                _tapAlignExcludeTargetFilter = If(commonOptions.ExcludeTargetFilterText, String.Empty)

                _tapAlignLastTol = tol
                _tapAlignLastUnit = unit
                _tapAlignLastDomain = domain

                _tapAlignRows = Services.TapAlignmentReviewService.Run(app,
                                                                       tol,
                                                                       unit,
                                                                       domain,
                                                                       _tapAlignExtraHeaders,
                                                                       _tapAlignTargetFilter,
                                                                       _tapAlignExcludeTargetFilter,
                                                                       AddressOf ReportTapAlignProgress)

                If _tapAlignRows Is Nothing Then
                    _tapAlignRows = New List(Of Dictionary(Of String, Object))()
                End If

                _tapAlignLastTargetCount = Services.TapAlignmentReviewService.CountTargets(app,
                                                                                           domain,
                                                                                           _tapAlignTargetFilter,
                                                                                           _tapAlignExcludeTargetFilter)

                Dim fileCount =
                    _tapAlignRows.
                    Select(Function(row) ReadTapAlignField(row, "File")).
                    Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    Count()

                SendToWeb("tapalign:done", New With {
                    .ok = True,
                    .rows = _tapAlignRows,
                    .unit = unit,
                    .domain = domain,
                    .tol = tol,
                    .extraHeaders = _tapAlignExtraHeaders,
                    .featureTargetFilter = _tapAlignFeatureTargetFilter,
                    .common = New With {
                        .extraParamsText = commonOptions.ExtraParamsText,
                        .targetFilterText = commonOptions.TargetFilterText,
                        .excludeTargetFilterText = commonOptions.ExcludeTargetFilterText,
                        .excludeEndDummy = False,
                        .includePointXY = commonOptions.IncludePointXY,
                        .includeLinearMetrics = commonOptions.IncludeLinearMetrics
                    },
                    .summary = New With {
                        .issueCount = _tapAlignRows.Count,
                        .fileCount = fileCount,
                        .domainLabel = ResolveTapAlignDomainLabel(domain)
                    }
                })
            Catch ex As Exception
                LogError("[tapalign] " & ex.Message)
                SendToWeb("revit:error", New With {.message = ex.Message})
                SendToWeb("tapalign:done", New With {
                    .ok = False,
                    .message = ex.Message,
                    .rows = New List(Of Dictionary(Of String, Object))(),
                    .unit = _tapAlignLastUnit,
                    .domain = _tapAlignLastDomain,
                    .extraHeaders = _tapAlignExtraHeaders
                })
            End Try
        End Sub

        Private Sub HandleTapAlignSaveExcel(app As UIApplication, payload As Object)
            Try
                If _tapAlignRows Is Nothing Then
                    _tapAlignRows = New List(Of Dictionary(Of String, Object))()
                End If

                Dim unit As String = _tapAlignLastUnit
                Dim locale As String = _tapAlignLastExportLocale

                Try
                    Dim rawUnit = TryCast(GetProp(payload, "unit"), String)
                    If Not String.IsNullOrWhiteSpace(rawUnit) Then unit = rawUnit
                Catch
                    unit = _tapAlignLastUnit
                End Try

                Try
                    Dim rawLocale = TryCast(GetProp(payload, "locale"), String)
                    If Not String.IsNullOrWhiteSpace(rawLocale) Then locale = rawLocale
                Catch
                    locale = _tapAlignLastExportLocale
                End Try

                unit = NormalizeTapAlignUnit(unit)
                locale = NormalizeTapAlignExportLocale(locale)
                _tapAlignLastExportLocale = locale
                Dim doAutoFit As Boolean = ParseExcelMode(payload)

                Dim table = BuildTapAlignDataTable(_tapAlignRows, unit, _tapAlignExtraHeaders, locale, _tapAlignLastTargetCount)
                Dim totalRows = Math.Max(1, table.Rows.Count)
                Dim defaultName = BridgeHandler.SanitizeFileName(String.Format(CultureInfo.InvariantCulture,
                                                                               "{0}_탭분기축검토_{1}건.xlsx",
                                                                               Date.Now.ToString("yyMMdd", CultureInfo.InvariantCulture),
                                                                               Math.Max(0, _tapAlignRows.Count)))
                ExcelProgressReporter.Reset(TapAlignProgressChannel)
                ExcelProgressReporter.Report(TapAlignProgressChannel, "EXCEL_INIT", "엑셀 워크북 준비", 0, totalRows, Nothing, True)
                Dim savedPath = ExcelCore.PickAndSaveXlsx("탭/분기 축 틀어짐 검토", table, defaultName, autoFit:=doAutoFit, progressKey:=TapAlignProgressChannel, exportKind:="tapalign", exportLocale:=locale)

                If String.IsNullOrWhiteSpace(savedPath) Then
                    SendToWeb("tapalign:saved", New With {.ok = False, .cancelled = True})
                    Return
                End If

                SendToWeb("tapalign:saved", New With {.ok = True, .path = savedPath})
            Catch ex As Exception
                ExcelProgressReporter.Report(TapAlignProgressChannel, "ERROR", ex.Message, 0, 0, Nothing, True)
                SendToWeb("tapalign:saved", New With {.ok = False, .message = ex.Message})
            End Try
        End Sub

        Private Sub ReportTapAlignProgress(pct As Double, text As String)
            Dim detail = If(text, String.Empty)
            SendToWeb(TapAlignProgressChannel, New With {
                .pct = pct,
                .text = text,
                .detail = detail
            })
        End Sub

        Private Function BuildTapAlignDataTable(rows As List(Of Dictionary(Of String, Object)),
                                                unit As String,
                                                extraHeaders As IList(Of String),
                                                locale As String,
                                                Optional targetCount As Integer = 0) As DataTable
            locale = NormalizeTapAlignExportLocale(locale)

            Dim table As New DataTable("TapAlignment")
            Dim fileHeader = ResolveTapAlignExcelHeader("File", unit, locale)
            Dim elementIdHeader = ResolveTapAlignExcelHeader("ElementId", unit, locale)
            Dim categoryHeader = ResolveTapAlignExcelHeader("Category", unit, locale)
            Dim familyHeader = ResolveTapAlignExcelHeader("Family", unit, locale)
            Dim typeHeader = ResolveTapAlignExcelHeader("Type", unit, locale)
            Dim hostIdHeader = ResolveTapAlignExcelHeader("HostId", unit, locale)
            Dim hostCategoryHeader = ResolveTapAlignExcelHeader("HostCategory", unit, locale)
            Dim hostTypeHeader = ResolveTapAlignExcelHeader("HostType", unit, locale)
            Dim domainHeader = ResolveTapAlignExcelHeader("Domain", unit, locale)
            Dim distanceHeader = ResolveTapAlignExcelHeader("DistanceFromCenter", unit, locale)
            Dim angleHeader = ResolveTapAlignExcelHeader("ModeledAngle", unit, locale)
            Dim reviewHeader = ResolveTapAlignExcelHeader("Review", unit, locale)
            Dim commentsHeader = ResolveTapAlignExcelHeader("Comments", unit, locale)

            table.Columns.Add(fileHeader, GetType(String))
            table.Columns.Add(elementIdHeader, GetType(String))
            table.Columns.Add(categoryHeader, GetType(String))
            table.Columns.Add(familyHeader, GetType(String))
            table.Columns.Add(typeHeader, GetType(String))
            table.Columns.Add(hostIdHeader, GetType(String))
            table.Columns.Add(hostCategoryHeader, GetType(String))
            table.Columns.Add(hostTypeHeader, GetType(String))
            table.Columns.Add(domainHeader, GetType(String))
            table.Columns.Add(distanceHeader, GetType(String))
            table.Columns.Add(angleHeader, GetType(String))
            table.Columns.Add(reviewHeader, GetType(String))
            table.Columns.Add(commentsHeader, GetType(String))

            If extraHeaders IsNot Nothing Then
                For Each name In extraHeaders
                    table.Columns.Add(ResolveTapAlignExtraHeader(name, locale, "branch"), GetType(String))
                    table.Columns.Add(ResolveTapAlignExtraHeader(name, locale, "host"), GetType(String))
                Next
            End If

            Dim exportRows As New List(Of Dictionary(Of String, Object))()
            If rows IsNot Nothing Then
                exportRows.AddRange(rows.Where(Function(row) row IsNot Nothing))
            End If

            If exportRows.Count = 0 Then
                exportRows.Add(New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"Status", If(targetCount > 0, "OK", "NO_TARGET")},
                    {"TargetCount", targetCount.ToString(CultureInfo.InvariantCulture)}
                })
            End If

            For Each row In exportRows
                Dim dr = table.NewRow()
                dr(fileHeader) = ReadTapAlignField(row, "File")
                dr(elementIdHeader) = ReadTapAlignField(row, "ElementId")
                dr(categoryHeader) = ReadTapAlignField(row, "Category")
                dr(familyHeader) = ReadTapAlignField(row, "Family")
                dr(typeHeader) = ReadTapAlignField(row, "Type")
                dr(hostIdHeader) = ReadTapAlignField(row, "HostId")
                dr(hostCategoryHeader) = ReadTapAlignField(row, "HostCategory")
                dr(hostTypeHeader) = ReadTapAlignField(row, "HostType")
                dr(domainHeader) = TranslateTapAlignDomain(ReadTapAlignField(row, "Domain"), locale)
                dr(distanceHeader) = ReadTapAlignField(row, "DistanceFromCenter")
                dr(angleHeader) = ReadTapAlignField(row, "ModeledAngle")
                dr(reviewHeader) = BuildTapAlignReviewText(row, locale)
                dr(commentsHeader) = BuildTapAlignCommentsText(row, locale)

                If extraHeaders IsNot Nothing Then
                    For Each name In extraHeaders
                        dr(ResolveTapAlignExtraHeader(name, locale, "branch")) = ReadTapAlignField(row, "BranchParam::" & name)
                        dr(ResolveTapAlignExtraHeader(name, locale, "host")) = ReadTapAlignField(row, "HostParam::" & name)
                    Next
                End If

                table.Rows.Add(dr)
            Next

            Return table
        End Function

        Private Function BuildTapDepthDataTable(rows As List(Of Dictionary(Of String, Object)),
                                                unit As String,
                                                extraHeaders As IList(Of String),
                                                locale As String,
                                                Optional targetCount As Integer = 0) As DataTable
            locale = NormalizeTapAlignExportLocale(locale)

            Dim table As New DataTable("TapSaddleEmbed")
            Dim fileHeader = "File"
            Dim elementIdHeader = "ElementId"
            Dim categoryHeader = "Category"
            Dim familyHeader = "Family"
            Dim typeHeader = "Type"
            Dim hostIdHeader = "Connected Host Id"
            Dim hostCategoryHeader = "Connected Host Category"
            Dim hostTypeHeader = "Connected Host Type"
            Dim domainHeader = "Domain"
            Dim projectionHeader = "Takeoff Length Projection (" & NormalizeTapAlignUnit(unit) & ")"
            Dim takeoffHeader = "Takeoff Length (" & NormalizeTapAlignUnit(unit) & ")"
            Dim standardNameHeader = "Compared Standard"
            Dim standardLengthHeader = "Compared Standard Length (" & NormalizeTapAlignUnit(unit) & ")"
            Dim actualHeader = "Actual Buried Length (" & NormalizeTapAlignUnit(unit) & ")"
            Dim differenceHeader = "Difference (" & NormalizeTapAlignUnit(unit) & ")"
            Dim reviewHeader = "Review"
            Dim commentsHeader = "Comments"

            table.Columns.Add(fileHeader, GetType(String))
            table.Columns.Add(elementIdHeader, GetType(String))
            table.Columns.Add(categoryHeader, GetType(String))
            table.Columns.Add(familyHeader, GetType(String))
            table.Columns.Add(typeHeader, GetType(String))
            table.Columns.Add(hostIdHeader, GetType(String))
            table.Columns.Add(hostCategoryHeader, GetType(String))
            table.Columns.Add(hostTypeHeader, GetType(String))
            table.Columns.Add(domainHeader, GetType(String))
            table.Columns.Add(projectionHeader, GetType(String))
            table.Columns.Add(takeoffHeader, GetType(String))
            table.Columns.Add(standardNameHeader, GetType(String))
            table.Columns.Add(standardLengthHeader, GetType(String))
            table.Columns.Add(actualHeader, GetType(String))
            table.Columns.Add(differenceHeader, GetType(String))
            table.Columns.Add(reviewHeader, GetType(String))
            table.Columns.Add(commentsHeader, GetType(String))

            If extraHeaders IsNot Nothing Then
                For Each name In extraHeaders
                    table.Columns.Add(ResolveTapAlignExtraHeader(name, locale, "branch"), GetType(String))
                    table.Columns.Add(ResolveTapAlignExtraHeader(name, locale, "host"), GetType(String))
                Next
            End If

            Dim exportRows As New List(Of Dictionary(Of String, Object))()
            If rows IsNot Nothing Then
                exportRows.AddRange(rows.Where(Function(row) row IsNot Nothing))
            End If

            If exportRows.Count = 0 Then
                exportRows.Add(New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase) From {
                    {"Status", If(targetCount > 0, "OK", "NO_TARGET")},
                    {"TargetCount", targetCount.ToString(CultureInfo.InvariantCulture)}
                })
            End If

            For Each row In exportRows
                Dim dr = table.NewRow()
                dr(fileHeader) = ReadTapAlignField(row, "File")
                dr(elementIdHeader) = ReadTapAlignField(row, "ElementId")
                dr(categoryHeader) = ReadTapAlignField(row, "Category")
                dr(familyHeader) = ReadTapAlignField(row, "Family")
                dr(typeHeader) = ReadTapAlignField(row, "Type")
                dr(hostIdHeader) = ReadTapAlignField(row, "HostId")
                dr(hostCategoryHeader) = ReadTapAlignField(row, "HostCategory")
                dr(hostTypeHeader) = ReadTapAlignField(row, "HostType")
                dr(domainHeader) = TranslateTapAlignDomain(ReadTapAlignField(row, "Domain"), locale)
                dr(projectionHeader) = ReadTapAlignField(row, "ProjectionLength")
                dr(takeoffHeader) = ReadTapAlignField(row, "TakeoffLength")
                dr(standardNameHeader) = ReadTapAlignField(row, "StandardName")
                dr(standardLengthHeader) = ReadTapAlignField(row, "StandardLength")
                dr(actualHeader) = ReadTapAlignField(row, "ActualBuriedLength")
                dr(differenceHeader) = ReadTapAlignField(row, "Difference")
                dr(reviewHeader) = BuildTapDepthReviewText(row, locale)
                dr(commentsHeader) = BuildTapDepthCommentsText(row, locale)

                If extraHeaders IsNot Nothing Then
                    For Each name In extraHeaders
                        dr(ResolveTapAlignExtraHeader(name, locale, "branch")) = ReadTapAlignField(row, "BranchParam::" & name)
                        dr(ResolveTapAlignExtraHeader(name, locale, "host")) = ReadTapAlignField(row, "HostParam::" & name)
                    Next
                End If

                table.Rows.Add(dr)
            Next

            Return table
        End Function

        Private Function ReadTapAlignCommonOptions(payload As Object) As Services.HubCommonOptionsStorageService.HubCommonOptions
            Dim stored = Services.HubCommonOptionsStorageService.Load()
            Dim result As New Services.HubCommonOptionsStorageService.HubCommonOptions() With {
                .ExtraParamsText = stored.ExtraParamsText,
                .TargetFilterText = stored.TargetFilterText,
                .ExcludeTargetFilterText = stored.ExcludeTargetFilterText,
                .ExcludeEndDummy = False,
                .IncludePointXY = stored.IncludePointXY,
                .IncludeLinearMetrics = stored.IncludeLinearMetrics
            }

            Dim commonPayload = GetProp(payload, "commonOptions")
            If commonPayload Is Nothing Then Return result

            result.ExtraParamsText = SafeTapAlignString(GetProp(commonPayload, "extraParamsText"), result.ExtraParamsText)
            result.TargetFilterText = SafeTapAlignString(GetProp(commonPayload, "targetFilterText"), result.TargetFilterText)
            result.ExcludeTargetFilterText = SafeTapAlignString(GetProp(commonPayload, "excludeTargetFilterText"), result.ExcludeTargetFilterText)
            If String.IsNullOrWhiteSpace(result.ExcludeTargetFilterText) Then
                result.ExcludeTargetFilterText = SafeTapAlignString(GetProp(commonPayload, "excludeTargetFilter"), result.ExcludeTargetFilterText)
            End If
            result.ExcludeEndDummy = False
            result.IncludePointXY = SafeTapAlignBool(GetProp(commonPayload, "includePointXY"), result.IncludePointXY)
            result.IncludeLinearMetrics = SafeTapAlignBool(GetProp(commonPayload, "includeLinearMetrics"), result.IncludeLinearMetrics)

            Return result
        End Function

        Private Shared Function ResolveTapAlignExcelHeader(key As String, unit As String, locale As String) As String
            Select Case key
                Case "File"
                    Return "File"
                Case "ElementId"
                    Return "ElementId"
                Case "Category"
                    Return "Category"
                Case "Family"
                    Return "Family"
                Case "Type"
                    Return "Type"
                Case "HostId"
                    Return "Connected Host Id"
                Case "HostCategory"
                    Return "Connected Host Category"
                Case "HostType"
                    Return "Connected Host Type"
                Case "Domain"
                    Return "Domain"
                Case "DistanceFromCenter"
                    Return "Distance From Center (" & NormalizeTapAlignUnit(unit) & ")"
                Case "ModeledAngle"
                    Return "Modeled Angle"
                Case "Review"
                    Return "Review"
                Case "Comments"
                    Return "Comments"
            End Select

            Return key
        End Function

        Private Shared Function ResolveTapAlignExtraHeader(name As String, locale As String, Optional scope As String = "host") As String
            scope = If(scope, String.Empty).Trim().ToLowerInvariant()

            If scope = "branch" Then
                Return name & " (Branch Element)"
            End If

            Return name & " (Connected Host)"
        End Function

        Private Shared Function BuildTapAlignReviewText(row As Dictionary(Of String, Object), locale As String) As String
            locale = NormalizeTapAlignExportLocale(locale)

            Dim status = ReadTapAlignField(row, "Status").Trim().ToUpperInvariant()
            Select Case status
                Case "NO_TARGET"
                    Return If(locale = "en", "No target elements found.", "검토 대상 객체가 없습니다.")
                Case "OK"
                    Dim count As Integer = 0
                    Integer.TryParse(ReadTapAlignField(row, "TargetCount"), NumberStyles.Integer, CultureInfo.InvariantCulture, count)
                    If count <= 0 Then
                        Return If(locale = "en", "No target elements found.", "검토 대상 객체가 없습니다.")
                    End If

                    If locale = "en" Then
                        Return $"All {count} connections are aligned with the centerline."
                    End If

                    Return $"전체 {count}개 연결이 중심축에 정렬되어 있습니다."
                Case "NO_HOST"
                    Dim typeName = ReadTapAlignField(row, "Type")
                    If String.IsNullOrWhiteSpace(typeName) Then typeName = If(locale = "en", "Unknown", "알 수 없음")

                    If locale = "en" Then
                        Return $"Review Target: {typeName} - Connected host centerline could not be resolved."
                    End If

                    Return $"검토 대상: {typeName} - 연결 중심축을 확인할 수 없습니다."
                Case Else
                    Dim typeName = ReadTapAlignField(row, "HostType")
                    If String.IsNullOrWhiteSpace(typeName) Then typeName = ReadTapAlignField(row, "Type")
                    If String.IsNullOrWhiteSpace(typeName) Then typeName = If(locale = "en", "Unknown", "알 수 없음")

                    Dim prefix = ResolveTapAlignReviewTypeLabel(row, locale)
                    If locale = "en" Then
                        Return $"{prefix}: {typeName} - The connection is misaligned from the centerline."
                    End If

                    Return $"{prefix}: {typeName} - 연결이 중심축에서 벗어났습니다."
            End Select
        End Function

        Private Shared Function BuildTapDepthReviewText(row As Dictionary(Of String, Object), locale As String) As String
            locale = NormalizeTapAlignExportLocale(locale)

            Dim status = ReadTapAlignField(row, "Status").Trim().ToUpperInvariant()
            Select Case status
                Case "NO_TARGET"
                    Return If(locale = "en", "No target elements found.", "검토 대상 객체가 없습니다.")
                Case "OK"
                    Dim count As Integer = 0
                    Integer.TryParse(ReadTapAlignField(row, "TargetCount"), NumberStyles.Integer, CultureInfo.InvariantCulture, count)
                    If count <= 0 Then
                        Return If(locale = "en", "No target elements found.", "검토 대상 객체가 없습니다.")
                    End If

                    If locale = "en" Then
                        Return $"All {count} tap/saddle elements match Takeoff Length Projection or Takeoff Length."
                    End If

                    Return $"전체 {count}개 Tap/Saddle 객체의 묻힘 깊이가 Takeoff Length Projection 또는 Takeoff Length 기준 안에 있습니다."
                Case "HOST_DISCONNECTED"
                    Dim typeName = ReadTapAlignField(row, "Type")
                    If String.IsNullOrWhiteSpace(typeName) Then typeName = If(locale = "en", "Unknown", "알 수 없음")
                    Dim hostType = ReadTapAlignField(row, "HostType")
                    If String.IsNullOrWhiteSpace(hostType) Then hostType = If(locale = "en", "host", "호스트")
                    Dim standardName = ReadTapAlignField(row, "StandardName")
                    Dim standardText = ReadTapAlignField(row, "StandardLength")
                    Dim actualText = ReadTapAlignField(row, "ActualBuriedLength")

                    If locale = "en" Then
                        Return $"Tap/Saddle: {typeName} - The host connector is disconnected. Reconnect it to the host. (host: {hostType}, closest: {standardName}, standard: {standardText}, actual: {actualText})"
                    End If

                    Return $"Tap/Saddle: {typeName} - 호스트 배관/덕트와 커넥터 연결이 끊어져 있습니다. 호스트와 다시 연결해 주세요. (호스트:{hostType}, 가장 가까운 기준:{standardName}, 기준값:{standardText}, 실제:{actualText})"
                Case Else
                    Dim typeName = ReadTapAlignField(row, "Type")
                    If String.IsNullOrWhiteSpace(typeName) Then typeName = If(locale = "en", "Unknown", "알 수 없음")
                    Dim standardName = ReadTapAlignField(row, "StandardName")
                    If String.IsNullOrWhiteSpace(standardName) Then standardName = "Takeoff Length Projection / Takeoff Length"
                    Dim standardText = ReadTapAlignField(row, "StandardLength")
                    Dim actualText = ReadTapAlignField(row, "ActualBuriedLength")
                    Dim differenceText = ReadTapAlignField(row, "Difference")

                    If locale = "en" Then
                        Return $"Tap/Saddle: {typeName} - Buried depth is outside Takeoff Length Projection and Takeoff Length. (closest: {standardName}, standard: {standardText}, actual: {actualText}, difference: {differenceText})"
                    End If

                    Return $"Tap/Saddle: {typeName} - Takeoff Length Projection / Takeoff Length 기준 묻힘 깊이가 벗어났습니다. (가장 가까운 기준:{standardName}, 기준값:{standardText}, 실제:{actualText}, 차이:{differenceText})"
            End Select
        End Function

        Private Shared Function BuildTapAlignCommentsText(row As Dictionary(Of String, Object), locale As String) As String
            locale = NormalizeTapAlignExportLocale(locale)

            Dim status = ReadTapAlignField(row, "Status").Trim().ToUpperInvariant()
            If status = "OK" OrElse status = "NO_TARGET" Then Return String.Empty

            If status = "NO_HOST" Then
                If locale = "en" Then
                    Return "Please verify the connection state and the connected host line."
                End If

                Return "연결 상태와 연결된 호스트 라인을 확인해주세요."
            End If

            If locale = "en" Then
                Return "Please reconnect the elements so that they are aligned with the central axis."
            End If

            Return "요소를 다시 연결하여 중심축에 정렬되도록 해주세요."
        End Function

        Private Shared Function BuildTapDepthCommentsText(row As Dictionary(Of String, Object), locale As String) As String
            locale = NormalizeTapAlignExportLocale(locale)

            Dim status = ReadTapAlignField(row, "Status").Trim().ToUpperInvariant()
            If status = "OK" OrElse status = "NO_TARGET" Then Return String.Empty

            If status = "HOST_DISCONNECTED" Then
                If locale = "en" Then
                    Return "Reconnect the tap/saddle to the host pipe or duct first, then run the embed-depth review again."
                End If

                Return "Tap/Saddle을 호스트 배관/덕트에 다시 연결한 뒤 묻힘 검토를 다시 실행해 주세요."
            End If

            If locale = "en" Then
                Return "Please adjust the tap/saddle insertion depth against the host centerline."
            End If

            Return "호스트 배관/덕트 기준 묻힘 깊이를 확인하고 Tap/Saddle 배치를 조정해 주세요."
        End Function

        Private Shared Function ResolveTapAlignReviewTypeLabel(row As Dictionary(Of String, Object), locale As String) As String
            locale = NormalizeTapAlignExportLocale(locale)

            Dim domain = ReadTapAlignField(row, "Domain")
            Dim hostCategory = ReadTapAlignField(row, "HostCategory")
            Dim isDuct As Boolean =
                String.Equals(domain, "Duct", StringComparison.OrdinalIgnoreCase) OrElse
                hostCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                hostCategory.IndexOf("덕트", StringComparison.OrdinalIgnoreCase) >= 0

            If locale = "en" Then
                Return If(isDuct, "Duct Type", "Pipe Type")
            End If

            Return If(isDuct, "덕트 타입", "배관 타입")
        End Function

        Private Shared Function TranslateTapAlignMessage(message As String, locale As String) As String
            Dim text = If(message, String.Empty).Trim()
            If text = String.Empty Then Return String.Empty

            locale = NormalizeTapAlignExportLocale(locale)
            If locale = "en" Then
                If String.Equals(text, "중심축에서 벗어났습니다.", StringComparison.OrdinalIgnoreCase) Then
                    Return "Offset from center axis."
                End If
                If String.Equals(text, "연결 중심축을 확인할 수 없습니다.", StringComparison.OrdinalIgnoreCase) Then
                    Return "Connected host centerline could not be resolved."
                End If
                Return text
            End If

            If String.Equals(text, "Offset from center axis.", StringComparison.OrdinalIgnoreCase) Then
                Return "중심축에서 벗어났습니다."
            End If
            If String.Equals(text, "Connected host centerline could not be resolved.", StringComparison.OrdinalIgnoreCase) Then
                Return "연결 중심축을 확인할 수 없습니다."
            End If
            Return text
        End Function

        Private Shared Function TranslateTapAlignDomain(domain As String, locale As String) As String
            Dim text = If(domain, String.Empty).Trim()
            If text = String.Empty Then Return String.Empty

            locale = NormalizeTapAlignExportLocale(locale)
            If locale = "en" Then
                If String.Equals(text, "배관", StringComparison.OrdinalIgnoreCase) Then Return "Pipe"
                If String.Equals(text, "덕트", StringComparison.OrdinalIgnoreCase) Then Return "Duct"
                Return text
            End If

            If String.Equals(text, "Pipe", StringComparison.OrdinalIgnoreCase) Then Return "배관"
            If String.Equals(text, "Duct", StringComparison.OrdinalIgnoreCase) Then Return "덕트"
            Return text
        End Function

        Private Shared Function SafeTapAlignString(raw As Object, fallback As String) As String
            Try
                Dim text = TryCast(raw, String)
                If text IsNot Nothing Then Return text
                If raw IsNot Nothing Then Return raw.ToString()
            Catch
            End Try
            Return If(fallback, String.Empty)
        End Function

        Private Shared Function SafeTapAlignBool(raw As Object, fallback As Boolean) As Boolean
            Try
                If raw Is Nothing Then Return fallback
                Return Convert.ToBoolean(raw, CultureInfo.InvariantCulture)
            Catch
                Return fallback
            End Try
        End Function

        Private Shared Function NormalizeTapAlignUnit(unit As String) As String
            Dim normalized = If(unit, String.Empty).Trim().ToLowerInvariant()
            If normalized = "inch" OrElse normalized = "in" OrElse normalized = "inches" Then Return "inch"
            Return "mm"
        End Function

        Private Shared Function NormalizeTapAlignDomain(domain As String) As String
            Dim normalized = If(domain, String.Empty).Trim().ToLowerInvariant()
            If normalized = "pipe" OrElse normalized = "piping" Then Return "pipe"
            If normalized = "duct" OrElse normalized = "hvac" Then Return "duct"
            Return "all"
        End Function

        Private Shared Function NormalizeTapAlignExportLocale(locale As String) As String
            Dim normalized = If(locale, String.Empty).Trim().ToLowerInvariant()
            If normalized = "en" OrElse normalized = "eng" OrElse normalized = "english" Then Return "en"
            Return "ko"
        End Function

        Private Shared Function ResolveTapAlignDomainLabel(domain As String) As String
            Dim normalized = NormalizeTapAlignDomain(domain)
            If normalized = "pipe" Then Return "배관"
            If normalized = "duct" Then Return "덕트"
            Return "배관 + 덕트"
        End Function

        Private Shared Function ReadTapAlignField(row As Dictionary(Of String, Object), key As String) As String
            If row Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return String.Empty

            Try
                If row.ContainsKey(key) AndAlso row(key) IsNot Nothing Then
                    Return row(key).ToString()
                End If
            Catch
            End Try

            Return String.Empty
        End Function

    End Class

End Namespace
