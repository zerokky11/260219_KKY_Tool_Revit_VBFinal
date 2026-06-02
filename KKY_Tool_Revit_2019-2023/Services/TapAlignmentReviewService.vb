Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Namespace Services

    Public NotInheritable Class TapAlignmentReviewService

        Private Const MisalignedMessage As String = "중심축에서 벗어났습니다."
        Private Const UnresolvedHostMessage As String = "연결 중심축을 확인할 수 없습니다."
        Private Const BranchAngleThresholdDeg As Double = 5.0R
        Private Const MidCurveEndpointToleranceFt As Double = 1.0R / 304.8R
        Private Const NearbyHostSearchMarginFt As Double = 3.0R
        Private Const NearbyHostBaseToleranceFt As Double = 150.0R / 304.8R
        Private Const NearbyHostMaxToleranceFt As Double = 12.0R
        Private Const ConnectedHostTraversalDepth As Integer = 2

        Private Sub New()
        End Sub

        Private Class ReviewHit
            Public Property Host As Element
            Public Property DistanceFt As Double
            Public Property AngleDeg As Double
            Public Property DisplayAngleDeg As Double
            Public Property HostPointDistanceFt As Double
            Public Property Status As String
            Public Property Message As String
        End Class

        Private Class ProjectionDepthHit
            Public Property Host As Element
            Public Property HasProjectionLength As Boolean
            Public Property ProjectionLengthFt As Double
            Public Property HasTakeoffLength As Boolean
            Public Property TakeoffLengthFt As Double
            Public Property StandardName As String
            Public Property StandardFt As Double
            Public Property ActualBuriedFt As Double
            Public Property DifferenceFt As Double
            Public Property ProjectionDifferenceFt As Double
            Public Property TakeoffDifferenceFt As Double
            Public Property RadialDistanceFt As Double
            Public Property HostHalfSizeFt As Double
            Public Property HostPointDistanceFt As Double
            Public Property HostConnectionMissing As Boolean
            Public Property Status As String
            Public Property Message As String
        End Class

        Private Class ProjectionDepthStandard
            Public Property Name As String
            Public Property ValueFt As Double
        End Class

        Private Class TargetFilter
            Public Property Evaluator As Func(Of Element, Boolean)
        End Class

        Private Class FilterToken
            Public Property Kind As String
            Public Property Text As String
        End Class

        Private Class FilterParser
            Private ReadOnly _tokens As List(Of FilterToken)
            Private _position As Integer

            Public Sub New(raw As String)
                _tokens = Tokenize(raw)
            End Sub

            Public Function Parse() As Func(Of Element, Boolean)
                If _tokens.Count = 0 Then Return Nothing

                Dim first = ParseExpression()
                If first Is Nothing Then Return Nothing

                Dim expressions As New List(Of Func(Of Element, Boolean)) From {first}
                While PeekIs("comma")
                    [Next]()
                    Dim nextExpression = ParseExpression()
                    If nextExpression Is Nothing Then Exit While
                    expressions.Add(nextExpression)
                End While

                If expressions.Count = 1 Then Return first

                Return Function(el As Element)
                           For Each expression In expressions
                               If expression IsNot Nothing AndAlso Not expression(el) Then Return False
                           Next
                           Return True
                       End Function
            End Function

            Private Function ParseExpression() As Func(Of Element, Boolean)
                If AtEnd() Then Return Nothing

                Dim token = Peek()
                If token Is Nothing OrElse token.Kind <> "ident" Then Return Nothing

                If PeekIs("lparen", 1) Then
                    Return ParseFunction()
                End If

                Return ParseComparison()
            End Function

            Private Function ParseFunction() As Func(Of Element, Boolean)
                Dim nameToken = Expect("ident")
                If nameToken Is Nothing Then Return Nothing

                Dim functionName = nameToken.Text.ToLowerInvariant()
                Expect("lparen")

                Dim arguments As New List(Of Func(Of Element, Boolean))()
                While Not AtEnd()
                    If PeekIs("rparen") Then Exit While

                    Dim argument = ParseExpression()
                    If argument Is Nothing Then Exit While
                    arguments.Add(argument)

                    If PeekIs("comma") Then
                        [Next]()
                    ElseIf PeekIs("rparen") Then
                        Exit While
                    ElseIf PeekIs("ident") OrElse PeekIs("lparen") Then
                        Continue While
                    Else
                        Exit While
                    End If
                End While

                Expect("rparen")

                Select Case functionName
                    Case "and"
                        Return Function(el As Element)
                                   For Each argument In arguments
                                       If argument IsNot Nothing AndAlso Not argument(el) Then Return False
                                   Next
                                   Return True
                               End Function
                    Case "or"
                        Return Function(el As Element)
                                   For Each argument In arguments
                                       If argument IsNot Nothing AndAlso argument(el) Then Return True
                                   Next
                                   Return False
                               End Function
                    Case "not"
                        Dim inner = If(arguments.Count > 0, arguments(0), Nothing)
                        Return Function(el As Element)
                                   If inner Is Nothing Then Return True
                                   Return Not inner(el)
                               End Function
                End Select

                Return Nothing
            End Function

            Private Function ParseComparison() As Func(Of Element, Boolean)
                Dim left = Expect("ident")
                If left Is Nothing Then Return Nothing
                Expect("eq")

                Dim right = ExpectValue()
                If right Is Nothing Then Return Nothing

                Dim paramName = left.Text
                Dim expected = right.Text

                Return Function(el As Element)
                           Dim candidates = ResolveParamTexts(el, paramName)
                           If candidates Is Nothing OrElse candidates.Count = 0 Then Return False

                           For Each actual In candidates
                               If String.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase) Then
                                   Return True
                               End If
                           Next

                           Return False
                       End Function
            End Function

            Private Function Expect(kind As String) As FilterToken
                If PeekIs(kind) Then Return [Next]()
                Return Nothing
            End Function

            Private Function ExpectValue() As FilterToken
                If PeekIs("string") OrElse PeekIs("ident") Then Return [Next]()
                Return Nothing
            End Function

            Private Function Peek() As FilterToken
                If _position < 0 OrElse _position >= _tokens.Count Then Return Nothing
                Return _tokens(_position)
            End Function

            Private Function PeekIs(kind As String, Optional offset As Integer = 0) As Boolean
                Dim index = _position + offset
                If index < 0 OrElse index >= _tokens.Count Then Return False
                Return String.Equals(_tokens(index).Kind, kind, StringComparison.OrdinalIgnoreCase)
            End Function

            Private Function [Next]() As FilterToken
                Dim token = Peek()
                _position += 1
                Return token
            End Function

            Private Function AtEnd() As Boolean
                Return _position >= _tokens.Count
            End Function

            Private Shared Function Tokenize(raw As String) As List(Of FilterToken)
                Dim tokens As New List(Of FilterToken)()
                If String.IsNullOrWhiteSpace(raw) Then Return tokens

                Dim index As Integer = 0
                While index < raw.Length
                    Dim ch = raw(index)

                    If Char.IsWhiteSpace(ch) Then
                        index += 1
                        Continue While
                    End If

                    If ch = "("c Then
                        tokens.Add(New FilterToken With {.Kind = "lparen", .Text = "("})
                        index += 1
                        Continue While
                    End If
                    If ch = ")"c Then
                        tokens.Add(New FilterToken With {.Kind = "rparen", .Text = ")"})
                        index += 1
                        Continue While
                    End If
                    If ch = ","c OrElse ch = ";"c Then
                        tokens.Add(New FilterToken With {.Kind = "comma", .Text = ","})
                        index += 1
                        Continue While
                    End If
                    If ch = "="c Then
                        tokens.Add(New FilterToken With {.Kind = "eq", .Text = "="})
                        index += 1
                        Continue While
                    End If

                    If ch = "'"c OrElse ch = """"c Then
                        Dim quote = ch
                        index += 1
                        Dim start = index
                        While index < raw.Length AndAlso raw(index) <> quote
                            index += 1
                        End While

                        Dim content = raw.Substring(start, index - start)
                        tokens.Add(New FilterToken With {.Kind = "string", .Text = content})

                        If index < raw.Length AndAlso raw(index) = quote Then index += 1
                        Continue While
                    End If

                    Dim wordStart = index
                    While index < raw.Length AndAlso
                          Not Char.IsWhiteSpace(raw(index)) AndAlso
                          raw(index) <> "("c AndAlso
                          raw(index) <> ")"c AndAlso
                          raw(index) <> ","c AndAlso
                          raw(index) <> ";"c AndAlso
                          raw(index) <> "="c
                        index += 1
                    End While

                    Dim word = raw.Substring(wordStart, index - wordStart)
                    If word <> String.Empty Then
                        tokens.Add(New FilterToken With {.Kind = "ident", .Text = word})
                    End If
                End While

                Return tokens
            End Function
        End Class

        Public Shared Function Run(app As UIApplication,
                                   tol As Double,
                                   unit As String,
                                   domain As String,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Return Run(app, tol, unit, domain, CType(Nothing, IEnumerable(Of String)), Nothing, False, progress)
        End Function

        Public Shared Function Run(app As UIApplication,
                                   tol As Double,
                                   unit As String,
                                   domain As String,
                                   extraParams As IEnumerable(Of String),
                                   targetFilter As String,
                                   excludeEndDummy As Boolean,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If app Is Nothing OrElse app.ActiveUIDocument Is Nothing OrElse app.ActiveUIDocument.Document Is Nothing Then
                Return rows
            End If

            Dim doc = app.ActiveUIDocument.Document
            Dim tolFt = ConnectorDiagnosticsService.ToTolFt(tol, unit)
            Return RunOnDocument(doc,
                                 tolFt,
                                 NormalizeUnit(unit),
                                 NormalizeDomain(domain),
                                 NormalizeExtraParams(extraParams),
                                 ParseTargetFilter(targetFilter),
                                 Nothing,
                                 excludeEndDummy,
                                 progress)
        End Function

        Public Shared Function Run(app As UIApplication,
                                   tol As Double,
                                   unit As String,
                                   domain As String,
                                   extraParams As IEnumerable(Of String),
                                   targetFilter As String,
                                   excludeTargetFilter As String,
                                   Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If app Is Nothing OrElse app.ActiveUIDocument Is Nothing OrElse app.ActiveUIDocument.Document Is Nothing Then
                Return rows
            End If

            Dim doc = app.ActiveUIDocument.Document
            Dim tolFt = ConnectorDiagnosticsService.ToTolFt(tol, unit)
            Return RunOnDocument(doc,
                                 tolFt,
                                 NormalizeUnit(unit),
                                 NormalizeDomain(domain),
                                 NormalizeExtraParams(extraParams),
                                 ParseTargetFilter(targetFilter),
                                 ParseTargetFilter(excludeTargetFilter),
                                 False,
                                 progress)
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             tol As Double,
                                             unit As String,
                                             domain As String,
                                             extraParams As IEnumerable(Of String),
                                             targetFilter As String,
                                             excludeEndDummy As Boolean,
                                             Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If doc Is Nothing Then Return rows

            Dim tolFt = ConnectorDiagnosticsService.ToTolFt(tol, unit)
            Return RunOnDocument(doc,
                                 tolFt,
                                 NormalizeUnit(unit),
                                 NormalizeDomain(domain),
                                 NormalizeExtraParams(extraParams),
                                 ParseTargetFilter(targetFilter),
                                 Nothing,
                                 excludeEndDummy,
                                 progress)
        End Function

        Public Shared Function RunOnDocument(doc As Document,
                                             tol As Double,
                                             unit As String,
                                             domain As String,
                                             extraParams As IEnumerable(Of String),
                                             targetFilter As String,
                                             excludeTargetFilter As String,
                                             Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If doc Is Nothing Then Return rows

            Dim tolFt = ConnectorDiagnosticsService.ToTolFt(tol, unit)
            Return RunOnDocument(doc,
                                 tolFt,
                                 NormalizeUnit(unit),
                                 NormalizeDomain(domain),
                                 NormalizeExtraParams(extraParams),
                                 ParseTargetFilter(targetFilter),
                                 ParseTargetFilter(excludeTargetFilter),
                                 False,
                                 progress)
        End Function

        Public Shared Function CountTargets(app As UIApplication,
                                            domain As String,
                                            targetFilter As String,
                                            excludeEndDummy As Boolean) As Integer
            If app Is Nothing OrElse app.ActiveUIDocument Is Nothing OrElse app.ActiveUIDocument.Document Is Nothing Then
                Return 0
            End If

            Return CountTargetsOnDocument(app.ActiveUIDocument.Document, domain, targetFilter, excludeEndDummy)
        End Function

        Public Shared Function CountTargets(app As UIApplication,
                                            domain As String,
                                            targetFilter As String,
                                            excludeTargetFilter As String) As Integer
            If app Is Nothing OrElse app.ActiveUIDocument Is Nothing OrElse app.ActiveUIDocument.Document Is Nothing Then
                Return 0
            End If

            Return CountTargetsOnDocument(app.ActiveUIDocument.Document, domain, targetFilter, excludeTargetFilter)
        End Function

        Public Shared Function CombineTargetFilterText(ParamArray filters() As String) As String
            If filters Is Nothing OrElse filters.Length = 0 Then Return String.Empty

            Dim parts = filters.
                Where(Function(raw) Not String.IsNullOrWhiteSpace(raw)).
                Select(Function(raw) raw.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            If parts.Count = 0 Then Return String.Empty
            Return String.Join("; ", parts)
        End Function

        Public Shared Function CountTargetsOnDocument(doc As Document,
                                                      domain As String,
                                                      targetFilter As String,
                                                      excludeEndDummy As Boolean) As Integer
            If doc Is Nothing Then Return 0

            Return CountTargetsOnDocument(doc,
                                          NormalizeDomain(domain),
                                          ParseTargetFilter(targetFilter),
                                          Nothing,
                                          excludeEndDummy)
        End Function

        Public Shared Function CountTargetsOnDocument(doc As Document,
                                                      domain As String,
                                                      targetFilter As String,
                                                      excludeTargetFilter As String) As Integer
            If doc Is Nothing Then Return 0

            Return CountTargetsOnDocument(doc,
                                          NormalizeDomain(domain),
                                          ParseTargetFilter(targetFilter),
                                          ParseTargetFilter(excludeTargetFilter),
                                          False)
        End Function

        Public Shared Function RunProjectionDepthOnDocument(doc As Document,
                                                            tol As Double,
                                                            unit As String,
                                                            domain As String,
                                                            extraParams As IEnumerable(Of String),
                                                            targetFilter As String,
                                                            excludeTargetFilter As String,
                                                            Optional progress As Action(Of Double, String) = Nothing) As List(Of Dictionary(Of String, Object))
            If doc Is Nothing Then Return New List(Of Dictionary(Of String, Object))()

            Dim tolFt = ConnectorDiagnosticsService.ToTolFt(tol, unit)
            Return RunProjectionDepthOnDocument(doc,
                                                tolFt,
                                                NormalizeUnit(unit),
                                                NormalizeDomain(domain),
                                                NormalizeExtraParams(extraParams),
                                                ParseTargetFilter(targetFilter),
                                                ParseTargetFilter(excludeTargetFilter),
                                                progress)
        End Function

        Private Shared Function CountTargetsOnDocument(doc As Document,
                                                       domain As String,
                                                       includeFilter As TargetFilter,
                                                       excludeFilter As TargetFilter,
                                                       excludeEndDummy As Boolean) As Integer
            If doc Is Nothing Then Return 0

            Return CollectCandidates(doc, domain).
                Where(Function(el) IsElementAllowed(el, includeFilter, excludeFilter, excludeEndDummy)).
                GroupBy(Function(el) el.Id.IntValue()).
                Count()
        End Function

        Private Shared Function RunOnDocument(doc As Document,
                                              tolFt As Double,
                                              unit As String,
                                              domain As String,
                                              extraParams As IList(Of String),
                                              includeFilter As TargetFilter,
                                              excludeFilter As TargetFilter,
                                              excludeEndDummy As Boolean,
                                              progress As Action(Of Double, String)) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If doc Is Nothing Then Return rows

            Dim candidates = CollectCandidates(doc, domain)
            candidates = candidates.
                Where(Function(el) IsElementAllowed(el, includeFilter, excludeFilter, excludeEndDummy)).
                GroupBy(Function(el) el.Id.IntValue()).
                Select(Function(g) g.First()).
                ToList()

            If progress IsNot Nothing Then
                progress(0.05R, "탭/분기 축 틀어짐 후보를 수집하는 중...")
            End If

            If candidates.Count = 0 Then
                If progress IsNot Nothing Then progress(1.0R, "완료")
                Return rows
            End If

            Dim total = Math.Max(1, candidates.Count)
            For index As Integer = 0 To candidates.Count - 1
                Dim candidate = candidates(index)
                Dim hit = FindBestMisalignment(candidate, domain, tolFt)
                If hit IsNot Nothing Then
                    rows.Add(BuildRow(doc, candidate, hit, unit, extraParams))
                End If

                If progress IsNot Nothing Then
                    Dim pct = 0.05R + (0.95R * CDbl(index + 1) / CDbl(total))
                    progress(pct, String.Format(CultureInfo.InvariantCulture, "탭/분기 축 틀어짐 검토 중... ({0}/{1})", index + 1, total))
                End If
            Next

            Return rows.
                OrderByDescending(Function(row) ParseNumber(ReadField(row, "DistanceFromCenter"))).
                ThenBy(Function(row) ParseInteger(ReadField(row, "ElementId"))).
                ToList()
        End Function

        Private Shared Function RunProjectionDepthOnDocument(doc As Document,
                                                             tolFt As Double,
                                                             unit As String,
                                                             domain As String,
                                                             extraParams As IList(Of String),
                                                             includeFilter As TargetFilter,
                                                             excludeFilter As TargetFilter,
                                                             progress As Action(Of Double, String)) As List(Of Dictionary(Of String, Object))
            Dim rows As New List(Of Dictionary(Of String, Object))()
            If doc Is Nothing Then Return rows

            Dim candidates = CollectCandidates(doc, domain)
            candidates = candidates.
                Where(Function(el) IsElementAllowed(el, includeFilter, excludeFilter, False)).
                GroupBy(Function(el) el.Id.IntValue()).
                Select(Function(g) g.First()).
                ToList()

            If progress IsNot Nothing Then
                progress(0.05R, "Tap, Saddle 모델링 검토(묻힘) 후보를 수집하는 중...")
            End If

            If candidates.Count = 0 Then
                If progress IsNot Nothing Then progress(1.0R, "완료")
                Return rows
            End If

            Dim total = Math.Max(1, candidates.Count)
            For index As Integer = 0 To candidates.Count - 1
                Dim candidate = candidates(index)
                Dim hit = FindBestProjectionDepthIssue(candidate, domain, tolFt)
                If hit IsNot Nothing Then
                    rows.Add(BuildProjectionDepthRow(doc, candidate, hit, unit, extraParams))
                End If

                If progress IsNot Nothing Then
                    Dim pct = 0.05R + (0.95R * CDbl(index + 1) / CDbl(total))
                    progress(pct, String.Format(CultureInfo.InvariantCulture, "Tap, Saddle 모델링 검토(묻힘) 중... ({0}/{1})", index + 1, total))
                End If
            Next

            Return rows.
                OrderByDescending(Function(row) Math.Abs(ParseNumber(ReadField(row, "Difference")))).
                ThenBy(Function(row) ParseInteger(ReadField(row, "ElementId"))).
                ToList()
        End Function

        Private Shared Function FindBestProjectionDepthIssue(candidate As Element,
                                                             domain As String,
                                                             tolFt As Double) As ProjectionDepthHit
            Dim connectors = GetConnectors(candidate)
            If connectors.Count = 0 Then Return Nothing

            Dim standards = ReadProjectionDepthStandards(candidate)
            If standards.Count = 0 Then
                Return Nothing
            End If

            Dim compared As Boolean = False
            Dim best = FindBestProjectionDepthHit(connectors,
                                                  GetCandidateHostCurves(candidate, domain, includeNearbyFallback:=False),
                                                  compared)

            If best Is Nothing AndAlso Not compared Then
                best = FindBestProjectionDepthHit(connectors,
                                                  GetCandidateHostCurves(candidate, domain, includeNearbyFallback:=True),
                                                  compared)
            End If

            If best Is Nothing Then Return Nothing
            ApplyProjectionDepthStandards(best, standards)
            If best.StandardFt <= 0.0R Then Return Nothing

            If Not HasConnectorPathToHost(connectors, TryCast(best.Host, MEPCurve), domain) Then
                best.HostConnectionMissing = True
                best.Status = "HOST_DISCONNECTED"
                best.Message = "호스트 배관/덕트와 커넥터 연결이 끊어져 있습니다. 호스트와 다시 연결해 주세요."
                Return best
            End If

            If Math.Abs(best.DifferenceFt) <= tolFt Then Return Nothing

            best.Status = If(best.DifferenceFt > 0.0R, "EMBED_TOO_DEEP", "EMBED_TOO_SHALLOW")
            best.Message = If(best.DifferenceFt > 0.0R,
                              $"{best.StandardName} 기준보다 깊게 묻혀 있습니다.",
                              $"{best.StandardName} 기준보다 얕게 묻혀 있습니다.")
            Return best
        End Function

        Private Shared Function FindBestProjectionDepthHit(connectors As IList(Of Connector),
                                                           candidateHosts As IList(Of MEPCurve),
                                                           ByRef compared As Boolean) As ProjectionDepthHit
            Dim best As ProjectionDepthHit = Nothing
            If connectors Is Nothing OrElse candidateHosts Is Nothing OrElse candidateHosts.Count = 0 Then Return best

            For Each host In candidateHosts
                Dim hit As ProjectionDepthHit = Nothing
                If Not TryEvaluateProjectionDepth(connectors, host, hit) Then Continue For

                compared = True
                If best Is Nothing OrElse hit.HostPointDistanceFt < best.HostPointDistanceFt Then
                    best = hit
                End If
            Next

            Return best
        End Function

        Private Shared Function FindBestMisalignment(candidate As Element,
                                                     domain As String,
                                                     tolFt As Double) As ReviewHit
            Dim connectors = GetConnectors(candidate)
            Dim compared As Boolean = False

            Dim best = FindBestHostHit(connectors,
                                       GetCandidateHostCurves(candidate, domain, includeNearbyFallback:=False),
                                       compared)

            If best Is Nothing AndAlso Not compared Then
                best = FindBestHostHit(connectors,
                                       GetCandidateHostCurves(candidate, domain, includeNearbyFallback:=True),
                                       compared)
            End If

            If best Is Nothing AndAlso connectors.Count > 0 AndAlso Not compared Then
                Return New ReviewHit With {
                    .Host = Nothing,
                    .DistanceFt = Double.NaN,
                    .AngleDeg = Double.NaN,
                    .DisplayAngleDeg = Double.NaN,
                    .Status = "NO_HOST",
                    .Message = UnresolvedHostMessage
                }
            End If

            If best Is Nothing Then Return Nothing
            If best.DistanceFt <= tolFt Then Return Nothing

            Return best
        End Function

        Private Shared Function FindBestHostHit(connectors As IList(Of Connector),
                                                candidateHosts As IList(Of MEPCurve),
                                                ByRef compared As Boolean) As ReviewHit
            Dim best As ReviewHit = Nothing
            If connectors Is Nothing OrElse candidateHosts Is Nothing OrElse candidateHosts.Count = 0 Then Return best

            For Each host In candidateHosts
                Dim hit As ReviewHit = Nothing
                If Not TryEvaluateHostBranch(connectors, host, hit) Then Continue For

                compared = True
                If best Is Nothing OrElse hit.HostPointDistanceFt < best.HostPointDistanceFt Then
                    best = hit
                End If
            Next

            Return best
        End Function

        Private Shared Function GetCandidateHostCurves(candidate As Element,
                                                       domain As String,
                                                       includeNearbyFallback As Boolean) As List(Of MEPCurve)
            Dim results As New List(Of MEPCurve)()
            Dim seen As New HashSet(Of Integer)()

            Dim connectors = GetConnectors(candidate)
            Dim midCurveHosts = GetMidCurveConnectedHostCurves(connectors, domain)
            For Each curve In midCurveHosts
                AddUniqueHostCurve(results, seen, curve, domain)
            Next

            Dim hostedCurve = TryGetHostedCurve(candidate)
            If hostedCurve IsNot Nothing Then
                AddUniqueHostCurve(results, seen, hostedCurve, domain)
            End If

            Dim branchAxis As XYZ = Nothing
            If Not TryResolveBranchAxisFromConnectorLayout(connectors, branchAxis) Then Return results

            For Each connector In connectors
                For Each curve In GetConnectedHostCurves(connector, domain)
                    Dim centerLine As Line = Nothing
                    If Not TryGetCenterLine(curve, centerLine) Then Continue For

                    Dim hostDirection = NormalizeVector(centerLine.Direction)
                    If hostDirection Is Nothing Then Continue For

                    Dim angleToBranch = AngleBetweenAxesDeg(branchAxis, hostDirection)
                    If Double.IsNaN(angleToBranch) OrElse Double.IsInfinity(angleToBranch) Then Continue For
                    If angleToBranch <= BranchAngleThresholdDeg Then Continue For

                    AddUniqueHostCurve(results, seen, curve, domain)
                Next
            Next

            If includeNearbyFallback Then
                AddNearbyMidCurveHostCurves(candidate, connectors, branchAxis, domain, results, seen)
            End If

            Return results
        End Function

        Private Shared Sub AddNearbyMidCurveHostCurves(candidate As Element,
                                                       connectors As IList(Of Connector),
                                                       branchAxis As XYZ,
                                                       domain As String,
                                                       results As IList(Of MEPCurve),
                                                       seen As HashSet(Of Integer))
            If candidate Is Nothing OrElse connectors Is Nothing OrElse connectors.Count = 0 OrElse branchAxis Is Nothing Then Return
            If results Is Nothing OrElse seen Is Nothing Then Return

            Dim hostCurves = CollectNearbyHostCurves(candidate, domain)
            If hostCurves.Count = 0 Then Return

            For Each connector In connectors
                Dim origin As XYZ = Nothing
                If Not TryGetConnectorOrigin(connector, origin) Then Continue For

                For Each curve In hostCurves
                    If curve Is Nothing OrElse curve.Id Is Nothing Then Continue For
                    If curve.Id.IntValue() = candidate.Id.IntValue() Then Continue For
                    If seen.Contains(curve.Id.IntValue()) Then Continue For

                    Dim centerLine As Line = Nothing
                    If Not TryGetCenterLine(curve, centerLine) Then Continue For

                    Dim hostDirection = NormalizeVector(centerLine.Direction)
                    If hostDirection Is Nothing Then Continue For

                    Dim angleToBranch = AngleBetweenAxesDeg(branchAxis, hostDirection)
                    If Double.IsNaN(angleToBranch) OrElse Double.IsInfinity(angleToBranch) Then Continue For
                    If angleToBranch <= BranchAngleThresholdDeg Then Continue For

                    If Not IsPointInsideCurveSpan(origin, centerLine) Then Continue For

                    Dim distanceFt = DistanceFromPointToInfiniteLine(origin, centerLine.GetEndPoint(0), hostDirection)
                    If Double.IsNaN(distanceFt) OrElse Double.IsInfinity(distanceFt) Then Continue For
                    If distanceFt > ResolveNearbyHostDistanceToleranceFt(curve) Then Continue For

                    AddUniqueHostCurve(results, seen, curve, domain)
                Next
            Next
        End Sub

        Private Shared Function CollectNearbyHostCurves(candidate As Element, domain As String) As List(Of MEPCurve)
            Dim results As New List(Of MEPCurve)()
            If candidate Is Nothing OrElse candidate.Document Is Nothing Then Return results

            Try
                Dim collector As FilteredElementCollector = New FilteredElementCollector(candidate.Document).OfClass(GetType(MEPCurve)).WhereElementIsNotElementType()
                Dim outline = BuildExpandedOutline(candidate, NearbyHostSearchMarginFt)
                If outline IsNot Nothing Then
                    collector = collector.WherePasses(New BoundingBoxIntersectsFilter(outline))
                End If

                For Each el As Element In collector
                    Dim curve = TryCast(el, MEPCurve)
                    If curve Is Nothing Then Continue For
                    If Not MatchesCurveDomain(curve, domain) Then Continue For
                    results.Add(curve)
                Next
            Catch
            End Try

            Return results
        End Function

        Private Shared Function BuildExpandedOutline(el As Element, marginFt As Double) As Outline
            If el Is Nothing Then Return Nothing

            Try
                Dim bbox = el.BoundingBox(Nothing)
                If bbox Is Nothing OrElse bbox.Min Is Nothing OrElse bbox.Max Is Nothing Then Return Nothing

                Dim minPoint = New XYZ(bbox.Min.X - marginFt, bbox.Min.Y - marginFt, bbox.Min.Z - marginFt)
                Dim maxPoint = New XYZ(bbox.Max.X + marginFt, bbox.Max.Y + marginFt, bbox.Max.Z + marginFt)
                Return New Outline(minPoint, maxPoint)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ResolveNearbyHostDistanceToleranceFt(curve As MEPCurve) As Double
            Dim halfSize = ResolveMepCurveHalfSizeFt(curve)
            Dim tolerance = Math.Max(NearbyHostBaseToleranceFt, halfSize + NearbyHostBaseToleranceFt)
            Return Math.Min(NearbyHostMaxToleranceFt, tolerance)
        End Function

        Private Shared Function ResolveMepCurveHalfSizeFt(curve As MEPCurve) As Double
            If curve Is Nothing Then Return 0.0R

            Try
                If TypeOf curve Is Autodesk.Revit.DB.Plumbing.Pipe Then
                    Dim diameter = ReadDoubleParameter(curve, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                    If diameter > 0.0R Then Return diameter / 2.0R
                End If

                If TypeOf curve Is Autodesk.Revit.DB.Mechanical.Duct Then
                    Dim diameter = ReadDoubleParameter(curve, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                    If diameter > 0.0R Then Return diameter / 2.0R

                    Dim width = ReadDoubleParameter(curve, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                    Dim height = ReadDoubleParameter(curve, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                    If width > 0.0R AndAlso height > 0.0R Then
                        Return Math.Sqrt((width * width) + (height * height)) / 2.0R
                    End If
                    If width > 0.0R Then Return width / 2.0R
                    If height > 0.0R Then Return height / 2.0R
                End If
            Catch
            End Try

            Return 0.0R
        End Function

        Private Shared Function ReadDoubleParameter(el As Element, bip As BuiltInParameter) As Double
            If el Is Nothing Then Return 0.0R

            Try
                Dim param = el.Parameter(bip)
                If param Is Nothing OrElse param.StorageType <> StorageType.Double Then Return 0.0R
                Return param.AsDouble()
            Catch
                Return 0.0R
            End Try
        End Function

        Private Shared Function ReadProjectionDepthStandards(el As Element) As List(Of ProjectionDepthStandard)
            Dim standards As New List(Of ProjectionDepthStandard)()
            If el Is Nothing Then Return standards

            Dim valueFt As Double = 0.0R
            If TryReadTakeoffLengthProjectionFt(el, valueFt) Then
                standards.Add(New ProjectionDepthStandard With {
                    .Name = "Takeoff Length Projection",
                    .ValueFt = valueFt
                })
            End If

            If TryReadTakeoffLengthFt(el, valueFt) Then
                standards.Add(New ProjectionDepthStandard With {
                    .Name = "Takeoff Length",
                    .ValueFt = valueFt
                })
            End If

            Return standards
        End Function

        Private Shared Function TryReadTakeoffLengthProjectionFt(el As Element, ByRef valueFt As Double) As Boolean
            valueFt = 0.0R
            If el Is Nothing Then Return False

            Dim param As Parameter = Nothing
            If TryGetBuiltInParameterOnElementOrType(el, BuiltInParameter.RBS_FAMILY_CONTENT_TAKEOFF_PROJLENGTH, param) AndAlso
               TryReadPositiveDoubleParameter(param, valueFt) Then
                Return True
            End If

            For Each name In New String() {"Takeoff Length Projection", "Takeoff Length Proj.", "Takeoff Length Proj", "Takeoff Projection Length"}
                param = FindParameterOnElementOrType(el, name)
                If TryReadPositiveDoubleParameter(param, valueFt) Then Return True
            Next

            Return False
        End Function

        Private Shared Function TryReadTakeoffLengthFt(el As Element, ByRef valueFt As Double) As Boolean
            valueFt = 0.0R
            If el Is Nothing Then Return False

            For Each name In New String() {"Takeoff Length", "Takeoff Len.", "Takeoff Len", "TakeoffLength"}
                Dim param = FindParameterOnElementOrType(el, name)
                If TryReadPositiveDoubleParameter(param, valueFt) Then Return True
            Next

            Return False
        End Function

        Private Shared Function TryReadPositiveDoubleParameter(param As Parameter, ByRef valueFt As Double) As Boolean
            valueFt = 0.0R
            If param Is Nothing Then Return False

            Try
                If param.StorageType = StorageType.Double Then
                    valueFt = param.AsDouble()
                    Return valueFt > 0.0R
                End If
            Catch
                valueFt = 0.0R
            End Try

            Return False
        End Function

        Private Shared Function TryEvaluateHostBranch(connectors As IList(Of Connector),
                                                      host As MEPCurve,
                                                      ByRef hit As ReviewHit) As Boolean
            hit = Nothing
            If connectors Is Nothing OrElse connectors.Count = 0 OrElse host Is Nothing Then Return False

            Dim centerLine As Line = Nothing
            If Not TryGetCenterLine(host, centerLine) Then Return False

            Dim hostDirection = NormalizeVector(centerLine.Direction)
            If hostDirection Is Nothing Then Return False

            Dim insertionConnector As Connector = Nothing
            Dim insertionOrigin As XYZ = Nothing
            Dim hostPointDistanceFt As Double = Double.NaN
            If Not TryResolveInsertionConnector(connectors,
                                                centerLine.GetEndPoint(0),
                                                hostDirection,
                                                insertionConnector,
                                                insertionOrigin,
                                                hostPointDistanceFt) Then
                Return False
            End If

            Dim branchAxis = ResolveBranchAxis(connectors, insertionConnector, hostDirection)
            If branchAxis Is Nothing Then Return False

            Dim angleDeg = AngleBetweenAxesDeg(branchAxis, hostDirection)
            If Double.IsNaN(angleDeg) OrElse Double.IsInfinity(angleDeg) Then Return False
            If angleDeg <= BranchAngleThresholdDeg Then Return False

            Dim distanceFt = ComputeConnectorMissDistance(insertionOrigin,
                                                          branchAxis,
                                                          centerLine.GetEndPoint(0),
                                                          hostDirection)
            If Double.IsNaN(distanceFt) OrElse Double.IsInfinity(distanceFt) Then Return False

            hit = New ReviewHit With {
                .Host = host,
                .DistanceFt = distanceFt,
                .AngleDeg = angleDeg,
                .DisplayAngleDeg = ResolveModeledAngleDeg(branchAxis),
                .HostPointDistanceFt = hostPointDistanceFt
            }
            Return True
        End Function

        Private Shared Function TryEvaluateProjectionDepth(connectors As IList(Of Connector),
                                                           host As MEPCurve,
                                                           ByRef hit As ProjectionDepthHit) As Boolean
            hit = Nothing
            If connectors Is Nothing OrElse connectors.Count = 0 OrElse host Is Nothing Then Return False

            Dim centerLine As Line = Nothing
            If Not TryGetCenterLine(host, centerLine) Then Return False

            Dim hostDirection = NormalizeVector(centerLine.Direction)
            If hostDirection Is Nothing Then Return False

            Dim insertionConnector As Connector = Nothing
            Dim insertionOrigin As XYZ = Nothing
            Dim hostPointDistanceFt As Double = Double.NaN
            If Not TryResolveInsertionConnector(connectors,
                                                centerLine.GetEndPoint(0),
                                                hostDirection,
                                                insertionConnector,
                                                insertionOrigin,
                                                hostPointDistanceFt) Then
                Return False
            End If

            Dim branchAxis = ResolveBranchAxis(connectors, insertionConnector, hostDirection)
            If branchAxis Is Nothing Then Return False

            Dim angleDeg = AngleBetweenAxesDeg(branchAxis, hostDirection)
            If Double.IsNaN(angleDeg) OrElse Double.IsInfinity(angleDeg) Then Return False
            If angleDeg <= BranchAngleThresholdDeg Then Return False

            Dim radialDirection = ResolveOutwardRadialDirection(insertionOrigin, candidateOwner:=insertionConnector.Owner, centerLine:=centerLine, hostDirection:=hostDirection, branchAxis:=branchAxis)
            If radialDirection Is Nothing Then Return False

            Dim hostHalfSizeFt = ResolveHostHalfSizeAlongDirectionFt(host, centerLine, radialDirection)
            If hostHalfSizeFt <= 0.0R Then Return False

            Dim actualBuriedFt As Double = 0.0R
            Dim radialDistanceFt As Double = 0.0R
            If Not TryComputeProjectionDepthFromGeometry(insertionConnector.Owner,
                                                         centerLine,
                                                         radialDirection,
                                                         hostHalfSizeFt,
                                                         actualBuriedFt,
                                                         radialDistanceFt) Then
                radialDistanceFt = DistanceFromPointToInfiniteLine(insertionOrigin,
                                                                   centerLine.GetEndPoint(0),
                                                                   hostDirection)
                If Double.IsNaN(radialDistanceFt) OrElse Double.IsInfinity(radialDistanceFt) Then Return False
                actualBuriedFt = hostHalfSizeFt - radialDistanceFt
            End If

            hit = New ProjectionDepthHit With {
                .Host = host,
                .ActualBuriedFt = actualBuriedFt,
                .RadialDistanceFt = radialDistanceFt,
                .HostHalfSizeFt = hostHalfSizeFt,
                .HostPointDistanceFt = hostPointDistanceFt
            }
            Return True
        End Function

        Private Shared Sub ApplyProjectionDepthStandards(hit As ProjectionDepthHit,
                                                         standards As IList(Of ProjectionDepthStandard))
            If hit Is Nothing OrElse standards Is Nothing Then Return

            Dim bestStandard As ProjectionDepthStandard = Nothing
            Dim bestDifference As Double = Double.MaxValue

            For Each depthStandard In standards
                If depthStandard Is Nothing OrElse depthStandard.ValueFt <= 0.0R Then Continue For

                Dim difference = hit.ActualBuriedFt - depthStandard.ValueFt
                If String.Equals(depthStandard.Name, "Takeoff Length Projection", StringComparison.OrdinalIgnoreCase) Then
                    hit.HasProjectionLength = True
                    hit.ProjectionLengthFt = depthStandard.ValueFt
                    hit.ProjectionDifferenceFt = difference
                ElseIf String.Equals(depthStandard.Name, "Takeoff Length", StringComparison.OrdinalIgnoreCase) Then
                    hit.HasTakeoffLength = True
                    hit.TakeoffLengthFt = depthStandard.ValueFt
                    hit.TakeoffDifferenceFt = difference
                End If

                If bestStandard Is Nothing OrElse Math.Abs(difference) < Math.Abs(bestDifference) Then
                    bestStandard = depthStandard
                    bestDifference = difference
                End If
            Next

            If bestStandard IsNot Nothing Then
                hit.StandardName = bestStandard.Name
                hit.StandardFt = bestStandard.ValueFt
                hit.DifferenceFt = bestDifference
            End If
        End Sub

        Private Shared Function ResolveOutwardRadialDirection(insertionOrigin As XYZ,
                                                              candidateOwner As Element,
                                                              centerLine As Line,
                                                              hostDirection As XYZ,
                                                              branchAxis As XYZ) As XYZ
            If centerLine Is Nothing OrElse hostDirection Is Nothing OrElse branchAxis Is Nothing Then Return Nothing

            Dim projectedBranch = branchAxis.Subtract(hostDirection.Multiply(branchAxis.DotProduct(hostDirection)))
            Dim radialDirection = NormalizeVector(projectedBranch)
            If radialDirection Is Nothing Then Return Nothing

            Dim probePoint As XYZ = Nothing
            If Not TryGetElementBoundingBoxCenter(candidateOwner, probePoint) Then
                probePoint = insertionOrigin
            End If

            If probePoint IsNot Nothing Then
                Dim foot = ProjectPointToInfiniteLine(probePoint, centerLine.GetEndPoint(0), hostDirection)
                If foot IsNot Nothing Then
                    Dim fromLine = probePoint.Subtract(foot)
                    If fromLine IsNot Nothing AndAlso fromLine.GetLength() > 0.0000001R AndAlso fromLine.DotProduct(radialDirection) < 0.0R Then
                        radialDirection = radialDirection.Multiply(-1.0R)
                    End If
                End If
            End If

            Return radialDirection
        End Function

        Private Shared Function ResolveHostHalfSizeAlongDirectionFt(host As MEPCurve,
                                                                    centerLine As Line,
                                                                    radialDirection As XYZ) As Double
            If host Is Nothing OrElse centerLine Is Nothing OrElse radialDirection Is Nothing Then Return 0.0R

            Try
                Dim origin = centerLine.GetEndPoint(0)
                Dim points = CollectElementGeometryPoints(host)
                Dim maxAbs = 0.0R

                For Each point In points
                    If point Is Nothing Then Continue For

                    Dim coord = point.Subtract(origin).DotProduct(radialDirection)
                    If Double.IsNaN(coord) OrElse Double.IsInfinity(coord) Then Continue For

                    maxAbs = Math.Max(maxAbs, Math.Abs(coord))
                Next

                If maxAbs > 0.0000001R Then Return maxAbs
            Catch
            End Try

            Return ResolveMepCurveHalfSizeFt(host)
        End Function

        Private Shared Function TryComputeProjectionDepthFromGeometry(candidate As Element,
                                                                      centerLine As Line,
                                                                      radialDirection As XYZ,
                                                                      hostHalfSizeFt As Double,
                                                                      ByRef actualBuriedFt As Double,
                                                                      ByRef radialDistanceFt As Double) As Boolean
            actualBuriedFt = 0.0R
            radialDistanceFt = Double.NaN
            If candidate Is Nothing OrElse centerLine Is Nothing OrElse radialDirection Is Nothing OrElse hostHalfSizeFt <= 0.0R Then Return False

            Dim points = CollectElementGeometryPoints(candidate)
            If points Is Nothing OrElse points.Count = 0 Then Return False

            Dim origin = centerLine.GetEndPoint(0)
            Dim minCoord = Double.MaxValue

            For Each point In points
                If point Is Nothing Then Continue For

                Dim coord = point.Subtract(origin).DotProduct(radialDirection)
                If Double.IsNaN(coord) OrElse Double.IsInfinity(coord) Then Continue For
                If coord < minCoord Then minCoord = coord
            Next

            If minCoord = Double.MaxValue Then Return False

            radialDistanceFt = Math.Abs(minCoord)
            actualBuriedFt = Math.Max(0.0R, hostHalfSizeFt - minCoord)
            Return True
        End Function

        Private Shared Function TryGetElementBoundingBoxCenter(el As Element, ByRef center As XYZ) As Boolean
            center = Nothing
            If el Is Nothing Then Return False

            Try
                Dim bbox = el.BoundingBox(Nothing)
                If bbox Is Nothing OrElse bbox.Min Is Nothing OrElse bbox.Max Is Nothing Then Return False

                center = New XYZ((bbox.Min.X + bbox.Max.X) / 2.0R,
                                 (bbox.Min.Y + bbox.Max.Y) / 2.0R,
                                 (bbox.Min.Z + bbox.Max.Z) / 2.0R)
                Return True
            Catch
                center = Nothing
                Return False
            End Try
        End Function

        Private Shared Function CollectElementGeometryPoints(el As Element) As List(Of XYZ)
            Dim points As New List(Of XYZ)()
            If el Is Nothing Then Return points

            Try
                Dim opt As New Options() With {
                    .ComputeReferences = False,
                    .DetailLevel = ViewDetailLevel.Fine,
                    .IncludeNonVisibleObjects = False
                }

                Dim geom = GetGeometryCompat(el, opt)
                CollectGeometryPoints(geom, Nothing, points)
            Catch
            End Try

            Return points
        End Function

        Private Shared Function GetGeometryCompat(el As Element, opt As Options) As GeometryElement
            If el Is Nothing Then Return Nothing
            If opt Is Nothing Then opt = New Options()

            Try
                Dim t = el.GetType()
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

                Return TryCast(mi.Invoke(el, New Object() {opt}), GeometryElement)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub CollectGeometryPoints(geom As GeometryElement,
                                                 xform As Transform,
                                                 acc As IList(Of XYZ))
            If geom Is Nothing OrElse acc Is Nothing Then Return

            For Each go As GeometryObject In geom
                If go Is Nothing Then Continue For

                Dim solid = TryCast(go, Solid)
                If solid IsNot Nothing Then
                    AddSolidEdgePoints(solid, xform, acc)
                    Continue For
                End If

                Dim curve = TryCast(go, Curve)
                If curve IsNot Nothing Then
                    AddCurvePoints(curve, xform, acc)
                    Continue For
                End If

                Dim instance = TryCast(go, GeometryInstance)
                If instance IsNot Nothing Then
                    Try
                        Dim instGeom As GeometryElement = Nothing
                        Try
                            instGeom = instance.GetInstanceGeometry()
                        Catch
                        End Try

                        If instGeom IsNot Nothing Then
                            CollectGeometryPoints(instGeom, xform, acc)
                        Else
                            Dim symbolGeom As GeometryElement = Nothing
                            Try
                                symbolGeom = instance.GetSymbolGeometry()
                            Catch
                            End Try

                            If symbolGeom IsNot Nothing Then
                                Dim nextXform = If(xform Is Nothing, instance.Transform, xform.Multiply(instance.Transform))
                                CollectGeometryPoints(symbolGeom, nextXform, acc)
                            End If
                        End If
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Shared Sub AddSolidEdgePoints(solid As Solid,
                                              xform As Transform,
                                              acc As IList(Of XYZ))
            If solid Is Nothing OrElse acc Is Nothing Then Return

            Try
                If solid.Edges Is Nothing Then Return

                For Each edge As Edge In solid.Edges
                    If edge Is Nothing Then Continue For

                    Dim curve As Curve = Nothing
                    Try
                        curve = edge.AsCurve()
                    Catch
                        curve = Nothing
                    End Try

                    If curve IsNot Nothing Then AddCurvePoints(curve, xform, acc)
                Next
            Catch
            End Try
        End Sub

        Private Shared Sub AddCurvePoints(curve As Curve,
                                          xform As Transform,
                                          acc As IList(Of XYZ))
            If curve Is Nothing OrElse acc Is Nothing Then Return

            Try
                Dim pts = curve.Tessellate()
                If pts Is Nothing Then Return

                For Each point As XYZ In pts
                    If point Is Nothing Then Continue For

                    If xform IsNot Nothing Then
                        Try
                            point = xform.OfPoint(point)
                        Catch
                        End Try
                    End If

                    acc.Add(point)
                Next
            Catch
            End Try
        End Sub

        Private Shared Function ProjectPointToInfiniteLine(point As XYZ,
                                                           linePoint As XYZ,
                                                           lineDirection As XYZ) As XYZ
            If point Is Nothing OrElse linePoint Is Nothing OrElse lineDirection Is Nothing Then Return Nothing

            Try
                Dim delta = point.Subtract(linePoint)
                Dim along = delta.DotProduct(lineDirection)
                Return linePoint.Add(lineDirection.Multiply(along))
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryResolveInsertionConnector(connectors As IList(Of Connector),
                                                             hostPoint As XYZ,
                                                             hostDirection As XYZ,
                                                             ByRef insertionConnector As Connector,
                                                             ByRef insertionOrigin As XYZ,
                                                             ByRef pointDistanceFt As Double) As Boolean
            insertionConnector = Nothing
            insertionOrigin = Nothing
            pointDistanceFt = Double.NaN
            If connectors Is Nothing OrElse connectors.Count = 0 Then Return False

            Dim bestDistance = Double.MaxValue
            For Each connector In connectors
                Dim origin As XYZ = Nothing
                If Not TryGetConnectorOrigin(connector, origin) Then Continue For

                Dim distanceFt = DistanceFromPointToInfiniteLine(origin, hostPoint, hostDirection)
                If Double.IsNaN(distanceFt) OrElse Double.IsInfinity(distanceFt) Then Continue For

                If insertionConnector Is Nothing OrElse distanceFt < bestDistance Then
                    insertionConnector = connector
                    insertionOrigin = origin
                    bestDistance = distanceFt
                End If
            Next

            If insertionConnector Is Nothing OrElse insertionOrigin Is Nothing Then Return False
            pointDistanceFt = bestDistance
            Return True
        End Function

        Private Shared Function ResolveBranchAxis(connectors As IList(Of Connector),
                                                  insertionConnector As Connector,
                                                  hostDirection As XYZ) As XYZ
            If connectors Is Nothing OrElse insertionConnector Is Nothing Then Return Nothing

            Dim insertionOrigin As XYZ = Nothing
            If Not TryGetConnectorOrigin(insertionConnector, insertionOrigin) Then Return Nothing

            Dim bestAxis As XYZ = Nothing
            Dim bestLength As Double = 0.0R
            For Each connector In connectors
                If connector Is Nothing OrElse Object.ReferenceEquals(connector, insertionConnector) Then Continue For

                Dim origin As XYZ = Nothing
                If Not TryGetConnectorOrigin(connector, origin) Then Continue For

                Dim delta = origin.Subtract(insertionOrigin)
                Dim length = delta.GetLength()
                If length <= bestLength OrElse length <= 0.0000001R Then Continue For

                Dim axis = NormalizeVector(delta)
                If axis Is Nothing Then Continue For

                Dim angleToHost = AngleBetweenAxesDeg(axis, hostDirection)
                If Double.IsNaN(angleToHost) OrElse Double.IsInfinity(angleToHost) Then Continue For
                If angleToHost <= BranchAngleThresholdDeg Then Continue For

                bestAxis = axis
                bestLength = length
            Next

            If bestAxis IsNot Nothing Then Return bestAxis

            For Each axis In GetConnectorAxisCandidates(insertionConnector)
                Dim angleToHost = AngleBetweenAxesDeg(axis, hostDirection)
                If Double.IsNaN(angleToHost) OrElse Double.IsInfinity(angleToHost) Then Continue For
                If angleToHost > BranchAngleThresholdDeg Then Return axis
            Next

            Return Nothing
        End Function

        Private Shared Function TryResolveBranchAxisFromConnectorLayout(connectors As IList(Of Connector),
                                                                        ByRef branchAxis As XYZ) As Boolean
            branchAxis = Nothing
            If connectors Is Nothing OrElse connectors.Count = 0 Then Return False

            Dim bestLength As Double = 0.0R
            For i As Integer = 0 To connectors.Count - 1
                Dim originA As XYZ = Nothing
                If Not TryGetConnectorOrigin(connectors(i), originA) Then Continue For

                For j As Integer = i + 1 To connectors.Count - 1
                    Dim originB As XYZ = Nothing
                    If Not TryGetConnectorOrigin(connectors(j), originB) Then Continue For

                    Dim delta = originB.Subtract(originA)
                    Dim length = delta.GetLength()
                    If length <= bestLength OrElse length <= 0.0000001R Then Continue For

                    Dim axis = NormalizeVector(delta)
                    If axis Is Nothing Then Continue For

                    branchAxis = axis
                    bestLength = length
                Next
            Next

            If branchAxis IsNot Nothing Then Return True

            For Each connector In connectors
                branchAxis = GetConnectorAxis(connector)
                If branchAxis IsNot Nothing Then Return True
            Next

            Return False
        End Function

        Private Shared Function BuildRow(doc As Document,
                                         candidate As Element,
                                         hit As ReviewHit,
                                         unit As String,
                                         extraParams As IList(Of String)) As Dictionary(Of String, Object)
            Dim status = If(String.IsNullOrWhiteSpace(hit.Status), "ERROR", hit.Status)
            Dim message = If(String.IsNullOrWhiteSpace(hit.Message), MisalignedMessage, hit.Message)
            Dim row As New Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                {"File", BuildFileLabel(doc)},
                {"ElementId", SafeElementId(candidate)},
                {"Category", SafeCategoryName(candidate)},
                {"Family", GetFamilyName(candidate)},
                {"Type", GetTypeName(candidate)},
                {"HostId", SafeElementId(hit.Host)},
                {"HostCategory", SafeCategoryName(hit.Host)},
                {"HostType", GetTypeName(hit.Host)},
                {"Domain", ResolveDomainLabel(hit.Host)},
                {"DistanceFromCenter", FormatNumber(ToDisplayDistance(hit.DistanceFt, unit), 3)},
                {"ModeledAngle", FormatNumber(ResolveDisplayAngleValue(hit), 3)},
                {"Status", status},
                {"Message", message}
            }

            If extraParams IsNot Nothing Then
                For Each name In extraParams
                    row("BranchParam::" & name) = ResolveExtraValue(candidate, name)
                    row("HostParam::" & name) = ResolveExtraValue(hit.Host, name)
                Next
            End If

            Return row
        End Function

        Private Shared Function BuildProjectionDepthRow(doc As Document,
                                                        candidate As Element,
                                                        hit As ProjectionDepthHit,
                                                        unit As String,
            extraParams As IList(Of String)) As Dictionary(Of String, Object)
            Dim status = If(String.IsNullOrWhiteSpace(hit.Status), "ERROR", hit.Status)
            Dim message = If(String.IsNullOrWhiteSpace(hit.Message), "Takeoff Length Projection / Takeoff Length 기준 묻힘 깊이를 확인해 주세요.", hit.Message)
            Dim row As New Dictionary(Of String, Object)(StringComparer.Ordinal) From {
                {"File", BuildFileLabel(doc)},
                {"ElementId", SafeElementId(candidate)},
                {"Category", SafeCategoryName(candidate)},
                {"Family", GetFamilyName(candidate)},
                {"Type", GetTypeName(candidate)},
                {"HostId", SafeElementId(hit.Host)},
                {"HostCategory", SafeCategoryName(hit.Host)},
                {"HostType", GetTypeName(hit.Host)},
                {"Domain", ResolveDomainLabel(hit.Host)},
                {"ProjectionLength", FormatOptionalDistance(hit.HasProjectionLength, hit.ProjectionLengthFt, unit)},
                {"TakeoffLength", FormatOptionalDistance(hit.HasTakeoffLength, hit.TakeoffLengthFt, unit)},
                {"StandardName", If(String.IsNullOrWhiteSpace(hit.StandardName), String.Empty, hit.StandardName)},
                {"StandardLength", If(hit.StandardFt > 0.0R, FormatNumber(ToDisplayDistance(hit.StandardFt, unit), 3), String.Empty)},
                {"ActualBuriedLength", FormatNumber(ToDisplayDistance(hit.ActualBuriedFt, unit), 3)},
                {"Difference", FormatNumber(ToDisplayDistance(hit.DifferenceFt, unit), 3)},
                {"ProjectionDifference", FormatOptionalDistance(hit.HasProjectionLength, hit.ProjectionDifferenceFt, unit)},
                {"TakeoffDifference", FormatOptionalDistance(hit.HasTakeoffLength, hit.TakeoffDifferenceFt, unit)},
                {"DistanceFromCenter", FormatNumber(ToDisplayDistance(hit.RadialDistanceFt, unit), 3)},
                {"HostHalfSize", FormatNumber(ToDisplayDistance(hit.HostHalfSizeFt, unit), 3)},
                {"HostConnectionMissing", If(hit.HostConnectionMissing, "True", "False")},
                {"Status", status},
                {"Message", message}
            }

            If extraParams IsNot Nothing Then
                For Each name In extraParams
                    row("BranchParam::" & name) = ResolveExtraValue(candidate, name)
                    row("HostParam::" & name) = ResolveExtraValue(hit.Host, name)
                Next
            End If

            Return row
        End Function

        Private Shared Function CollectCandidates(doc As Document, domain As String) As List(Of Element)
            Dim categoryIds As New List(Of BuiltInCategory)()
            If domain = "all" OrElse domain = "pipe" Then
                categoryIds.Add(BuiltInCategory.OST_PipeFitting)
                categoryIds.Add(BuiltInCategory.OST_PipeAccessory)
            End If
            If domain = "all" OrElse domain = "duct" Then
                categoryIds.Add(BuiltInCategory.OST_DuctFitting)
                categoryIds.Add(BuiltInCategory.OST_DuctAccessory)
            End If

            Dim results As New List(Of Element)()
            For Each catId In categoryIds
                For Each el As Element In New FilteredElementCollector(doc).OfCategory(catId).WhereElementIsNotElementType()
                    If HasConnectors(el) AndAlso IsInlineBranchCandidate(el, domain) Then results.Add(el)
                Next
            Next
            Return results
        End Function

        Private Shared Function IsInlineBranchCandidate(el As Element, domain As String) As Boolean
            If el Is Nothing Then Return False

            Dim family = TryCast(el, FamilyInstance)
            If family Is Nothing Then Return False

            Dim partType As PartType
            If TryGetInlineBranchPartType(family, partType) Then
                If MatchesInlineBranchPartType(family, partType, domain) Then Return True
            End If

            Return IsInlineBranchByMidCurveConnection(family, domain)
        End Function

        Private Shared Function IsInlineBranchByMidCurveConnection(family As FamilyInstance, domain As String) As Boolean
            If family Is Nothing Then Return False

            Dim connectors = GetConnectors(family)
            If connectors.Count < 2 Then Return False

            Return GetMidCurveConnectedHostCurves(connectors, domain).Count > 0
        End Function

        Private Shared Function GetMidCurveConnectedHostCurves(connectors As IList(Of Connector),
                                                               domain As String) As List(Of MEPCurve)
            Dim results As New List(Of MEPCurve)()
            Dim seen As New HashSet(Of Integer)()
            If connectors Is Nothing OrElse connectors.Count < 2 Then Return results

            For Each connector In connectors
                Dim origin As XYZ = Nothing
                If Not TryGetConnectorOrigin(connector, origin) Then Continue For

                For Each curve In GetConnectedHostCurves(connector, domain)
                    Dim centerLine As Line = Nothing
                    If Not TryGetCenterLine(curve, centerLine) Then Continue For

                    Dim hostDirection = NormalizeVector(centerLine.Direction)
                    If hostDirection Is Nothing Then Continue For
                    If Not IsPointInsideCurveSpan(origin, centerLine) Then Continue For
                    If ResolveBranchAxis(connectors, connector, hostDirection) Is Nothing Then Continue For

                    AddUniqueHostCurve(results, seen, curve, domain)
                Next
            Next

            Return results
        End Function

        Private Shared Function IsPointInsideCurveSpan(point As XYZ, centerLine As Line) As Boolean
            If point Is Nothing OrElse centerLine Is Nothing Then Return False

            Dim startPoint = centerLine.GetEndPoint(0)
            Dim endPoint = centerLine.GetEndPoint(1)
            If startPoint Is Nothing OrElse endPoint Is Nothing Then Return False

            Dim direction = NormalizeVector(endPoint.Subtract(startPoint))
            If direction Is Nothing Then Return False

            Dim length = startPoint.DistanceTo(endPoint)
            If length <= 0.0000001R Then Return False

            Dim along = point.Subtract(startPoint).DotProduct(direction)
            Dim endpointTol = Math.Min(MidCurveEndpointToleranceFt, length * 0.25R)

            Return along > endpointTol AndAlso along < (length - endpointTol)
        End Function

        Private Shared Function TryGetInlineBranchPartType(el As Element, ByRef partType As PartType) As Boolean
            partType = PartType.Undefined
            If el Is Nothing Then Return False

            Dim param As Parameter = Nothing
            If TryGetBuiltInParameterOnElementOrType(el, BuiltInParameter.RBS_PART_TYPE, param) AndAlso TryReadPartType(param, partType) Then
                Return True
            End If

            If TryGetBuiltInParameterOnElementOrType(el, BuiltInParameter.FAMILY_CONTENT_PART_TYPE, param) AndAlso TryReadPartType(param, partType) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function TryGetBuiltInParameterOnElementOrType(el As Element,
                                                                      bip As BuiltInParameter,
                                                                      ByRef param As Parameter) As Boolean
            param = Nothing
            If el Is Nothing Then Return False

            Try
                param = el.Parameter(bip)
                If param IsNot Nothing Then Return True
            Catch
            End Try

            Try
                Dim typeId = el.GetTypeId()
                If typeId IsNot Nothing AndAlso typeId.IntValue() > 0 Then
                    Dim typeEl = el.Document.GetElement(typeId)
                    If typeEl IsNot Nothing Then
                        param = typeEl.Parameter(bip)
                        If param IsNot Nothing Then Return True
                    End If
                End If
            Catch
            End Try

            Return False
        End Function

        Private Shared Function TryReadPartType(param As Parameter, ByRef partType As PartType) As Boolean
            partType = PartType.Undefined
            If param Is Nothing Then Return False

            Try
                If param.StorageType = StorageType.Integer Then
                    Dim raw = param.AsInteger()
                    If [Enum].IsDefined(GetType(PartType), raw) Then
                        partType = CType(raw, PartType)
                        Return True
                    End If
                End If
            Catch
            End Try

            Dim rawText = ResolveParamText(param)
            If String.IsNullOrWhiteSpace(rawText) Then Return False

            For Each name As String In [Enum].GetNames(GetType(PartType))
                If rawText.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    partType = CType([Enum].Parse(GetType(PartType), name, True), PartType)
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function MatchesInlineBranchPartType(el As Element,
                                                            partType As PartType,
                                                            domain As String) As Boolean
            If el Is Nothing Then Return False

            Select Case NormalizeDomain(domain)
                Case "pipe"
                    If Not IsPipeInlineBranchCategory(el) Then Return False
                    Return IsPipeInlineBranchPartType(partType)
                Case "duct"
                    If Not IsDuctInlineBranchCategory(el) Then Return False
                    Return IsDuctInlineBranchPartType(partType)
                Case Else
                    If IsPipeInlineBranchCategory(el) Then Return IsPipeInlineBranchPartType(partType)
                    If IsDuctInlineBranchCategory(el) Then Return IsDuctInlineBranchPartType(partType)
                    Return IsPipeInlineBranchPartType(partType) OrElse IsDuctInlineBranchPartType(partType)
            End Select
        End Function

        Private Shared Function IsPipeInlineBranchPartType(partType As PartType) As Boolean
            Return partType = PartType.SpudAdjustable OrElse partType = PartType.SpudPerpendicular
        End Function

        Private Shared Function IsDuctInlineBranchPartType(partType As PartType) As Boolean
            Return partType = PartType.TapAdjustable OrElse partType = PartType.TapPerpendicular
        End Function

        Private Shared Function IsPipeInlineBranchCategory(el As Element) As Boolean
            If el Is Nothing OrElse el.Category Is Nothing Then Return False

            Dim catId = CType(el.Category.Id.IntValue(), BuiltInCategory)
            Return catId = BuiltInCategory.OST_PipeFitting OrElse catId = BuiltInCategory.OST_PipeAccessory
        End Function

        Private Shared Function IsDuctInlineBranchCategory(el As Element) As Boolean
            If el Is Nothing OrElse el.Category Is Nothing Then Return False

            Dim catId = CType(el.Category.Id.IntValue(), BuiltInCategory)
            Return catId = BuiltInCategory.OST_DuctFitting OrElse catId = BuiltInCategory.OST_DuctAccessory
        End Function

        Private Shared Function GetConnectedHostCurves(connector As Connector, domain As String) As List(Of MEPCurve)
            Dim results As New List(Of MEPCurve)()
            If connector Is Nothing Then Return results

            Dim seen As New HashSet(Of Integer)()
            Try
                For Each ref As Connector In connector.AllRefs.Cast(Of Connector)()
                    If ref Is Nothing OrElse ref.Owner Is Nothing Then Continue For
                    Dim curve = TryCast(ref.Owner, MEPCurve)
                    If curve Is Nothing Then Continue For
                    AddUniqueHostCurve(results, seen, curve, domain)
                Next
            Catch
            End Try

            AddConnectedHostCurvesFromConnector(connector,
                                                domain,
                                                results,
                                                seen,
                                                New HashSet(Of String)(StringComparer.Ordinal),
                                                0)

            Return results.ToList()
        End Function

        Private Shared Function HasConnectorPathToHost(connectors As IList(Of Connector),
                                                       host As MEPCurve,
                                                       domain As String) As Boolean
            If connectors Is Nothing OrElse connectors.Count = 0 OrElse host Is Nothing OrElse host.Id Is Nothing Then Return False

            Dim hostId = host.Id.IntValue()
            For Each connector In connectors
                If connector Is Nothing Then Continue For

                For Each curve In GetConnectedHostCurves(connector, domain)
                    If curve Is Nothing OrElse curve.Id Is Nothing Then Continue For
                    If curve.Id.IntValue() = hostId Then Return True
                Next
            Next

            Return False
        End Function

        Private Shared Sub AddConnectedHostCurvesFromConnector(connector As Connector,
                                                               domain As String,
                                                               results As IList(Of MEPCurve),
                                                               seen As HashSet(Of Integer),
                                                               visited As HashSet(Of String),
                                                               depth As Integer)
            If connector Is Nothing OrElse results Is Nothing OrElse seen Is Nothing OrElse visited Is Nothing Then Return
            If depth > ConnectedHostTraversalDepth Then Return

            Dim visitKey = BuildConnectorVisitKey(connector)
            If visitKey <> String.Empty AndAlso Not visited.Add(visitKey) Then Return

            Dim ownerCurve = TryCast(connector.Owner, MEPCurve)
            If ownerCurve IsNot Nothing Then
                AddUniqueHostCurve(results, seen, ownerCurve, domain)
            End If

            If depth >= ConnectedHostTraversalDepth Then Return
            If Not IsConnectorConnected(connector) Then Return

            Dim refs = GetConnectorReferences(connector)
            If refs.Count = 0 Then Return

            For Each ref In refs
                If ref Is Nothing OrElse ref.Owner Is Nothing Then Continue For

                Dim curve = TryCast(ref.Owner, MEPCurve)
                If curve IsNot Nothing Then
                    AddUniqueHostCurve(results, seen, curve, domain)
                    Continue For
                End If

                For Each nextConnector In GetConnectors(ref.Owner)
                    AddConnectedHostCurvesFromConnector(nextConnector, domain, results, seen, visited, depth + 1)
                Next
            Next
        End Sub

        Private Shared Function GetConnectorReferences(connector As Connector) As List(Of Connector)
            Dim results As New List(Of Connector)()
            If connector Is Nothing Then Return results

            Try
                For Each ref As Connector In connector.AllRefs.Cast(Of Connector)()
                    If ref Is Nothing Then Continue For
                    results.Add(ref)
                Next
            Catch
            End Try

            Return results
        End Function

        Private Shared Function IsConnectorConnected(connector As Connector) As Boolean
            If connector Is Nothing Then Return False

            Try
                Return connector.IsConnected
            Catch
                Return False
            End Try
        End Function

        Private Shared Function BuildConnectorVisitKey(connector As Connector) As String
            If connector Is Nothing Then Return String.Empty

            Dim ownerId = String.Empty
            Try
                If connector.Owner IsNot Nothing AndAlso connector.Owner.Id IsNot Nothing Then
                    ownerId = connector.Owner.Id.IntValue().ToString(CultureInfo.InvariantCulture)
                End If
            Catch
                ownerId = String.Empty
            End Try

            Dim originText = String.Empty
            Try
                Dim origin = connector.Origin
                If origin IsNot Nothing Then
                    originText = String.Format(CultureInfo.InvariantCulture,
                                               "{0:0.######},{1:0.######},{2:0.######}",
                                               origin.X,
                                               origin.Y,
                                               origin.Z)
                End If
            Catch
                originText = String.Empty
            End Try

            Return ownerId & "|" & originText & "|" & connector.GetHashCode().ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function TryGetHostedCurve(candidate As Element) As MEPCurve
            If candidate Is Nothing Then Return Nothing

            Try
                Dim family = TryCast(candidate, FamilyInstance)
                If family Is Nothing Then Return Nothing
                Return TryCast(family.Host, MEPCurve)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub AddUniqueHostCurve(results As IList(Of MEPCurve),
                                              seen As HashSet(Of Integer),
                                              curve As MEPCurve,
                                              domain As String)
            If results Is Nothing OrElse seen Is Nothing OrElse curve Is Nothing Then Return
            If Not MatchesCurveDomain(curve, domain) Then Return

            Dim id = curve.Id.IntValue()
            If seen.Add(id) Then results.Add(curve)
        End Sub

        Private Shared Function MatchesCurveDomain(curve As MEPCurve, domain As String) As Boolean
            If curve Is Nothing OrElse curve.Category Is Nothing Then Return False
            Dim catId = CType(curve.Category.Id.IntValue(), BuiltInCategory)
            If domain = "pipe" Then Return catId = BuiltInCategory.OST_PipeCurves
            If domain = "duct" Then Return catId = BuiltInCategory.OST_DuctCurves
            Return catId = BuiltInCategory.OST_PipeCurves OrElse catId = BuiltInCategory.OST_DuctCurves
        End Function

        Private Shared Function HasConnectors(el As Element) As Boolean
            Return GetConnectors(el).Count > 0
        End Function

        Private Shared Function GetConnectors(el As Element) As List(Of Connector)
            If el Is Nothing Then Return New List(Of Connector)()

            Try
                Dim fi = TryCast(el, FamilyInstance)
                If fi IsNot Nothing AndAlso fi.MEPModel IsNot Nothing AndAlso fi.MEPModel.ConnectorManager IsNot Nothing Then
                    Return fi.MEPModel.ConnectorManager.Connectors.Cast(Of Connector)().ToList()
                End If
            Catch
            End Try

            Try
                Dim curve = TryCast(el, MEPCurve)
                If curve IsNot Nothing AndAlso curve.ConnectorManager IsNot Nothing Then
                    Return curve.ConnectorManager.Connectors.Cast(Of Connector)().ToList()
                End If
            Catch
            End Try

            Return New List(Of Connector)()
        End Function

        Private Shared Function TryGetConnectorOrigin(connector As Connector,
                                                      ByRef origin As XYZ) As Boolean
            origin = Nothing
            If connector Is Nothing Then Return False

            Try
                origin = connector.Origin
                Return origin IsNot Nothing
            Catch
                origin = Nothing
                Return False
            End Try
        End Function

        Private Shared Function GetConnectorAxis(connector As Connector) As XYZ
            If connector Is Nothing Then Return Nothing

            Try
                Dim basis = NormalizeVector(connector.CoordinateSystem.BasisZ)
                If basis IsNot Nothing Then Return basis
            Catch
            End Try

            Try
                Dim basis = NormalizeVector(connector.CoordinateSystem.BasisX)
                If basis IsNot Nothing Then Return basis
            Catch
            End Try

            Try
                Dim basis = NormalizeVector(connector.CoordinateSystem.BasisY)
                If basis IsNot Nothing Then Return basis
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function GetConnectorAxisCandidates(connector As Connector) As List(Of XYZ)
            Dim results As New List(Of XYZ)()
            If connector Is Nothing Then Return results

            Try
                TryAddConnectorAxisCandidate(results, connector.CoordinateSystem.BasisZ)
            Catch
            End Try

            Try
                TryAddConnectorAxisCandidate(results, connector.CoordinateSystem.BasisX)
            Catch
            End Try

            Try
                TryAddConnectorAxisCandidate(results, connector.CoordinateSystem.BasisY)
            Catch
            End Try

            Return results
        End Function

        Private Shared Sub TryAddConnectorAxisCandidate(results As IList(Of XYZ), candidate As XYZ)
            If results Is Nothing Then Return

            Dim axis = NormalizeVector(candidate)
            If axis Is Nothing Then Return

            For Each existing In results
                If AngleBetweenAxesDeg(existing, axis) <= 0.1R Then Return
            Next

            results.Add(axis)
        End Sub

        Private Shared Function NormalizeVector(vector As XYZ) As XYZ
            If vector Is Nothing Then Return Nothing
            Try
                If vector.GetLength() <= 0.0000001R Then Return Nothing
                Return vector.Normalize()
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryGetCenterLine(curve As MEPCurve, ByRef line As Line) As Boolean
            line = Nothing
            If curve Is Nothing Then Return False

            Try
                Dim location = TryCast(curve.Location, LocationCurve)
                If location Is Nothing OrElse location.Curve Is Nothing Then Return False
                line = TryCast(location.Curve, Line)
            Catch
                line = Nothing
            End Try

            Return line IsNot Nothing
        End Function

        Private Shared Function AngleBetweenAxesDeg(a As XYZ, b As XYZ) As Double
            Dim na = NormalizeVector(a)
            Dim nb = NormalizeVector(b)
            If na Is Nothing OrElse nb Is Nothing Then Return Double.NaN
            Dim dot = Math.Abs(na.DotProduct(nb))
            dot = Math.Max(-1.0R, Math.Min(1.0R, dot))
            Return Math.Acos(dot) * (180.0R / Math.PI)
        End Function

        Private Shared Function ResolveModeledAngleDeg(connectorAxis As XYZ) As Double
            Return AngleToXyPlaneDeg(connectorAxis)
        End Function

        Private Shared Function ResolveDisplayAngleValue(hit As ReviewHit) As Double
            If hit Is Nothing Then Return Double.NaN
            If Not Double.IsNaN(hit.DisplayAngleDeg) AndAlso Not Double.IsInfinity(hit.DisplayAngleDeg) Then
                Return hit.DisplayAngleDeg
            End If
            Return hit.AngleDeg
        End Function

        Private Shared Function AngleToXyPlaneDeg(vector As XYZ) As Double
            Dim direction = NormalizeVector(vector)
            If direction Is Nothing Then Return Double.NaN

            Dim zComponent = Math.Abs(direction.Z)
            zComponent = Math.Max(0.0R, Math.Min(1.0R, zComponent))
            Return Math.Asin(zComponent) * (180.0R / Math.PI)
        End Function

        Private Shared Function ComputeConnectorMissDistance(origin As XYZ,
                                                             connectorAxis As XYZ,
                                                             hostPoint As XYZ,
                                                             hostDirection As XYZ) As Double
            Dim axis = NormalizeVector(connectorAxis)
            Dim hostDir = NormalizeVector(hostDirection)
            If origin Is Nothing OrElse hostPoint Is Nothing OrElse axis Is Nothing OrElse hostDir Is Nothing Then Return Double.NaN

            Dim foot = ProjectPointOntoLine(origin, hostPoint, hostDir)
            If foot Is Nothing Then Return Double.NaN

            Dim offset = foot.Subtract(origin)
            If offset Is Nothing Then Return Double.NaN

            Dim offsetLength = offset.GetLength()
            If offsetLength <= 0.0000001R Then Return 0.0R

            Dim projectedAxis = NormalizeVector(ProjectVectorOntoPlane(axis, hostDir))
            If projectedAxis Is Nothing Then Return offsetLength

            Return offset.CrossProduct(projectedAxis).GetLength()
        End Function

        Private Shared Function ProjectPointOntoLine(point As XYZ,
                                                     lineOrigin As XYZ,
                                                     lineDirection As XYZ) As XYZ
            Dim direction = NormalizeVector(lineDirection)
            If point Is Nothing OrElse lineOrigin Is Nothing OrElse direction Is Nothing Then Return Nothing

            Dim delta = point.Subtract(lineOrigin)
            Dim along = delta.DotProduct(direction)
            Return lineOrigin.Add(direction.Multiply(along))
        End Function

        Private Shared Function ProjectVectorOntoPlane(vector As XYZ, planeNormal As XYZ) As XYZ
            Dim direction = NormalizeVector(vector)
            Dim normal = NormalizeVector(planeNormal)
            If direction Is Nothing OrElse normal Is Nothing Then Return Nothing

            Return direction.Subtract(normal.Multiply(direction.DotProduct(normal)))
        End Function

        Private Shared Function DistanceBetweenLineSegments(startA As XYZ,
                                                            endA As XYZ,
                                                            startB As XYZ,
                                                            endB As XYZ) As Double
            If startA Is Nothing OrElse endA Is Nothing OrElse startB Is Nothing OrElse endB Is Nothing Then Return Double.NaN

            Dim u = endA.Subtract(startA)
            Dim v = endB.Subtract(startB)
            Dim w = startA.Subtract(startB)

            Dim a = u.DotProduct(u)
            Dim b = u.DotProduct(v)
            Dim c = v.DotProduct(v)
            Dim d = u.DotProduct(w)
            Dim e = v.DotProduct(w)
            Dim denominator = (a * c) - (b * b)
            Dim epsilon As Double = 0.0000001R

            Dim sN As Double
            Dim sD As Double = denominator
            Dim tN As Double
            Dim tD As Double = denominator

            If denominator < epsilon Then
                sN = 0.0R
                sD = 1.0R
                tN = e
                tD = c
            Else
                sN = (b * e) - (c * d)
                tN = (a * e) - (b * d)

                If sN < 0.0R Then
                    sN = 0.0R
                    tN = e
                    tD = c
                ElseIf sN > sD Then
                    sN = sD
                    tN = e + b
                    tD = c
                End If
            End If

            If tN < 0.0R Then
                tN = 0.0R
                If -d < 0.0R Then
                    sN = 0.0R
                ElseIf -d > a Then
                    sN = sD
                Else
                    sN = -d
                    sD = a
                End If
            ElseIf tN > tD Then
                tN = tD
                If (-d + b) < 0.0R Then
                    sN = 0.0R
                ElseIf (-d + b) > a Then
                    sN = sD
                Else
                    sN = -d + b
                    sD = a
                End If
            End If

            Dim sc = If(Math.Abs(sN) < epsilon, 0.0R, sN / sD)
            Dim tc = If(Math.Abs(tN) < epsilon, 0.0R, tN / tD)
            Dim delta = w.Add(u.Multiply(sc)).Subtract(v.Multiply(tc))
            Return delta.GetLength()
        End Function

        Private Shared Function DistanceBetweenInfiniteLines(originA As XYZ,
                                                             directionA As XYZ,
                                                             originB As XYZ,
                                                             directionB As XYZ) As Double
            Dim da = NormalizeVector(directionA)
            Dim db = NormalizeVector(directionB)
            If originA Is Nothing OrElse originB Is Nothing OrElse da Is Nothing OrElse db Is Nothing Then Return Double.NaN

            Dim cross = da.CrossProduct(db)
            Dim crossLength = cross.GetLength()
            If crossLength <= 0.0000001R Then
                Return originB.Subtract(originA).CrossProduct(da).GetLength()
            End If

            Dim delta = originB.Subtract(originA)
            Return Math.Abs(delta.DotProduct(cross)) / crossLength
        End Function

        Private Shared Function DistanceFromPointToInfiniteLine(point As XYZ,
                                                                lineOrigin As XYZ,
                                                                lineDirection As XYZ) As Double
            Dim direction = NormalizeVector(lineDirection)
            If point Is Nothing OrElse lineOrigin Is Nothing OrElse direction Is Nothing Then Return Double.NaN

            Dim delta = point.Subtract(lineOrigin)
            Return delta.CrossProduct(direction).GetLength()
        End Function

        Private Shared Function NormalizeUnit(unit As String) As String
            Dim normalized = If(unit, String.Empty).Trim().ToLowerInvariant()
            If normalized = "inch" OrElse normalized = "in" OrElse normalized = "inches" Then Return "inch"
            Return "mm"
        End Function

        Private Shared Function NormalizeDomain(domain As String) As String
            Dim normalized = If(domain, String.Empty).Trim().ToLowerInvariant()
            If normalized = "pipe" OrElse normalized = "piping" Then Return "pipe"
            If normalized = "duct" OrElse normalized = "hvac" Then Return "duct"
            Return "all"
        End Function

        Private Shared Function NormalizeExtraParams(extraParams As IEnumerable(Of String)) As List(Of String)
            Dim results As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If extraParams Is Nothing Then Return results

            For Each raw In extraParams
                Dim name = If(raw, String.Empty).Trim()
                If name = String.Empty Then Continue For
                If seen.Add(name) Then results.Add(name)
            Next

            Return results
        End Function

        Private Shared Function ResolveExtraValue(el As Element, name As String) As String
            Dim syntheticValue As String = String.Empty
            If TryResolveSyntheticExtraValue(el, name, syntheticValue) Then
                Return syntheticValue
            End If
            Return ResolveParamText(el, name)
        End Function

        Private Shared Function TryResolveSyntheticExtraValue(el As Element,
                                                              rawName As String,
                                                              ByRef value As String) As Boolean
            value = String.Empty
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(rawName) Then Return False

            Dim normalized = NormalizeSyntheticName(rawName)
            If normalized = String.Empty Then Return False

            Dim point As XYZ = Nothing
            Dim direction As XYZ = Nothing
            Dim lengthFt As Double = 0.0R

            Select Case normalized
                Case "pointx"
                    If TryGetRepresentativePoint(el, point) Then value = FormatNumber(point.X, 6)
                    Return True
                Case "pointy"
                    If TryGetRepresentativePoint(el, point) Then value = FormatNumber(point.Y, 6)
                    Return True
                Case "pointz"
                    If TryGetRepresentativePoint(el, point) Then value = FormatNumber(point.Z, 6)
                    Return True
                Case "curvelength", "length", "linelength"
                    If TryGetCurveLengthFt(el, lengthFt) Then value = FormatNumber(lengthFt, 6)
                    Return True
                Case "directionx", "vectorx", "dirx"
                    If TryGetDirectionVector(el, direction) Then value = FormatNumber(direction.X, 6)
                    Return True
                Case "directiony", "vectory", "diry"
                    If TryGetDirectionVector(el, direction) Then value = FormatNumber(direction.Y, 6)
                    Return True
                Case "directionz", "vectorz", "dirz"
                    If TryGetDirectionVector(el, direction) Then value = FormatNumber(direction.Z, 6)
                    Return True
            End Select

            Return False
        End Function

        Private Shared Function NormalizeSyntheticName(rawName As String) As String
            If String.IsNullOrWhiteSpace(rawName) Then Return String.Empty
            Return New String(rawName.Trim().ToLowerInvariant().Where(Function(ch) Char.IsLetterOrDigit(ch)).ToArray())
        End Function

        Private Shared Function TryGetRepresentativePoint(el As Element, ByRef point As XYZ) As Boolean
            point = Nothing
            If el Is Nothing Then Return False

            Try
                Dim locationPoint = TryCast(el.Location, LocationPoint)
                If locationPoint IsNot Nothing AndAlso locationPoint.Point IsNot Nothing Then
                    point = locationPoint.Point
                    Return True
                End If
            Catch
            End Try

            Try
                Dim locationCurve = TryCast(el.Location, LocationCurve)
                If locationCurve IsNot Nothing AndAlso locationCurve.Curve IsNot Nothing Then
                    point = locationCurve.Curve.Evaluate(0.5R, True)
                    If point IsNot Nothing Then Return True
                End If
            Catch
            End Try

            Try
                Dim connectors = GetConnectors(el)
                If connectors.Count = 0 Then Return False

                Dim sumX As Double = 0.0R
                Dim sumY As Double = 0.0R
                Dim sumZ As Double = 0.0R
                Dim count As Integer = 0

                For Each connector In connectors
                    Dim origin As XYZ = Nothing
                    Try
                        origin = connector.Origin
                    Catch
                        origin = Nothing
                    End Try
                    If origin Is Nothing Then Continue For

                    sumX += origin.X
                    sumY += origin.Y
                    sumZ += origin.Z
                    count += 1
                Next

                If count > 0 Then
                    point = New XYZ(sumX / count, sumY / count, sumZ / count)
                    Return True
                End If
            Catch
            End Try

            Return False
        End Function

        Private Shared Function TryGetCurveLengthFt(el As Element, ByRef lengthFt As Double) As Boolean
            lengthFt = 0.0R
            If el Is Nothing Then Return False

            Try
                Dim locationCurve = TryCast(el.Location, LocationCurve)
                If locationCurve Is Nothing OrElse locationCurve.Curve Is Nothing Then Return False
                lengthFt = locationCurve.Curve.Length
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function TryGetDirectionVector(el As Element, ByRef direction As XYZ) As Boolean
            direction = Nothing
            If el Is Nothing Then Return False

            Try
                Dim locationCurve = TryCast(el.Location, LocationCurve)
                Dim line = If(locationCurve Is Nothing, Nothing, TryCast(locationCurve.Curve, Line))
                If line IsNot Nothing Then
                    direction = NormalizeVector(line.Direction)
                    Return direction IsNot Nothing
                End If
            Catch
            End Try

            Try
                For Each connector In GetConnectors(el)
                    direction = GetConnectorAxis(connector)
                    If direction IsNot Nothing Then Return True
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function ResolveParamText(el As Element, name As String) As String
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return String.Empty
            Return ResolveParamText(FindParameterOnElementOrType(el, name))
        End Function

        Private Shared Function ResolveParamTexts(el As Element, name As String) As List(Of String)
            Dim results As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return results

            For Each param In FindParametersOnElementOrType(el, name)
                Dim text = ResolveParamText(param)
                If seen.Add(text) Then results.Add(text)
            Next

            Return results
        End Function

        Private Shared Function FindParameterOnElementOrType(el As Element, name As String) As Parameter
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing

            Dim parameters = FindParametersOnElementOrType(el, name)
            If parameters Is Nothing OrElse parameters.Count = 0 Then Return Nothing
            Return parameters(0)
        End Function

        Private Shared Function FindParametersOnElementOrType(el As Element, name As String) As List(Of Parameter)
            Dim results As New List(Of Parameter)()
            If el Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return results

            Try
                Dim instanceParams = el.GetParameters(name)
                If instanceParams IsNot Nothing Then
                    For Each param In instanceParams
                        If param IsNot Nothing Then results.Add(param)
                    Next
                End If
            Catch
            End Try

            Try
                Dim typeId = el.GetTypeId()
                If typeId IsNot Nothing AndAlso typeId.IntValue() > 0 Then
                    Dim elementType = el.Document.GetElement(typeId)
                    If elementType IsNot Nothing Then
                        Dim typeParams = elementType.GetParameters(name)
                        If typeParams IsNot Nothing Then
                            For Each param In typeParams
                                If param IsNot Nothing Then results.Add(param)
                            Next
                        End If
                    End If
                End If
            Catch
            End Try

            Return results
        End Function

        Private Shared Function ResolveParamText(param As Parameter) As String
            If param Is Nothing Then Return String.Empty

            Dim hasValue As Boolean = False
            Try
                hasValue = param.HasValue
            Catch
                hasValue = False
            End Try
            If Not hasValue Then Return String.Empty

            Dim raw As String = Nothing
            Try
                If param.StorageType = StorageType.[String] Then
                    raw = param.AsString()
                Else
                    raw = param.AsValueString()
                    If String.IsNullOrWhiteSpace(raw) Then raw = param.AsString()
                End If
            Catch
                raw = Nothing
            End Try

            If raw Is Nothing Then raw = String.Empty
            Return raw.Trim()
        End Function

        Private Shared Function ParseTargetFilter(raw As String) As TargetFilter
            Dim result As New TargetFilter()
            If String.IsNullOrWhiteSpace(raw) Then Return result

            Try
                Dim parser As New FilterParser(raw)
                Dim evaluator = parser.Parse()
                If evaluator Is Nothing Then Return result
                result.Evaluator = evaluator
            Catch
                Return result
            End Try

            Return result
        End Function

        Private Shared Function IsElementAllowed(el As Element,
                                                 includeFilter As TargetFilter,
                                                 excludeFilter As TargetFilter,
                                                 excludeEndDummy As Boolean) As Boolean
            If el Is Nothing Then Return False
            If excludeEndDummy AndAlso ShouldExcludeEndDummy(el) Then Return False
            If excludeFilter IsNot Nothing AndAlso excludeFilter.Evaluator IsNot Nothing AndAlso excludeFilter.Evaluator(el) Then Return False
            If includeFilter Is Nothing OrElse includeFilter.Evaluator Is Nothing Then Return True
            Return includeFilter.Evaluator(el)
        End Function

        Private Shared Function ShouldExcludeEndDummy(el As Element) As Boolean
            Dim familyName = GetFamilyName(el)
            If familyName.IndexOf("Dummy", StringComparison.OrdinalIgnoreCase) < 0 Then Return False

            Return familyName.IndexOf("End_", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   familyName.IndexOf("_End", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   familyName.IndexOf("End-", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   familyName.IndexOf("-End", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   familyName.IndexOf("End ", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   familyName.IndexOf(" End", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function ResolveDomainLabel(el As Element) As String
            If el Is Nothing OrElse el.Category Is Nothing Then Return String.Empty
            Dim catId = CType(el.Category.Id.IntValue(), BuiltInCategory)
            If catId = BuiltInCategory.OST_PipeCurves Then Return "Pipe"
            If catId = BuiltInCategory.OST_DuctCurves Then Return "Duct"
            Return el.Category.Name
        End Function

        Private Shared Function ToDisplayDistance(valueFt As Double, unit As String) As Double
            Select Case NormalizeUnit(unit)
                Case "inch"
                    Return valueFt * 12.0R
                Case Else
                    Return valueFt * 304.8R
            End Select
        End Function

        Private Shared Function FormatNumber(value As Double, decimals As Integer) As String
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then Return String.Empty
            Return Math.Round(value, decimals).ToString("0." & New String("#"c, decimals), CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatOptionalDistance(hasValue As Boolean, valueFt As Double, unit As String) As String
            If Not hasValue Then Return String.Empty
            Return FormatNumber(ToDisplayDistance(valueFt, unit), 3)
        End Function

        Private Shared Function ParseNumber(value As String) As Double
            Dim parsed As Double
            If Double.TryParse(value, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, parsed) Then
                Return parsed
            End If
            Return Double.MinValue
        End Function

        Private Shared Function ParseInteger(value As String) As Integer
            Dim parsed As Integer
            If Integer.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                Return parsed
            End If
            Return Integer.MaxValue
        End Function

        Private Shared Function BuildFileLabel(doc As Document) As String
            If doc Is Nothing Then Return String.Empty
            If Not String.IsNullOrWhiteSpace(doc.PathName) Then Return Path.GetFileName(doc.PathName)
            Return doc.Title
        End Function

        Private Shared Function SafeElementId(el As Element) As String
            If el Is Nothing OrElse el.Id Is Nothing Then Return String.Empty
            Return el.Id.IntValue().ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function SafeCategoryName(el As Element) As String
            If el Is Nothing OrElse el.Category Is Nothing Then Return String.Empty
            Return el.Category.Name
        End Function

        Private Shared Function GetFamilyName(el As Element) As String
            If el Is Nothing Then Return String.Empty
            Try
                Dim fi = TryCast(el, FamilyInstance)
                If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing AndAlso fi.Symbol.Family IsNot Nothing Then
                    Return fi.Symbol.Family.Name
                End If

                Dim elementType = TryCast(el.Document.GetElement(el.GetTypeId()), ElementType)
                If elementType IsNot Nothing Then Return elementType.FamilyName
            Catch
            End Try
            Return String.Empty
        End Function

        Private Shared Function GetTypeName(el As Element) As String
            If el Is Nothing Then Return String.Empty
            Try
                Dim fi = TryCast(el, FamilyInstance)
                If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then Return fi.Symbol.Name

                Dim elementType = TryCast(el.Document.GetElement(el.GetTypeId()), ElementType)
                If elementType IsNot Nothing Then Return elementType.Name
            Catch
            End Try
            Return String.Empty
        End Function

        Private Shared Function ReadField(row As Dictionary(Of String, Object), key As String) As String
            If row Is Nothing OrElse String.IsNullOrWhiteSpace(key) OrElse Not row.ContainsKey(key) Then Return String.Empty
            Return If(row(key), String.Empty).ToString()
        End Function

    End Class

End Namespace
