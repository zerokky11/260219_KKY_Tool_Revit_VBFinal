Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application
Imports WpfBorder = System.Windows.Controls.Border
Imports WpfButton = System.Windows.Controls.Button
Imports WpfCheckBox = System.Windows.Controls.CheckBox
Imports WpfColor = System.Windows.Media.Color
Imports WpfComboBox = System.Windows.Controls.ComboBox
Imports WpfControl = System.Windows.Controls.Control
Imports WpfFrameworkElement = System.Windows.FrameworkElement
Imports WpfPanel = System.Windows.Controls.Panel
Imports WpfTabItem = System.Windows.Controls.TabItem
Imports WpfTextBlock = System.Windows.Controls.TextBlock

Namespace UI

    Friend NotInheritable Class DocumentVisualAidService

        Private Shared ReadOnly SyncRoot As New Object()
        Private Shared ReadOnly RefreshDelay As TimeSpan = TimeSpan.FromMilliseconds(180)
        Private Shared ReadOnly SharedSyncDelay As TimeSpan = TimeSpan.FromMilliseconds(900)
        Private Shared ReadOnly SettingsDirectoryPath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY_Tool_Revit")
        Private Shared ReadOnly SettingsFilePath As String =
            Path.Combine(SettingsDirectoryPath, "document-visual-aid.settings")
        Private Shared ReadOnly Palette As WpfColor() = {
            WpfColor.FromRgb(73, 152, 255),
            WpfColor.FromRgb(62, 186, 121),
            WpfColor.FromRgb(244, 153, 54),
            WpfColor.FromRgb(212, 91, 117),
            WpfColor.FromRgb(130, 118, 255),
            WpfColor.FromRgb(46, 188, 196),
            WpfColor.FromRgb(214, 126, 209),
            WpfColor.FromRgb(170, 166, 84)
        }

        Private Shared ReadOnly ColorSlotByKey As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        Private Shared _refreshTimer As DispatcherTimer
        Private Shared _sharedSyncTimer As DispatcherTimer
        Private Shared _legendWindow As DocumentColorLegendWindow
        Private Shared _revitApp As RevitApp
        Private Shared _pendingActiveKey As String = String.Empty
        Private Shared _nextColorSlot As Integer
        Private Shared _started As Boolean
        Private Shared _enabled As Boolean = True
        Private Shared _settingsLoaded As Boolean
        Private Shared _legendWindowSuppressed As Boolean
        Private Shared _lastEntries As List(Of DocumentColorEntry) = New List(Of DocumentColorEntry)()
        Private Shared _lastNavigator As DocumentViewNavigatorSnapshot

        Friend Shared ReadOnly Property IsEnabled As Boolean
            Get
                EnsureSettingsLoaded()

                SyncLock SyncRoot
                    Return _enabled
                End SyncLock
            End Get
        End Property

        Friend Shared Function CreateHostSettingsPayload() As Object
            Return New With {
                .enabled = IsEnabled
            }
        End Function

        Public Shared Sub Start()
            EnsureSettingsLoaded()

            SyncLock SyncRoot
                If _started Then Return
                _started = True
            End SyncLock

            DocumentViewNavigationBridge.Initialize()

            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then Return

            dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(
                                       Sub()
                                           EnsureRefreshTimer()
                                           EnsureSharedSyncTimer()
                                           _sharedSyncTimer.Start()
                                       End Sub))
        End Sub

        Public Shared Sub [Stop]()
            SyncLock SyncRoot
                _started = False
                _pendingActiveKey = String.Empty
            End SyncLock

            DocumentViewNavigationBridge.[Stop]()

            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then Return

            dispatcher.Invoke(Sub()
                                  If _refreshTimer IsNot Nothing Then
                                      _refreshTimer.Stop()
                                  End If

                                  If _sharedSyncTimer IsNot Nothing Then
                                      _sharedSyncTimer.Stop()
                                  End If

                                  If _legendWindow IsNot Nothing Then
                                      _legendWindow.RequestClose()
                                      _legendWindow = Nothing
                                  End If
                              End Sub)

            DocumentVisualAidSharedSessionCoordinator.ClearLocalState()
        End Sub

        Public Shared Sub NotifyActiveDocumentChanged(doc As Document)
            If doc Is Nothing Then
                QueueRefresh(Nothing)
                Return
            End If

            Try
                _revitApp = doc.Application
            Catch
            End Try

            QueueRefresh(doc)
        End Sub

        Public Shared Sub NotifyDocumentOpened(doc As Document)
            If doc IsNot Nothing Then
                Try
                    _revitApp = doc.Application
                Catch
                End Try
            End If

            QueueRefresh(doc)
        End Sub

        Public Shared Sub NotifyDocumentClosed()
            QueueRefresh(Nothing)
        End Sub

        Friend Shared Sub SetEnabled(enabled As Boolean)
            EnsureSettingsLoaded()

            SyncLock SyncRoot
                _enabled = enabled
                If enabled Then
                    _legendWindowSuppressed = False
                End If
            End SyncLock

            SaveEnabledSetting(enabled)

            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then
                ApplyEnabledVisualState(enabled)
                Return
            End If

            dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(
                                       Sub()
                                           ApplyEnabledVisualState(enabled)
                                       End Sub))
        End Sub

        Private Shared Sub QueueRefresh(activeDoc As Document)
            Start()

            SyncLock SyncRoot
                If activeDoc IsNot Nothing Then
                    _pendingActiveKey = BuildDocumentKey(activeDoc)
                Else
                    _pendingActiveKey = String.Empty
                End If
            End SyncLock

            TriggerRefresh()
        End Sub

        Private Shared Sub TriggerRefresh()
            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then Return

            dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(
                                       Sub()
                                           EnsureRefreshTimer()
                                           _refreshTimer.Stop()
                                           _refreshTimer.Start()
                                       End Sub))
        End Sub

        Private Shared Sub EnsureRefreshTimer()
            If _refreshTimer IsNot Nothing Then Return

            _refreshTimer = New DispatcherTimer(DispatcherPriority.Background)
            _refreshTimer.Interval = RefreshDelay
            AddHandler _refreshTimer.Tick, AddressOf OnRefreshTimerTick
        End Sub

        Private Shared Sub OnRefreshTimerTick(sender As Object, e As EventArgs)
            If _refreshTimer IsNot Nothing Then
                _refreshTimer.Stop()
            End If

            RefreshVisualAids()
        End Sub

        Private Shared Sub EnsureSharedSyncTimer()
            If _sharedSyncTimer IsNot Nothing Then Return

            _sharedSyncTimer = New DispatcherTimer(DispatcherPriority.Background)
            _sharedSyncTimer.Interval = SharedSyncDelay
            AddHandler _sharedSyncTimer.Tick, AddressOf OnSharedSyncTimerTick
        End Sub

        Private Shared Sub OnSharedSyncTimerTick(sender As Object, e As EventArgs)
            SyncSharedSessions()
        End Sub

        Private Shared Sub RefreshVisualAids()
            If Not IsEnabled Then
                ApplyEnabledVisualState(False)
                Return
            End If

            Dim entries = BuildEntriesSnapshot()
            Dim navigator = BuildViewNavigatorSnapshot()

            SyncLock SyncRoot
                _lastEntries = entries.ToList()
                _lastNavigator = navigator
            End SyncLock

            RevitDocumentTabStyler.Apply(entries)
            SyncSharedSessions(entries, navigator)
        End Sub

        Private Shared Sub ApplyEnabledVisualState(enabled As Boolean)
            If Not enabled Then
                If _refreshTimer IsNot Nothing Then
                    _refreshTimer.Stop()
                End If

                If _sharedSyncTimer IsNot Nothing Then
                    _sharedSyncTimer.Stop()
                End If

                If _legendWindow IsNot Nothing AndAlso _legendWindow.IsVisible Then
                    _legendWindow.Hide()
                End If

                DocumentVisualAidSharedSessionCoordinator.ClearLocalState()
                RevitDocumentTabStyler.Apply(New List(Of DocumentColorEntry)())
                Return
            End If

            EnsureSharedSyncTimer()
            _sharedSyncTimer.Start()
            TriggerRefresh()
        End Sub

        Friend Shared Sub HideLegendWindowPreservingTabs()
            SyncLock SyncRoot
                _legendWindowSuppressed = True
            End SyncLock

            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then
                If _legendWindow IsNot Nothing AndAlso _legendWindow.IsVisible Then
                    _legendWindow.Hide()
                End If

                Return
            End If

            dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(
                                       Sub()
                                           If _legendWindow IsNot Nothing AndAlso _legendWindow.IsVisible Then
                                               _legendWindow.Hide()
                                           End If
                                       End Sub))
        End Sub

        Private Shared Function IsLegendWindowSuppressed() As Boolean
            SyncLock SyncRoot
                Return _legendWindowSuppressed
            End SyncLock
        End Function

        Friend Shared Sub RequestDocumentActivation(documentKey As String)
            If String.IsNullOrWhiteSpace(documentKey) Then Return

            Dim dispatcher = ResolveDispatcher()
            If dispatcher Is Nothing Then
                RevitDocumentTabStyler.TryActivateDocument(documentKey)
                Return
            End If

            dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(
                                       Sub()
                                           RevitDocumentTabStyler.TryActivateDocument(documentKey)
                                       End Sub))
        End Sub

        Private Shared Sub DispatchDocumentActivation(targetProcessId As Integer, documentKey As String)
            If String.IsNullOrWhiteSpace(documentKey) Then Return

            Dim localProcessId = Process.GetCurrentProcess().Id
            If targetProcessId <= 0 OrElse targetProcessId = localProcessId Then
                RequestDocumentActivation(documentKey)
                FocusSessionProcess(localProcessId)
                Return
            End If

            DocumentVisualAidSharedSessionCoordinator.EnqueueDocumentActivation(targetProcessId, documentKey)
            DocumentVisualAidSharedSessionCoordinator.FocusProcessWindow(targetProcessId)
        End Sub

        Private Shared Sub DispatchNavigateRequested(targetProcessId As Integer, target As DocumentViewOption)
            If target Is Nothing Then Return

            Dim localProcessId = Process.GetCurrentProcess().Id
            If targetProcessId <= 0 OrElse targetProcessId = localProcessId Then
                HandleNavigateRequested(target)
                FocusSessionProcess(localProcessId)
                Return
            End If

            DocumentVisualAidSharedSessionCoordinator.EnqueueViewNavigation(targetProcessId, target)
            DocumentVisualAidSharedSessionCoordinator.FocusProcessWindow(targetProcessId)
        End Sub

        Private Shared Sub FocusSessionProcess(targetProcessId As Integer)
            Dim resolvedProcessId = targetProcessId
            If resolvedProcessId <= 0 Then
                resolvedProcessId = Process.GetCurrentProcess().Id
            End If

            DocumentVisualAidSharedSessionCoordinator.FocusProcessWindow(resolvedProcessId)
        End Sub

        Private Shared Sub SyncSharedSessions(Optional currentEntries As IReadOnlyList(Of DocumentColorEntry) = Nothing,
                                              Optional currentNavigator As DocumentViewNavigatorSnapshot = Nothing)
            If Not IsEnabled Then
                DocumentVisualAidSharedSessionCoordinator.ClearLocalState()
                If _legendWindow IsNot Nothing AndAlso _legendWindow.IsVisible Then
                    _legendWindow.Hide()
                End If
                Return
            End If

            DocumentVisualAidSharedSessionCoordinator.ProcessLocalCommands(AddressOf RequestDocumentActivation,
                                                                           AddressOf HandleNavigateRequested)

            Dim effectiveEntries As IReadOnlyList(Of DocumentColorEntry) = currentEntries
            Dim effectiveNavigator = currentNavigator

            SyncLock SyncRoot
                If effectiveEntries Is Nothing Then
                    effectiveEntries = If(_lastEntries, New List(Of DocumentColorEntry)())
                End If

                If effectiveNavigator Is Nothing Then
                    effectiveNavigator = _lastNavigator
                End If
            End SyncLock

            Dim localState = DocumentVisualAidSharedSessionCoordinator.BuildLocalState(_revitApp, effectiveEntries, effectiveNavigator)
            DocumentVisualAidSharedSessionCoordinator.PublishLocalState(localState)

            Dim aggregate = DocumentVisualAidSharedSessionCoordinator.BuildAggregateState(localState)
            Dim aggregateLocalSession =
                If(aggregate?.Sessions, New List(Of DocumentVisualAidSessionState)()).
                    FirstOrDefault(Function(session) session IsNot Nothing AndAlso session.IsLocalProcess)
            If aggregateLocalSession IsNot Nothing AndAlso aggregateLocalSession.Entries IsNot Nothing Then
                Dim globallyColoredEntries = aggregateLocalSession.Entries.ToList()
                SyncLock SyncRoot
                    _lastEntries = globallyColoredEntries.ToList()
                End SyncLock

                RevitDocumentTabStyler.Apply(globallyColoredEntries)
            End If

            Dim shouldOwnLegend = DocumentVisualAidSharedSessionCoordinator.TryBecomeOwner()
            Dim suppressLegendWindow = IsLegendWindowSuppressed()
            Dim shouldShowLegend As Boolean =
                shouldOwnLegend AndAlso
                aggregate IsNot Nothing AndAlso
                aggregate.Sessions IsNot Nothing AndAlso
                aggregate.Sessions.Count > 0

            If shouldShowLegend Then
                Dim legend = EnsureLegendWindow()
                If legend IsNot Nothing Then
                    legend.UpdateContents(aggregate)
                    If suppressLegendWindow Then
                        If legend.IsVisible Then
                            legend.Hide()
                        End If
                    ElseIf Not legend.IsVisible Then
                        PositionLegendWindow(legend)
                        legend.Show()
                    End If
                End If
            ElseIf _legendWindow IsNot Nothing AndAlso _legendWindow.IsVisible Then
                _legendWindow.Hide()
            End If
        End Sub

        Private Shared Sub EnsureSettingsLoaded()
            Dim shouldLoad As Boolean = False

            SyncLock SyncRoot
                shouldLoad = Not _settingsLoaded
                If shouldLoad Then
                    _settingsLoaded = True
                End If
            End SyncLock

            If Not shouldLoad Then Return

            Dim enabled = LoadEnabledSetting()
            SyncLock SyncRoot
                _enabled = enabled
            End SyncLock
        End Sub

        Private Shared Function LoadEnabledSetting() As Boolean
            Try
                If Not File.Exists(SettingsFilePath) Then
                    Return True
                End If

                Dim raw = File.ReadAllText(SettingsFilePath).Trim()
                If String.IsNullOrWhiteSpace(raw) Then
                    Return True
                End If

                Select Case raw.ToLowerInvariant()
                    Case "0", "false", "off", "disabled"
                        Return False
                    Case "1", "true", "on", "enabled"
                        Return True
                End Select
            Catch
            End Try

            Return True
        End Function

        Private Shared Sub SaveEnabledSetting(enabled As Boolean)
            Try
                Directory.CreateDirectory(SettingsDirectoryPath)
                File.WriteAllText(SettingsFilePath, If(enabled, "true", "false"))
            Catch
            End Try
        End Sub

        Private Shared Function BuildEntriesSnapshot() As List(Of DocumentColorEntry)
            Dim results As New List(Of DocumentColorEntry)()
            Dim app = _revitApp
            If app Is Nothing Then Return results

            Dim activeKey As String = String.Empty
            SyncLock SyncRoot
                activeKey = _pendingActiveKey
            End SyncLock

            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing Then Continue For
                    If Not doc.IsValidObject Then Continue For

                    Dim isLinked As Boolean = False
                    Try
                        isLinked = doc.IsLinked
                    Catch
                    End Try
                    If isLinked Then Continue For

                    Dim entry = CreateEntry(doc)
                    If entry Is Nothing Then Continue For
                    entry.IsActive = Not String.IsNullOrWhiteSpace(activeKey) AndAlso
                                     String.Equals(entry.Key, activeKey, StringComparison.OrdinalIgnoreCase)
                    results.Add(entry)
                Next
            Catch
            End Try

            results.Sort(Function(left, right)
                             Dim slotCompare = left.ColorSlot.CompareTo(right.ColorSlot)
                             If slotCompare <> 0 Then Return slotCompare
                             Return StringComparer.CurrentCultureIgnoreCase.Compare(left.DisplayName, right.DisplayName)
                         End Function)

            Return results
        End Function

        Private Shared Function BuildViewNavigatorSnapshot() As DocumentViewNavigatorSnapshot
            Dim app = _revitApp
            If app Is Nothing Then Return Nothing

            Dim activeKey As String = String.Empty
            SyncLock SyncRoot
                activeKey = _pendingActiveKey
            End SyncLock

            Dim documents As New List(Of DocumentViewNavigatorDocument)()

            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing OrElse Not doc.IsValidObject Then Continue For

                    Dim isLinked As Boolean = False
                    Try
                        isLinked = doc.IsLinked
                    Catch
                    End Try

                    If isLinked Then Continue For

                    Dim navigatorDocument = BuildNavigatorDocument(doc)
                    If navigatorDocument Is Nothing Then Continue For

                    documents.Add(navigatorDocument)
                Next
            Catch
            End Try

            If documents.Count = 0 Then Return Nothing

            Dim selectedKey = activeKey
            If String.IsNullOrWhiteSpace(selectedKey) OrElse
               Not documents.Any(Function(item) String.Equals(item.DocumentKey, selectedKey, StringComparison.OrdinalIgnoreCase)) Then
                selectedKey = documents(0).DocumentKey
            End If

            Return New DocumentViewNavigatorSnapshot With {
                .SelectedDocumentKey = selectedKey,
                .Documents = documents
            }
        End Function

        Private Shared Function BuildNavigatorDocument(doc As Document) As DocumentViewNavigatorDocument
            If doc Is Nothing Then Return Nothing

            Dim documentKey = BuildDocumentKey(doc)
            If String.IsNullOrWhiteSpace(documentKey) Then Return Nothing

            Dim views As New List(Of DocumentViewOption)()
            Dim activeViewId As Integer = -1
            Dim activeCategoryKey As String = "all"

            Try
                If doc.ActiveView IsNot Nothing Then
                    activeViewId = doc.ActiveView.Id.IntegerValue
                    activeCategoryKey = ClassifyViewCategory(doc.ActiveView)
                End If
            Catch
            End Try

            Try
                Dim collector As New FilteredElementCollector(doc)
                For Each view In collector.OfClass(GetType(View)).Cast(Of View)()
                    If Not IsNavigableView(view) Then Continue For

                    views.Add(New DocumentViewOption With {
                        .DocumentKey = documentKey,
                        .ViewId = view.Id.IntegerValue,
                        .CategoryKey = ClassifyViewCategory(view),
                        .DisplayName = BuildViewDisplayName(view)
                    })
                Next
            Catch
            End Try

            views.Sort(AddressOf CompareViewOptions)

            Dim categories As New List(Of DocumentViewCategory)()
            Dim allViews = views.ToList()
            If allViews.Count > 0 Then
                categories.Add(CreateCategory("all", "All", allViews))
            End If

            AddCategoryIfAny(categories, "2d", "2D", views)
            AddCategoryIfAny(categories, "3d", "3D", views)
            AddCategoryIfAny(categories, "schedule", "Schedule", views)
            AddCategoryIfAny(categories, "sheet", "Sheet", views)

            Dim defaultCategoryKey = "all"
            If categories.Any(Function(category) String.Equals(category.Key, activeCategoryKey, StringComparison.OrdinalIgnoreCase)) Then
                defaultCategoryKey = activeCategoryKey
            End If

            Return New DocumentViewNavigatorDocument With {
                .DocumentKey = documentKey,
                .DocumentName = SafeGetDocumentTitle(doc),
                .ActiveViewId = activeViewId,
                .DefaultCategoryKey = defaultCategoryKey,
                .Categories = categories
            }
        End Function

        Private Shared Sub AddCategoryIfAny(target As IList(Of DocumentViewCategory),
                                            key As String,
                                            label As String,
                                            source As IEnumerable(Of DocumentViewOption))
            If target Is Nothing OrElse source Is Nothing Then Return

            Dim items = source.Where(Function(viewOption) String.Equals(viewOption.CategoryKey, key, StringComparison.OrdinalIgnoreCase)).
                               OrderBy(Function(viewOption) viewOption.DisplayName, StringComparer.CurrentCultureIgnoreCase).
                               ToList()
            If items.Count = 0 Then Return

            target.Add(CreateCategory(key, label, items))
        End Sub

        Private Shared Function CreateCategory(key As String,
                                               label As String,
                                               items As IList(Of DocumentViewOption)) As DocumentViewCategory
            Return New DocumentViewCategory With {
                .Key = key,
                .DisplayName = $"{label} ({items.Count})",
                .Views = items.ToList()
            }
        End Function

        Private Shared Function IsNavigableView(view As View) As Boolean
            If view Is Nothing Then Return False

            Try
                If view.IsTemplate Then Return False
            Catch
            End Try

            Select Case view.ViewType
                Case ViewType.ProjectBrowser,
                     ViewType.SystemBrowser,
                     ViewType.Internal,
                     ViewType.Undefined
                    Return False
            End Select

            Return True
        End Function

        Private Shared Function ClassifyViewCategory(view As View) As String
            If view Is Nothing Then Return "2d"
            If TypeOf view Is ViewSheet Then Return "sheet"
            If TypeOf view Is ViewSchedule Then Return "schedule"
            If TypeOf view Is View3D Then Return "3d"
            Return "2d"
        End Function

        Private Shared Function BuildViewDisplayName(view As View) As String
            If view Is Nothing Then Return String.Empty

            Dim prefix = GetViewDisplayPrefix(view)

            Dim name As String = String.Empty
            Try
                name = view.Name
            Catch
                name = String.Empty
            End Try

            Dim sheet = TryCast(view, ViewSheet)
            If sheet IsNot Nothing Then
                Dim sheetNumber As String = String.Empty
                Try
                    sheetNumber = sheet.SheetNumber
                Catch
                    sheetNumber = String.Empty
                End Try

                If Not String.IsNullOrWhiteSpace(sheetNumber) Then
                    name = $"{sheetNumber} - {name}"
                End If
            End If

            If String.IsNullOrWhiteSpace(prefix) Then
                Return name
            End If

            Return $"{prefix} : {name}"
        End Function

        Private Shared Function GetViewDisplayPrefix(view As View) As String
            If view Is Nothing Then Return String.Empty

            Select Case view.ViewType
                Case ViewType.FloorPlan
                    Return "Floor Plan"
                Case ViewType.EngineeringPlan
                    Return "Structural Plan"
                Case ViewType.CeilingPlan
                    Return "Ceiling Plan"
                Case ViewType.AreaPlan
                    Return "Area Plan"
                Case ViewType.Section
                    Return "Section"
                Case ViewType.Elevation
                    Return "Elevation"
                Case ViewType.Detail
                    Return "Detail"
                Case ViewType.DraftingView
                    Return "Drafting View"
                Case ViewType.Legend
                    Return "Legend"
                Case ViewType.Schedule
                    Return "Schedule"
                Case ViewType.DrawingSheet
                    Return "Sheet"
                Case ViewType.ThreeD
                    Return "3D View"
                Case Else
                    Return NormalizeViewTypeLabel(view.ViewType.ToString())
            End Select
        End Function

        Private Shared Function NormalizeViewTypeLabel(rawValue As String) As String
            If String.IsNullOrWhiteSpace(rawValue) Then Return String.Empty

            Dim chars = rawValue.Trim().ToCharArray()
            Dim builder As New System.Text.StringBuilder()
            For index As Integer = 0 To chars.Length - 1
                Dim current = chars(index)
                If index > 0 AndAlso Char.IsUpper(current) AndAlso Not Char.IsWhiteSpace(chars(index - 1)) Then
                    builder.Append(" "c)
                End If
                builder.Append(current)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function CompareViewOptions(left As DocumentViewOption, right As DocumentViewOption) As Integer
            If left Is Nothing AndAlso right Is Nothing Then Return 0
            If left Is Nothing Then Return -1
            If right Is Nothing Then Return 1

            Return StringComparer.CurrentCultureIgnoreCase.Compare(left.DisplayName, right.DisplayName)
        End Function

        Private Shared Sub HandleNavigateRequested(target As DocumentViewOption)
            If target Is Nothing Then Return
            DocumentViewNavigationBridge.RequestNavigate(target)
        End Sub

        Private Shared Function CreateEntry(doc As Document) As DocumentColorEntry
            Dim key = BuildDocumentKey(doc)
            If String.IsNullOrWhiteSpace(key) Then Return Nothing

            Dim slot = GetOrAssignColorSlot(key)
            Dim baseColor = ResolveBaseColor(slot)
            Dim chipColor = BlendWithWhite(baseColor, 0.28R)
            Dim tabColor = BlendWithWhite(baseColor, 0.72R)
            Dim activeTabColor = BlendWithWhite(baseColor, 0.64R)
            Dim borderColor = BlendWithWhite(baseColor, 0.55R)
            Dim accentColor = BlendWithWhite(baseColor, 0.48R)

            Dim title = SafeGetDocumentTitle(doc)
            If String.IsNullOrWhiteSpace(title) Then
                title = "Untitled"
            End If

            Dim fullPath As String = String.Empty
            Try
                fullPath = doc.PathName
            Catch
            End Try

            Dim matchTokens As New List(Of String)()
            AddMatchToken(matchTokens, title)
            If Not String.IsNullOrWhiteSpace(fullPath) Then
                AddMatchToken(matchTokens, IO.Path.GetFileNameWithoutExtension(fullPath))
            End If

            Return New DocumentColorEntry With {
                .Key = key,
                .DisplayName = title,
                .FullPath = fullPath,
                .ColorSlot = slot,
                .ChipBrush = CreateFrozenBrush(WpfColor.FromArgb(230, chipColor.R, chipColor.G, chipColor.B)),
                .TabFillBrush = CreateFrozenBrush(WpfColor.FromArgb(34, tabColor.R, tabColor.G, tabColor.B)),
                .ActiveTabFillBrush = CreateFrozenBrush(WpfColor.FromArgb(52, activeTabColor.R, activeTabColor.G, activeTabColor.B)),
                .BorderBrush = CreateFrozenBrush(WpfColor.FromArgb(145, borderColor.R, borderColor.G, borderColor.B)),
                .AccentBrush = CreateFrozenBrush(WpfColor.FromArgb(160, accentColor.R, accentColor.G, accentColor.B)),
                .MatchTokens = matchTokens
            }
        End Function

        Private Shared Sub AddMatchToken(target As IList(Of String), token As String)
            If target Is Nothing Then Return
            If String.IsNullOrWhiteSpace(token) Then Return

            Dim normalized = token.Trim()
            If normalized.Length = 0 Then Return
            If target.Any(Function(existing) String.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase)) Then Return
            target.Add(normalized)
        End Sub

        Private Shared Function GetOrAssignColorSlot(key As String) As Integer
            SyncLock SyncRoot
                Dim slot As Integer = 0
                If ColorSlotByKey.TryGetValue(key, slot) Then
                    Return slot
                End If

                slot = _nextColorSlot
                _nextColorSlot += 1
                ColorSlotByKey(key) = slot
                Return slot
            End SyncLock
        End Function

        Private Shared Function ResolveBaseColor(slot As Integer) As WpfColor
            Dim normalizedSlot = Math.Max(0, slot)
            If normalizedSlot < Palette.Length Then
                Return Palette(normalizedSlot)
            End If

            Return CreateGeneratedBaseColor(normalizedSlot)
        End Function

        Private Shared Function CreateGeneratedBaseColor(slot As Integer) As WpfColor
            Dim hue = ((CDbl(slot) * 137.508R) Mod 360.0R) / 360.0R
            Dim saturation = 0.62R + (CDbl(slot Mod 3) * 0.06R)
            Dim value = 0.82R - (CDbl((slot \ Palette.Length) Mod 3) * 0.05R)
            Return FromHsv(hue, saturation, value)
        End Function

        Private Shared Function FromHsv(hue As Double, saturation As Double, value As Double) As WpfColor
            Dim normalizedHue = hue - Math.Floor(hue)
            Dim h = normalizedHue * 6.0R
            Dim sector = CInt(Math.Floor(h)) Mod 6
            Dim f = h - Math.Floor(h)
            Dim p = value * (1.0R - saturation)
            Dim q = value * (1.0R - (saturation * f))
            Dim t = value * (1.0R - (saturation * (1.0R - f)))

            Dim r As Double
            Dim g As Double
            Dim b As Double

            Select Case sector
                Case 0
                    r = value
                    g = t
                    b = p
                Case 1
                    r = q
                    g = value
                    b = p
                Case 2
                    r = p
                    g = value
                    b = t
                Case 3
                    r = p
                    g = q
                    b = value
                Case 4
                    r = t
                    g = p
                    b = value
                Case Else
                    r = value
                    g = p
                    b = q
            End Select

            Return WpfColor.FromRgb(ToColorByte(r), ToColorByte(g), ToColorByte(b))
        End Function

        Private Shared Function ToColorByte(component As Double) As Byte
            Dim normalized = Math.Max(0.0R, Math.Min(1.0R, component))
            Return CByte(Math.Max(0, Math.Min(255, CInt(Math.Round(normalized * 255.0R, MidpointRounding.AwayFromZero)))))
        End Function

        Friend Shared Function BuildDocumentKey(doc As Document) As String
            If doc Is Nothing Then Return String.Empty

            Dim fullPath As String = String.Empty
            Try
                fullPath = doc.PathName
            Catch
            End Try

            If Not String.IsNullOrWhiteSpace(fullPath) Then
                Return fullPath.Trim()
            End If

            Dim title = SafeGetDocumentTitle(doc)
            If Not String.IsNullOrWhiteSpace(title) Then
                Return "__title__" & title.Trim()
            End If

            Return String.Empty
        End Function

        Friend Shared Function FindOpenDocument(app As RevitApp, documentKey As String) As Document
            If app Is Nothing OrElse String.IsNullOrWhiteSpace(documentKey) Then Return Nothing

            Try
                For Each doc As Document In app.Documents
                    If doc Is Nothing OrElse Not doc.IsValidObject Then Continue For
                    If String.Equals(BuildDocumentKey(doc), documentKey, StringComparison.OrdinalIgnoreCase) Then
                        Return doc
                    End If
                Next
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Function SafeGetDocumentTitle(doc As Document) As String
            If doc Is Nothing Then Return String.Empty

            Try
                Return If(doc.Title, String.Empty).Trim()
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function EnsureLegendWindow() As DocumentColorLegendWindow
            If _legendWindow Is Nothing Then
                _legendWindow = New DocumentColorLegendWindow(AddressOf DispatchNavigateRequested,
                                                             AddressOf DispatchDocumentActivation,
                                                             AddressOf FocusSessionProcess)
                Dim owner = ResolveRevitMainWindow()
                If owner IsNot Nothing Then
                    Try
                        _legendWindow.Owner = owner
                    Catch
                    End Try
                End If
            End If

            Return _legendWindow
        End Function

        Private Shared Sub PositionLegendWindow(window As Window)
            If window Is Nothing Then Return

            Dim owner = window.Owner
            If owner Is Nothing Then
                owner = ResolveRevitMainWindow()
            End If

            If owner Is Nothing Then Return

            Dim width = If(Double.IsNaN(window.Width) OrElse window.Width <= 0, 260.0R, window.Width)
            window.Left = owner.Left + Math.Max(20.0R, owner.Width - width - 28.0R)
            window.Top = owner.Top + 122.0R
        End Sub

        Private Shared Function ResolveRevitMainWindow() As Window
            Dim app = System.Windows.Application.Current
            If app Is Nothing Then Return Nothing

            For Each candidate As Window In app.Windows
                If candidate Is Nothing Then Continue For
                If String.Equals(candidate.GetType().FullName, "UIFramework.MainWindow", StringComparison.Ordinal) Then
                    Return candidate
                End If
            Next

            Return app.MainWindow
        End Function

        Private Shared Function ResolveDispatcher() As Dispatcher
            Dim app = System.Windows.Application.Current
            If app IsNot Nothing Then
                Return app.Dispatcher
            End If

            Try
                Return Dispatcher.CurrentDispatcher
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function CreateFrozenBrush(color As WpfColor) As SolidColorBrush
            Dim brush As New SolidColorBrush(color)
            brush.Freeze()
            Return brush
        End Function

        Private Shared Function BlendWithWhite(color As WpfColor, amount As Double) As WpfColor
            Dim normalizedAmount = Math.Max(0.0R, Math.Min(1.0R, amount))

            Dim blendComponent =
                Function(component As Byte) As Byte
                    Dim value = CDbl(component) + ((255.0R - CDbl(component)) * normalizedAmount)
                    Return CByte(Math.Max(0, Math.Min(255, CInt(Math.Round(value, MidpointRounding.AwayFromZero)))))
                End Function

            Return WpfColor.FromRgb(blendComponent(color.R),
                                    blendComponent(color.G),
                                    blendComponent(color.B))
        End Function

    End Class

    Friend NotInheritable Class DocumentViewNavigationBridge

        Private Shared ReadOnly SyncRoot As New Object()
        Private Shared _externalEvent As ExternalEvent
        Private Shared _handler As DocumentViewNavigationHandler
        Private Shared _pendingTarget As DocumentViewOption

        Friend Shared Sub Initialize()
            SyncLock SyncRoot
                If _externalEvent IsNot Nothing Then Return

                Try
                    _handler = New DocumentViewNavigationHandler()
                    _externalEvent = ExternalEvent.Create(_handler)
                Catch
                    _handler = Nothing
                    _externalEvent = Nothing
                End Try
            End SyncLock
        End Sub

        Friend Shared Sub [Stop]()
            SyncLock SyncRoot
                _pendingTarget = Nothing
            End SyncLock
        End Sub

        Friend Shared Sub RequestNavigate(target As DocumentViewOption)
            If target Is Nothing Then Return

            SyncLock SyncRoot
                _pendingTarget = target
            End SyncLock

            Try
                If _externalEvent IsNot Nothing Then
                    _externalEvent.Raise()
                End If
            Catch
            End Try
        End Sub

        Friend Shared Function TakePendingTarget() As DocumentViewOption
            SyncLock SyncRoot
                Dim target = _pendingTarget
                _pendingTarget = Nothing
                Return target
            End SyncLock
        End Function

    End Class

    Friend NotInheritable Class DocumentViewNavigationHandler
        Implements IExternalEventHandler

        Public Sub Execute(app As UIApplication) Implements IExternalEventHandler.Execute
            Dim target = DocumentViewNavigationBridge.TakePendingTarget()
            If target Is Nothing OrElse app Is Nothing Then Return

            Dim uiDoc = app.ActiveUIDocument
            Dim activeDoc As Document = Nothing
            If uiDoc IsNot Nothing Then
                activeDoc = uiDoc.Document
            End If

            If activeDoc IsNot Nothing AndAlso
               Not String.Equals(DocumentVisualAidService.BuildDocumentKey(activeDoc), target.DocumentKey, StringComparison.OrdinalIgnoreCase) Then
                If RevitDocumentTabStyler.TryActivateDocument(target.DocumentKey) Then
                    Dim retryTarget = target
                    Dim dispatcher As Dispatcher = Nothing
                    If System.Windows.Application.Current IsNot Nothing Then
                        dispatcher = System.Windows.Application.Current.Dispatcher
                    Else
                        dispatcher = Dispatcher.CurrentDispatcher
                    End If
                    dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                           New Action(
                                               Sub()
                                                   DocumentViewNavigationBridge.RequestNavigate(retryTarget)
                                               End Sub))
                End If
                Return
            End If

            Dim doc = DocumentVisualAidService.FindOpenDocument(app.Application, target.DocumentKey)
            If doc Is Nothing Then Return

            Dim view As View = Nothing
            Try
                view = TryCast(doc.GetElement(New ElementId(target.ViewId)), View)
            Catch
                view = Nothing
            End Try

            If view Is Nothing Then Return

            If uiDoc Is Nothing Then Return

            Try
                uiDoc.RequestViewChange(view)
            Catch
                Try
                    uiDoc.ActiveView = view
                Catch
                End Try
            End Try
        End Sub

        Public Function GetName() As String Implements IExternalEventHandler.GetName
            Return "KKY Document View Navigation"
        End Function

    End Class

    Friend NotInheritable Class RevitDocumentTabStyler

        Private Shared ReadOnly LayoutDocumentTabItemTypeName As String = "Xceed.Wpf.AvalonDock.Controls.LayoutDocumentTabItem"
        Private Shared ReadOnly LayoutDocumentPaneControlTypeName As String = "Xceed.Wpf.AvalonDock.Controls.LayoutDocumentPaneControl"
        Private Shared ReadOnly LayoutDocumentPaneGroupControlTypeName As String = "Xceed.Wpf.AvalonDock.Controls.LayoutDocumentPaneGroupControl"
        Private Shared ReadOnly StyledByKkyProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("StyledByKky",
                                                GetType(Boolean),
                                                GetType(RevitDocumentTabStyler),
                                                New PropertyMetadata(False))
        Private Shared ReadOnly TraceSyncRoot As New Object()
        Private Shared ReadOnly TraceDirectoryPath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY_Tool_Revit")
        Private Shared ReadOnly TraceFilePath As String = Path.Combine(TraceDirectoryPath, "document-tabs.log")

        Friend Shared Sub Apply(entries As IReadOnlyList(Of DocumentColorEntry))
            Dim visualRoots = GetVisualRoots()
            If visualRoots.Count = 0 Then
                TraceApply(0, 0, New List(Of String) From {"no-visual-roots"})
                Return
            End If

            Dim tabItems = FindDocumentTabItems(visualRoots)
            Dim matchedCount As Integer = 0
            Dim sampledTitles As New List(Of String)()

            For Each tabItem In tabItems
                ClearStyling(tabItem)

                Dim fullTitle = TryResolveTabContext(tabItem)
                If sampledTitles.Count < 6 AndAlso Not String.IsNullOrWhiteSpace(fullTitle) Then
                    sampledTitles.Add(fullTitle)
                End If

                Dim matchedEntry As DocumentColorEntry = Nothing
                If entries IsNot Nothing AndAlso entries.Count = 1 Then
                    matchedEntry = entries(0)
                Else
                    matchedEntry = MatchEntry(entries, fullTitle)
                End If
                If matchedEntry Is Nothing Then Continue For

                matchedCount += 1

                Dim dataContext = tabItem.DataContext
                Dim model = GetPropertyValue(dataContext, "Model")
                Dim isActive As Boolean = tabItem.IsSelected OrElse
                                          SafeBool(GetPropertyValue(dataContext, "IsActive")) OrElse
                                          SafeBool(GetPropertyValue(dataContext, "IsSelected")) OrElse
                                          SafeBool(GetPropertyValue(model, "IsActive")) OrElse
                                          SafeBool(GetPropertyValue(model, "IsSelected")) OrElse
                                          matchedEntry.IsActive

                ApplyTemplateFallback(tabItem, matchedEntry, isActive)

                Dim adornerLayer As AdornerLayer = AdornerLayer.GetAdornerLayer(tabItem)
                If adornerLayer IsNot Nothing Then
                    adornerLayer.Add(New DocumentTabColorAdorner(tabItem, matchedEntry, isActive))
                End If
            Next

            TraceApply(tabItems.Count, matchedCount, sampledTitles)
        End Sub

        Friend Shared Function TryActivateDocument(documentKey As String) As Boolean
            If String.IsNullOrWhiteSpace(documentKey) Then Return False

            Dim visualRoots = GetVisualRoots()
            If visualRoots.Count = 0 Then Return False

            Dim matchTokens = BuildDocumentMatchTokens(documentKey)
            If matchTokens.Count = 0 Then Return False

            For Each tabItem In FindDocumentTabItems(visualRoots)
                If tabItem Is Nothing Then Continue For

                Dim fullTitle = TryResolveTabContext(tabItem)
                If Not ContainsAnyToken(fullTitle, matchTokens) Then Continue For

                Try
                    tabItem.IsSelected = True
                Catch
                End Try

                Try
                    tabItem.Focus()
                Catch
                End Try

                Dim selector = TryCast(GetParentObject(tabItem), Primitives.Selector)
                If selector IsNot Nothing Then
                    Try
                        selector.SelectedItem = tabItem
                    Catch
                    End Try
                End If

                Return True
            Next

            Return False
        End Function

        Private Shared Function GetVisualRoots() As List(Of DependencyObject)
            Dim results As New List(Of DependencyObject)()
            Dim seen As New HashSet(Of DependencyObject)()

            Dim app = System.Windows.Application.Current
            If app IsNot Nothing Then
                For Each window As Window In app.Windows
                    If window Is Nothing Then Continue For
                    If Not window.IsLoaded Then Continue For

                    If seen.Add(window) Then
                        results.Add(window)
                    End If
                Next
            End If

            Try
                For Each source As PresentationSource In PresentationSource.CurrentSources
                    If source Is Nothing Then Continue For

                    Dim root = TryCast(source.RootVisual, DependencyObject)
                    If root Is Nothing Then Continue For

                    If seen.Add(root) Then
                        results.Add(root)
                    End If
                Next
            Catch
            End Try

            Return results
        End Function

        Private Shared Function FindDocumentTabItems(visualRoots As IEnumerable(Of DependencyObject)) As List(Of WpfTabItem)
            Dim results As New List(Of WpfTabItem)()
            Dim seen As New HashSet(Of WpfTabItem)()

            If visualRoots Is Nothing Then Return results

            For Each root In visualRoots
                If root Is Nothing Then Continue For

                Dim window = TryCast(root, Window)
                If window IsNot Nothing AndAlso Not window.IsLoaded Then Continue For

                For Each child In EnumerateVisualDescendants(root)
                    If Not String.Equals(child.GetType().FullName, LayoutDocumentPaneGroupControlTypeName, StringComparison.Ordinal) Then
                        Continue For
                    End If

                    For Each descendant In EnumerateVisualDescendants(child)
                        Dim tabItem = TryCast(descendant, WpfTabItem)
                        If tabItem Is Nothing Then Continue For

                        If seen.Add(tabItem) Then
                            results.Add(tabItem)
                        End If
                    Next
                Next

                For Each child In EnumerateVisualDescendants(root)
                    Dim tabItem = TryCast(child, WpfTabItem)
                    If tabItem Is Nothing Then Continue For
                    If Not IsDocumentTabItem(tabItem) Then Continue For

                    If seen.Add(tabItem) Then
                        results.Add(tabItem)
                    End If
                Next
            Next

            Return results
        End Function

        Private Shared Function IsDocumentTabItem(tabItem As WpfTabItem) As Boolean
            If tabItem Is Nothing Then Return False

            Dim typeName = tabItem.GetType().FullName
            If String.Equals(typeName, LayoutDocumentTabItemTypeName, StringComparison.Ordinal) Then
                Return True
            End If

            Return HasAncestorType(tabItem, LayoutDocumentPaneGroupControlTypeName) OrElse
                   HasAncestorType(tabItem, LayoutDocumentPaneControlTypeName)
        End Function

        Private Shared Function TryResolveTabContext(tabItem As WpfTabItem) As String
            If tabItem Is Nothing Then Return String.Empty

            Dim candidates As New List(Of String)()
            AddCandidate(candidates, TryConvertToText(tabItem.ToolTip))
            AddCandidate(candidates, TryConvertToText(tabItem.Header))
            AddContextCandidates(candidates, tabItem.DataContext, 2)

            For Each descendant In EnumerateVisualDescendants(tabItem)
                Dim textBlock = TryCast(descendant, WpfTextBlock)
                If textBlock Is Nothing Then Continue For
                AddCandidate(candidates, textBlock.Text)
            Next

            If candidates.Count = 0 Then Return String.Empty

            Dim distinctCandidates = candidates.Distinct(StringComparer.OrdinalIgnoreCase)
            Return String.Join(" | ", distinctCandidates.ToArray())
        End Function

        Private Shared Sub AddCandidate(target As IList(Of String), candidate As String)
            If target Is Nothing Then Return
            If String.IsNullOrWhiteSpace(candidate) Then Return

            Dim normalized = candidate.Trim()
            If normalized.Length = 0 Then Return
            If target.Any(Function(existing) String.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase)) Then Return

            target.Add(normalized)
        End Sub

        Private Shared Sub AddContextCandidates(target As IList(Of String), source As Object, depth As Integer)
            If target Is Nothing OrElse source Is Nothing OrElse depth < 0 Then Return

            AddCandidate(target, TryConvertToText(source))

            If TypeOf source Is String Then Return
            If TypeOf source Is WpfTextBlock Then
                AddCandidate(target, DirectCast(source, WpfTextBlock).Text)
                Return
            End If

            Dim candidateProperties = {"Title", "ToolTip", "ContentId", "Description", "DocumentTitle", "Header", "Name"}
            For Each propertyName In candidateProperties
                AddCandidate(target, TryConvertToText(GetPropertyValue(source, propertyName)))
            Next

            If depth = 0 Then Return

            Dim nestedProperties = {"Model", "LayoutItem", "Document", "Root"}
            For Each propertyName In nestedProperties
                Dim nestedValue = GetPropertyValue(source, propertyName)
                If nestedValue Is Nothing OrElse Object.ReferenceEquals(nestedValue, source) Then Continue For
                AddContextCandidates(target, nestedValue, depth - 1)
            Next
        End Sub

        Private Shared Function TryConvertToText(value As Object) As String
            If value Is Nothing Then Return String.Empty

            If TypeOf value Is String Then
                Return DirectCast(value, String)
            End If

            Dim textBlock = TryCast(value, WpfTextBlock)
            If textBlock IsNot Nothing Then
                Return textBlock.Text
            End If

            Dim frameworkElement = TryCast(value, WpfFrameworkElement)
            If frameworkElement IsNot Nothing AndAlso TypeOf frameworkElement.ToolTip Is String Then
                Return DirectCast(frameworkElement.ToolTip, String)
            End If

            Try
                Dim text = Convert.ToString(value, CultureInfo.InvariantCulture)
                If String.IsNullOrWhiteSpace(text) Then Return String.Empty
                If String.Equals(text, value.GetType().FullName, StringComparison.Ordinal) Then Return String.Empty
                Return text
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Sub ApplyTemplateFallback(tabItem As WpfTabItem, entry As DocumentColorEntry, isActive As Boolean)
            If tabItem Is Nothing OrElse entry Is Nothing Then Return

            Dim fillBrush = If(isActive, entry.ActiveTabFillBrush, entry.TabFillBrush)
            Dim borderThickness = If(isActive, New Thickness(0, 3, 0, 0), New Thickness(0, 2, 0, 0))

            tabItem.Background = fillBrush
            tabItem.BorderBrush = entry.BorderBrush
            tabItem.BorderThickness = borderThickness
            tabItem.SetValue(StyledByKkyProperty, True)

            Dim primaryPanel As WpfPanel = Nothing
            Dim primaryBorder As WpfBorder = Nothing

            For Each descendant In EnumerateVisualDescendants(tabItem)
                If primaryPanel Is Nothing Then
                    primaryPanel = TryCast(descendant, WpfPanel)
                End If

                If primaryBorder Is Nothing Then
                    primaryBorder = TryCast(descendant, WpfBorder)
                End If

                If primaryPanel IsNot Nothing AndAlso primaryBorder IsNot Nothing Then
                    Exit For
                End If
            Next

            If primaryPanel IsNot Nothing Then
                primaryPanel.Background = fillBrush
                primaryPanel.SetValue(StyledByKkyProperty, True)
            End If

            If primaryBorder IsNot Nothing Then
                primaryBorder.Background = fillBrush
                primaryBorder.BorderBrush = entry.BorderBrush
                primaryBorder.BorderThickness = borderThickness
                primaryBorder.SetValue(StyledByKkyProperty, True)
            End If
        End Sub

        Private Shared Function MatchEntry(entries As IReadOnlyList(Of DocumentColorEntry), fullTitle As String) As DocumentColorEntry
            If entries Is Nothing OrElse entries.Count = 0 Then Return Nothing
            If String.IsNullOrWhiteSpace(fullTitle) Then Return Nothing

            Dim title = fullTitle.Trim()
            Dim bestMatch As DocumentColorEntry = Nothing
            Dim bestScore As Integer = Integer.MinValue

            For Each entry In entries
                If entry Is Nothing OrElse entry.MatchTokens Is Nothing Then Continue For

                For Each token In entry.MatchTokens
                    If String.IsNullOrWhiteSpace(token) Then Continue For

                    Dim score As Integer = Integer.MinValue
                    If title.Equals(token, StringComparison.OrdinalIgnoreCase) Then
                        score = token.Length + 4000
                    ElseIf title.StartsWith(token & " - ", StringComparison.OrdinalIgnoreCase) Then
                        score = token.Length + 3000
                    ElseIf title.IndexOf(token & " - ", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        score = token.Length + 2000
                    ElseIf title.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0 Then
                        score = token.Length + 1000
                    End If

                    If score > bestScore Then
                        bestMatch = entry
                        bestScore = score
                    End If
                Next
            Next

            Return bestMatch
        End Function

        Private Shared Function BuildDocumentMatchTokens(documentKey As String) As List(Of String)
            Dim tokens As New List(Of String)()
            If String.IsNullOrWhiteSpace(documentKey) Then Return tokens

            Dim normalized = documentKey.Trim()
            If normalized.StartsWith("__title__", StringComparison.OrdinalIgnoreCase) Then
                AddCandidate(tokens, normalized.Substring("__title__".Length))
                Return tokens
            End If

            AddCandidate(tokens, IO.Path.GetFileNameWithoutExtension(normalized))
            AddCandidate(tokens, IO.Path.GetFileName(normalized))
            Return tokens
        End Function

        Private Shared Function ContainsAnyToken(fullTitle As String, tokens As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(fullTitle) OrElse tokens Is Nothing Then Return False

            For Each token In tokens
                If String.IsNullOrWhiteSpace(token) Then Continue For
                If fullTitle.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Sub ClearStyling(tabItem As WpfTabItem)
            If tabItem Is Nothing Then Return

            RemoveDocumentAdorners(tabItem)
            tabItem.ClearValue(WpfControl.BackgroundProperty)
            tabItem.ClearValue(WpfControl.BorderBrushProperty)
            tabItem.ClearValue(WpfControl.BorderThicknessProperty)
            tabItem.ClearValue(StyledByKkyProperty)

            For Each descendant In EnumerateVisualDescendants(tabItem)
                If Not SafeBool(GetDependencyValue(descendant, StyledByKkyProperty)) Then Continue For

                Dim panel = TryCast(descendant, WpfPanel)
                If panel IsNot Nothing Then
                    panel.ClearValue(WpfPanel.BackgroundProperty)
                End If

                Dim border = TryCast(descendant, WpfBorder)
                If border IsNot Nothing Then
                    border.ClearValue(WpfBorder.BackgroundProperty)
                    border.ClearValue(WpfBorder.BorderBrushProperty)
                    border.ClearValue(WpfBorder.BorderThicknessProperty)
                End If

                Dim control = TryCast(descendant, WpfControl)
                If control IsNot Nothing Then
                    control.ClearValue(WpfControl.BackgroundProperty)
                    control.ClearValue(WpfControl.BorderBrushProperty)
                    control.ClearValue(WpfControl.BorderThicknessProperty)
                End If

                descendant.ClearValue(StyledByKkyProperty)
            Next
        End Sub

        Private Shared Sub RemoveDocumentAdorners(tabItem As WpfTabItem)
            Dim adornerLayer As AdornerLayer = AdornerLayer.GetAdornerLayer(tabItem)
            If adornerLayer Is Nothing Then Return

            Dim adorners As Adorner() = adornerLayer.GetAdorners(tabItem)
            If adorners Is Nothing Then Return

            For Each adorner In adorners
                If TypeOf adorner Is DocumentTabColorAdorner Then
                    adornerLayer.Remove(adorner)
                End If
            Next
        End Sub

        Private Shared Function GetPropertyValue(instance As Object, propertyName As String) As Object
            If instance Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return Nothing

            Try
                Dim prop = instance.GetType().GetProperty(propertyName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                If prop Is Nothing Then Return Nothing
                Return prop.GetValue(instance, Nothing)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function GetDependencyValue(instance As DependencyObject, [property] As DependencyProperty) As Object
            If instance Is Nothing OrElse [property] Is Nothing Then Return Nothing

            Try
                Return instance.GetValue([property])
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function SafeBool(value As Object) As Boolean
            If value Is Nothing Then Return False

            Try
                If TypeOf value Is Boolean Then
                    Return DirectCast(value, Boolean)
                End If

                Return Convert.ToBoolean(value, CultureInfo.InvariantCulture)
            Catch
                Return False
            End Try
        End Function

        Private Shared Function HasAncestorType(element As DependencyObject, expectedTypeName As String) As Boolean
            If element Is Nothing OrElse String.IsNullOrWhiteSpace(expectedTypeName) Then Return False

            Dim current = GetParentObject(element)
            Do While current IsNot Nothing
                If String.Equals(current.GetType().FullName, expectedTypeName, StringComparison.Ordinal) Then
                    Return True
                End If

                current = GetParentObject(current)
            Loop

            Return False
        End Function

        Private Shared Function GetParentObject(element As DependencyObject) As DependencyObject
            If element Is Nothing Then Return Nothing

            Try
                Dim visualParent = VisualTreeHelper.GetParent(element)
                If visualParent IsNot Nothing Then
                    Return visualParent
                End If
            Catch
            End Try

            Dim frameworkElement = TryCast(element, WpfFrameworkElement)
            If frameworkElement IsNot Nothing Then
                Return TryCast(frameworkElement.Parent, DependencyObject)
            End If

            Return Nothing
        End Function

        Private Shared Sub TraceApply(tabCount As Integer, matchedCount As Integer, sampledTitles As IList(Of String))
            Try
                Dim samples = If(sampledTitles Is Nothing, String.Empty, String.Join(" || ", sampledTitles))
                Dim line = String.Format(CultureInfo.InvariantCulture,
                                         "{0:yyyy-MM-dd HH:mm:ss.fff} tabs={1} matched={2} samples={3}",
                                         DateTime.Now,
                                         tabCount,
                                         matchedCount,
                                         samples)

                SyncLock TraceSyncRoot
                    Directory.CreateDirectory(TraceDirectoryPath)
                    File.AppendAllLines(TraceFilePath, {line})
                End SyncLock
            Catch
            End Try
        End Sub

        Private Shared Iterator Function EnumerateVisualDescendants(root As DependencyObject) As IEnumerable(Of DependencyObject)
            If root Is Nothing Then
                Return
            End If

            Dim childCount As Integer = 0
            Try
                childCount = VisualTreeHelper.GetChildrenCount(root)
            Catch
                childCount = 0
            End Try

            For index As Integer = 0 To childCount - 1
                Dim child As DependencyObject = Nothing
                Try
                    child = VisualTreeHelper.GetChild(root, index)
                Catch
                    child = Nothing
                End Try

                If child Is Nothing Then Continue For

                Yield child

                For Each descendant In EnumerateVisualDescendants(child)
                    Yield descendant
                Next
            Next
        End Function

    End Class

    Friend NotInheritable Class DocumentTabColorAdorner
        Inherits Adorner

        Private ReadOnly _entry As DocumentColorEntry
        Private ReadOnly _isActive As Boolean

        Public Sub New(adornedElement As UIElement, entry As DocumentColorEntry, isActive As Boolean)
            MyBase.New(adornedElement)
            _entry = entry
            _isActive = isActive
            IsHitTestVisible = False
        End Sub

        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            MyBase.OnRender(drawingContext)

            If _entry Is Nothing OrElse drawingContext Is Nothing Then Return

            Dim renderWidth = AdornedElement.RenderSize.Width
            Dim renderHeight = AdornedElement.RenderSize.Height
            If renderWidth <= 4 OrElse renderHeight <= 4 Then Return

            Dim bounds As New Rect(0.75R, 0.75R, Math.Max(0.0R, renderWidth - 1.5R), Math.Max(0.0R, renderHeight - 1.5R))
            Dim fillBrush = If(_isActive, _entry.ActiveTabFillBrush, _entry.TabFillBrush)
            Dim borderPen As New Pen(_entry.BorderBrush, If(_isActive, 1.25R, 1.0R))
            borderPen.Freeze()

            drawingContext.DrawRoundedRectangle(fillBrush, borderPen, bounds, 4.0R, 4.0R)

            Dim accentHeight = Math.Min(4.0R, Math.Max(2.0R, renderHeight / 4.5R))
            Dim accentRect As New Rect(bounds.Left + 1.0R, bounds.Top + 1.0R, Math.Max(0.0R, bounds.Width - 2.0R), accentHeight)
            drawingContext.DrawRoundedRectangle(_entry.AccentBrush, Nothing, accentRect, 3.0R, 3.0R)
        End Sub

    End Class

    Friend NotInheritable Class DocumentColorLegendWindow
        Inherits Window

        Private ReadOnly _itemsHost As New StackPanel()
        Private ReadOnly _summaryText As New TextBlock()
        Private ReadOnly _emptyText As New TextBlock()
        Private ReadOnly _navigatorSummaryText As New TextBlock()
        Private ReadOnly _navigatorEmptyText As New TextBlock()
        Private ReadOnly _documentCombo As New WpfComboBox()
        Private ReadOnly _viewTypeCombo As New WpfComboBox()
        Private ReadOnly _viewCombo As New WpfComboBox()
        Private ReadOnly _alwaysOnTopCheck As New WpfCheckBox()
        Private ReadOnly _hideNavigatorCheck As New WpfCheckBox()
        Private ReadOnly _hideWindowButton As New WpfButton()
        Private ReadOnly _navigatorBodyHost As New StackPanel()
        Private ReadOnly _goButton As New WpfButton()
        Private ReadOnly _versionTabsHost As New WrapPanel()
        Private ReadOnly _versionTabsHint As New TextBlock()
        Private ReadOnly _navigateAction As Action(Of Integer, DocumentViewOption)
        Private ReadOnly _documentActivateAction As Action(Of Integer, String)
        Private ReadOnly _focusProcessAction As Action(Of Integer)
        Private _navigatorSnapshot As DocumentViewNavigatorSnapshot
        Private _sessions As List(Of DocumentVisualAidSessionState) = New List(Of DocumentVisualAidSessionState)()
        Private _selectedSessionProcessId As Integer
        Private _manualSessionProcessId As Integer
        Private _foregroundSessionProcessId As Integer
        Private _manualNavigatorSessionProcessId As Integer
        Private _manualNavigatorDocumentKey As String = String.Empty
        Private _manualNavigatorCategoryKey As String = String.Empty
        Private _manualNavigatorViewId As Integer = -1
        Private _isUpdatingNavigatorControls As Boolean
        Private _forceClosing As Boolean

        Public Sub New(navigateAction As Action(Of Integer, DocumentViewOption),
                       documentActivateAction As Action(Of Integer, String),
                       focusProcessAction As Action(Of Integer))
            _navigateAction = navigateAction
            _documentActivateAction = documentActivateAction
            _focusProcessAction = focusProcessAction

            Title = "KKY Document Colors"
            Width = 324
            MinWidth = 286
            MaxWidth = 404
            SizeToContent = SizeToContent.Height
            ResizeMode = ResizeMode.NoResize
            WindowStyle = WindowStyle.ToolWindow
            ShowInTaskbar = False
            ShowActivated = False
            Topmost = False
            Background = New SolidColorBrush(WpfColor.FromRgb(248, 250, 254))

            Dim chromeBorder As New Border With {
                .BorderBrush = New SolidColorBrush(WpfColor.FromRgb(206, 214, 227)),
                .BorderThickness = New Thickness(1),
                .Background = Background,
                .Padding = New Thickness(12, 12, 12, 10)
            }

            Dim root As New StackPanel()

            Dim titleText As New TextBlock With {
                .Text = "Document Colors",
                .FontWeight = FontWeights.SemiBold,
                .FontSize = 14,
                .Foreground = New SolidColorBrush(WpfColor.FromRgb(32, 41, 55))
            }

            _summaryText.Text = "Checking open documents"
            _summaryText.Margin = New Thickness(0, 4, 0, 10)
            _summaryText.FontSize = 11
            _summaryText.Foreground = New SolidColorBrush(WpfColor.FromRgb(102, 114, 128))

            _itemsHost.Orientation = Orientation.Vertical

            _emptyText.Text = "No open documents were detected."
            _emptyText.FontSize = 11
            _emptyText.Foreground = New SolidColorBrush(WpfColor.FromRgb(121, 132, 148))
            _emptyText.TextWrapping = TextWrapping.Wrap
            _emptyText.Visibility = System.Windows.Visibility.Collapsed

            root.Children.Add(titleText)
            root.Children.Add(CreateVersionTabsCard())
            root.Children.Add(CreateDocumentSectionCard())
            root.Children.Add(CreateNavigatorCard())

            chromeBorder.Child = root
            Content = chromeBorder
        End Sub

        Friend Sub UpdateContents(aggregate As DocumentVisualAidAggregateState)
            UpdateSessions(aggregate)
            RefreshSelectedSessionContents()
        End Sub

        Friend Sub RequestClose()
            _forceClosing = True
            Close()
        End Sub

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            If Not _forceClosing Then
                e.Cancel = True
                Hide()
                Return
            End If

            MyBase.OnClosing(e)
        End Sub

        Private Function CreateVersionTabsCard() As UIElement
            _versionTabsHint.Text = "Connected Revit versions"
            _versionTabsHint.Margin = New Thickness(0, 0, 0, 8)
            _versionTabsHint.FontSize = 11
            _versionTabsHint.Foreground = New SolidColorBrush(WpfColor.FromRgb(102, 114, 128))

            _versionTabsHost.Orientation = Orientation.Horizontal
            _versionTabsHost.ItemHeight = Double.NaN
            _versionTabsHost.ItemWidth = Double.NaN

            Dim host As New StackPanel()
            host.Children.Add(_versionTabsHint)
            host.Children.Add(_versionTabsHost)

            Return CreateSectionCard(host, marginBottom:=10)
        End Function

        Private Sub UpdateSessions(aggregate As DocumentVisualAidAggregateState)
            Dim previousSelectedSessionProcessId = _selectedSessionProcessId
            _sessions = If(aggregate?.Sessions, New List(Of DocumentVisualAidSessionState)()).ToList()

            _foregroundSessionProcessId =
                DocumentVisualAidSharedSessionCoordinator.TryGetForegroundSessionProcessId(
                    _sessions.Where(Function(item) item IsNot Nothing).
                              Select(Function(item) item.ProcessId))

            Dim hasManualSession =
                _manualSessionProcessId > 0 AndAlso
                _sessions.Any(Function(item) item IsNot Nothing AndAlso item.ProcessId = _manualSessionProcessId)

            If _manualSessionProcessId > 0 AndAlso Not hasManualSession Then
                _manualSessionProcessId = 0
            End If

            If hasManualSession Then
                _selectedSessionProcessId = _manualSessionProcessId
            ElseIf _foregroundSessionProcessId > 0 AndAlso Not IsActive Then
                _selectedSessionProcessId = _foregroundSessionProcessId
            End If

            If _selectedSessionProcessId <= 0 OrElse
               Not _sessions.Any(Function(item) item IsNot Nothing AndAlso item.ProcessId = _selectedSessionProcessId) Then
                Dim preferredSession =
                    If(_foregroundSessionProcessId > 0,
                       _sessions.FirstOrDefault(Function(item) item IsNot Nothing AndAlso item.ProcessId = _foregroundSessionProcessId),
                       Nothing)

                If preferredSession Is Nothing Then
                    preferredSession = _sessions.FirstOrDefault(Function(item) item IsNot Nothing AndAlso item.IsLocalProcess)
                End If

                If preferredSession Is Nothing Then
                    preferredSession = _sessions.FirstOrDefault()
                End If

                _selectedSessionProcessId = If(preferredSession Is Nothing, 0, preferredSession.ProcessId)
            End If

            If previousSelectedSessionProcessId > 0 AndAlso
               previousSelectedSessionProcessId <> _selectedSessionProcessId Then
                ClearManualNavigatorSelection()
            End If

            RenderVersionTabs()
        End Sub

        Private Sub RenderVersionTabs()
            _versionTabsHost.Children.Clear()

            If _sessions Is Nothing OrElse _sessions.Count = 0 Then
                _versionTabsHint.Text = "Waiting for running Revit versions"
                Return
            End If

            Dim selectedSession = GetSelectedSession()
            If selectedSession IsNot Nothing Then
                Dim versionLabel = If(String.IsNullOrWhiteSpace(selectedSession.VersionLabel),
                                      If(String.IsNullOrWhiteSpace(selectedSession.TabLabel), "Revit", selectedSession.TabLabel),
                                      selectedSession.VersionLabel)
                If _manualSessionProcessId > 0 Then
                    _versionTabsHint.Text = $"Selected Revit version: {versionLabel}"
                ElseIf _foregroundSessionProcessId > 0 Then
                    _versionTabsHint.Text = $"Active Revit version: {versionLabel}"
                Else
                    _versionTabsHint.Text = $"Selected Revit version: {versionLabel}"
                End If
            Else
                _versionTabsHint.Text = If(_sessions.Count = 1, "Connected Revit version", "Connected Revit versions")
            End If

            For Each session In _sessions
                If session Is Nothing Then Continue For

                Dim capturedSession = session
                Dim isForeground = _foregroundSessionProcessId > 0 AndAlso session.ProcessId = _foregroundSessionProcessId
                Dim isActive = session.ProcessId = _selectedSessionProcessId
                Dim tabButton As New WpfButton With {
                    .Content = If(String.IsNullOrWhiteSpace(session.TabLabel), session.VersionLabel, session.TabLabel),
                    .ToolTip = If(String.IsNullOrWhiteSpace(session.TabToolTip), session.VersionLabel, session.TabToolTip),
                    .Margin = New Thickness(0, 0, 6, 6),
                    .MinWidth = 64,
                    .Padding = New Thickness(12, 6, 12, 6),
                    .FontSize = 10.5,
                    .FontWeight = If(isActive, FontWeights.SemiBold, FontWeights.Medium),
                    .Foreground = If(isActive,
                                     Brushes.White,
                                     New SolidColorBrush(WpfColor.FromRgb(72, 83, 99))),
                    .Background = If(isActive,
                                     New SolidColorBrush(WpfColor.FromRgb(64, 103, 196)),
                                     New SolidColorBrush(WpfColor.FromRgb(255, 255, 255))),
                    .BorderBrush = If(isForeground,
                                      New SolidColorBrush(WpfColor.FromRgb(35, 84, 188)),
                                      If(isActive,
                                         New SolidColorBrush(WpfColor.FromRgb(64, 103, 196)),
                                         New SolidColorBrush(WpfColor.FromRgb(208, 216, 228)))),
                    .BorderThickness = If(isForeground, New Thickness(2), New Thickness(1)),
                    .Cursor = System.Windows.Input.Cursors.Hand
                }
                AddHandler tabButton.Click,
                    Sub()
                        _manualSessionProcessId = capturedSession.ProcessId
                        _selectedSessionProcessId = capturedSession.ProcessId
                        ClearManualNavigatorSelection()
                        RefreshSelectedSessionContents()
                        If _focusProcessAction IsNot Nothing Then
                            _focusProcessAction.Invoke(capturedSession.ProcessId)
                        End If
                    End Sub

                _versionTabsHost.Children.Add(tabButton)
            Next
        End Sub

        Private Function GetSelectedSession() As DocumentVisualAidSessionState
            If _sessions Is Nothing OrElse _sessions.Count = 0 Then Return Nothing

            Dim selectedSession = _sessions.FirstOrDefault(
                Function(item) item IsNot Nothing AndAlso item.ProcessId = _selectedSessionProcessId)
            If selectedSession IsNot Nothing Then Return selectedSession

            Return _sessions.FirstOrDefault()
        End Function

        Private Sub RefreshSelectedSessionContents()
            Dim selectedSession = GetSelectedSession()
            If selectedSession Is Nothing Then
                UpdateEntries(Nothing, Nothing)
                UpdateNavigator(Nothing)
                Return
            End If

            UpdateEntries(selectedSession.Entries, selectedSession.VersionLabel)
            UpdateNavigator(selectedSession.Navigator)
        End Sub

        Private Sub UpdateEntries(entries As IList(Of DocumentColorEntry), versionLabel As String)
            _itemsHost.Children.Clear()

            Dim entryCount As Integer = If(entries Is Nothing, 0, entries.Count)
            If String.IsNullOrWhiteSpace(versionLabel) Then
                _summaryText.Text = $"{entryCount} document(s)"
            Else
                _summaryText.Text = $"{versionLabel} - {entryCount} document(s)"
            End If

            If entries Is Nothing OrElse entries.Count = 0 Then
                _emptyText.Visibility = System.Windows.Visibility.Visible
                Return
            End If

            _emptyText.Visibility = System.Windows.Visibility.Collapsed

            For Each entry In entries
                If entry Is Nothing Then Continue For
                _itemsHost.Children.Add(CreateEntryRow(entry))
            Next
        End Sub

        Private Function CreateDocumentSectionCard() As UIElement
            _summaryText.Margin = New Thickness(0, 0, 0, 10)
            _summaryText.FontSize = 11
            _summaryText.Foreground = New SolidColorBrush(WpfColor.FromRgb(102, 114, 128))

            Dim host As New StackPanel()
            host.Children.Add(_summaryText)
            host.Children.Add(_itemsHost)
            host.Children.Add(_emptyText)

            Return CreateSectionCard(host, marginBottom:=10)
        End Function

        Private Sub UpdateNavigator(navigator As DocumentViewNavigatorSnapshot)
            _navigatorSnapshot = navigator
            _isUpdatingNavigatorControls = True

            Try
                If navigator Is Nothing OrElse navigator.Documents Is Nothing OrElse navigator.Documents.Count = 0 Then
                    _navigatorSummaryText.Text = "Open a document to browse views."
                    _navigatorEmptyText.Text = "Pick a document, then a view type, then jump to the view you want."
                    _navigatorEmptyText.Visibility = System.Windows.Visibility.Visible
                    _documentCombo.ItemsSource = Nothing
                    _documentCombo.IsEnabled = False
                    _viewTypeCombo.ItemsSource = Nothing
                    _viewTypeCombo.IsEnabled = False
                    _viewCombo.ItemsSource = Nothing
                    _viewCombo.IsEnabled = False
                    _goButton.IsEnabled = False
                    Return
                End If

                _navigatorSummaryText.Text = "Active document follows automatically."
                _navigatorEmptyText.Visibility = System.Windows.Visibility.Collapsed

                _documentCombo.ItemsSource = navigator.Documents
                _documentCombo.IsEnabled = True

                Dim hasManualNavigatorSelection =
                    _manualNavigatorSessionProcessId = _selectedSessionProcessId AndAlso
                    Not String.IsNullOrWhiteSpace(_manualNavigatorDocumentKey)

                Dim selectedDocument As DocumentViewNavigatorDocument = Nothing
                If hasManualNavigatorSelection Then
                    selectedDocument = navigator.Documents.FirstOrDefault(
                        Function(item) String.Equals(item.DocumentKey, _manualNavigatorDocumentKey, StringComparison.OrdinalIgnoreCase))
                    If selectedDocument Is Nothing Then
                        ClearManualNavigatorSelection()
                        hasManualNavigatorSelection = False
                    End If
                End If

                If selectedDocument Is Nothing Then
                    selectedDocument = navigator.Documents.FirstOrDefault(
                        Function(item) String.Equals(item.DocumentKey, navigator.SelectedDocumentKey, StringComparison.OrdinalIgnoreCase))
                End If
                If selectedDocument Is Nothing Then
                    selectedDocument = navigator.Documents.FirstOrDefault()
                End If

                _documentCombo.SelectedItem = selectedDocument

                If selectedDocument IsNot Nothing Then
                    Dim preferredCategoryKey = selectedDocument.DefaultCategoryKey
                    Dim preferredViewId = selectedDocument.ActiveViewId

                    If hasManualNavigatorSelection AndAlso
                       String.Equals(selectedDocument.DocumentKey, _manualNavigatorDocumentKey, StringComparison.OrdinalIgnoreCase) Then
                        If Not String.IsNullOrWhiteSpace(_manualNavigatorCategoryKey) Then
                            preferredCategoryKey = _manualNavigatorCategoryKey
                            preferredViewId = _manualNavigatorViewId
                        ElseIf _manualNavigatorViewId > 0 Then
                            preferredViewId = _manualNavigatorViewId
                        End If
                    End If

                    UpdateCategorySelection(selectedDocument,
                                            preferredCategoryKey:=preferredCategoryKey,
                                            preferredViewId:=preferredViewId)
                Else
                    _viewTypeCombo.ItemsSource = Nothing
                    _viewTypeCombo.IsEnabled = False
                    _viewCombo.ItemsSource = Nothing
                    _viewCombo.IsEnabled = False
                    _goButton.IsEnabled = False
                End If
            Finally
                _isUpdatingNavigatorControls = False
            End Try
        End Sub

        Private Shared Function CreateSectionCard(content As UIElement, Optional marginBottom As Double = 0.0R) As UIElement
            Return New Border With {
                .Background = New SolidColorBrush(WpfColor.FromRgb(255, 255, 255)),
                .BorderBrush = New SolidColorBrush(WpfColor.FromRgb(223, 229, 238)),
                .BorderThickness = New Thickness(1),
                .CornerRadius = New CornerRadius(8),
                .Padding = New Thickness(10),
                .Margin = New Thickness(0, 0, 0, marginBottom),
                .Child = content
            }
        End Function

        Private Function CreateNavigatorCard() As UIElement
            Dim titleText As New TextBlock With {
                .Text = "View Navigator",
                .FontWeight = FontWeights.SemiBold,
                .FontSize = 13,
                .Foreground = New SolidColorBrush(WpfColor.FromRgb(32, 41, 55))
            }

            _navigatorSummaryText.Text = "Active document follows automatically."
            _navigatorSummaryText.Margin = New Thickness(0, 4, 0, 8)
            _navigatorSummaryText.FontSize = 11
            _navigatorSummaryText.Foreground = New SolidColorBrush(WpfColor.FromRgb(102, 114, 128))
            _navigatorSummaryText.TextTrimming = TextTrimming.CharacterEllipsis

            Dim documentLabel = CreateNavigatorLabel("Document")
            Dim viewTypeLabel = CreateNavigatorLabel("Type")
            Dim viewLabel = CreateNavigatorLabel("View")

            _documentCombo.Margin = New Thickness(0, 0, 0, 6)
            _documentCombo.MinWidth = 236
            AddHandler _documentCombo.SelectionChanged, AddressOf OnSelectedDocumentChanged
            StyleComboBox(_documentCombo)

            _viewTypeCombo.Margin = New Thickness(0, 0, 0, 6)
            _viewTypeCombo.MinWidth = 236
            AddHandler _viewTypeCombo.SelectionChanged, AddressOf OnViewTypeChanged
            StyleComboBox(_viewTypeCombo)

            _viewCombo.Margin = New Thickness(0, 0, 0, 8)
            _viewCombo.MinWidth = 236
            AddHandler _viewCombo.SelectionChanged, AddressOf OnSelectedViewChanged
            StyleComboBox(_viewCombo)

            _alwaysOnTopCheck.Content = "Always on top"
            _alwaysOnTopCheck.FontSize = 10.5
            _alwaysOnTopCheck.FontWeight = FontWeights.Medium
            _alwaysOnTopCheck.Foreground = New SolidColorBrush(WpfColor.FromRgb(82, 94, 112))
            _alwaysOnTopCheck.VerticalAlignment = VerticalAlignment.Center
            _alwaysOnTopCheck.IsChecked = Topmost
            AddHandler _alwaysOnTopCheck.Checked, AddressOf OnAlwaysOnTopChanged
            AddHandler _alwaysOnTopCheck.Unchecked, AddressOf OnAlwaysOnTopChanged

            _hideWindowButton.Content = "Hide window"
            _hideWindowButton.MinWidth = 92
            _hideWindowButton.Margin = New Thickness(8, 0, 0, 0)
            AddHandler _hideWindowButton.Click, AddressOf OnHideWindowClicked
            StyleSecondaryButton(_hideWindowButton)

            _hideNavigatorCheck.Content = "Hide view navigator"
            _hideNavigatorCheck.FontSize = 10.5
            _hideNavigatorCheck.FontWeight = FontWeights.Medium
            _hideNavigatorCheck.Foreground = New SolidColorBrush(WpfColor.FromRgb(82, 94, 112))
            _hideNavigatorCheck.Margin = New Thickness(0, 6, 0, 0)
            AddHandler _hideNavigatorCheck.Checked, AddressOf OnHideNavigatorChanged
            AddHandler _hideNavigatorCheck.Unchecked, AddressOf OnHideNavigatorChanged

            _goButton.Content = "Go"
            _goButton.Width = 72
            _goButton.IsEnabled = False
            _goButton.Margin = New Thickness(0, 8, 0, 0)
            AddHandler _goButton.Click, AddressOf OnGoClicked
            StylePrimaryButton(_goButton)

            _navigatorEmptyText.Text = "Pick a document, then a view type, then jump to the view you want."
            _navigatorEmptyText.FontSize = 10.5
            _navigatorEmptyText.Foreground = New SolidColorBrush(WpfColor.FromRgb(121, 132, 148))
            _navigatorEmptyText.TextWrapping = TextWrapping.Wrap
            _navigatorEmptyText.Visibility = System.Windows.Visibility.Collapsed

            Dim persistentActionRow As New StackPanel With {
                .Orientation = Orientation.Horizontal,
                .Margin = New Thickness(0, 6, 0, 0)
            }
            persistentActionRow.Children.Add(_alwaysOnTopCheck)
            persistentActionRow.Children.Add(_hideWindowButton)

            Dim goRow As New DockPanel With {
                .LastChildFill = False
            }
            DockPanel.SetDock(_goButton, Dock.Right)
            goRow.Children.Add(_goButton)

            _navigatorBodyHost.Children.Add(_navigatorSummaryText)
            _navigatorBodyHost.Children.Add(documentLabel)
            _navigatorBodyHost.Children.Add(_documentCombo)
            _navigatorBodyHost.Children.Add(viewTypeLabel)
            _navigatorBodyHost.Children.Add(_viewTypeCombo)
            _navigatorBodyHost.Children.Add(viewLabel)
            _navigatorBodyHost.Children.Add(_viewCombo)
            _navigatorBodyHost.Children.Add(goRow)
            _navigatorBodyHost.Children.Add(_navigatorEmptyText)
            UpdateNavigatorBodyVisibility()

            Dim host As New StackPanel()
            host.Children.Add(titleText)
            host.Children.Add(persistentActionRow)
            host.Children.Add(_hideNavigatorCheck)
            host.Children.Add(_navigatorBodyHost)

            Return CreateSectionCard(host)
        End Function

        Private Shared Function CreateNavigatorLabel(text As String) As UIElement
            Return New TextBlock With {
                .Text = text,
                .Margin = New Thickness(0, 0, 0, 4),
                .FontSize = 10.5,
                .FontWeight = FontWeights.Medium,
                .Foreground = New SolidColorBrush(WpfColor.FromRgb(102, 114, 128))
            }
        End Function

        Private Shared Sub StyleComboBox(combo As WpfComboBox)
            If combo Is Nothing Then Return

            combo.Height = 30
            combo.Background = New SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            combo.BorderBrush = New SolidColorBrush(WpfColor.FromRgb(208, 216, 228))
            combo.BorderThickness = New Thickness(1)
            combo.Foreground = New SolidColorBrush(WpfColor.FromRgb(34, 43, 58))
            combo.FontSize = 11.5
            combo.VerticalContentAlignment = VerticalAlignment.Center
        End Sub

        Private Shared Sub StylePrimaryButton(button As WpfButton)
            If button Is Nothing Then Return

            button.Height = 30
            button.Padding = New Thickness(12, 0, 12, 0)
            button.HorizontalAlignment = HorizontalAlignment.Right
            button.FontSize = 11.5
            button.FontWeight = FontWeights.SemiBold
            button.Foreground = Brushes.White
            button.Background = New SolidColorBrush(WpfColor.FromRgb(64, 103, 196))
            button.BorderBrush = New SolidColorBrush(WpfColor.FromRgb(64, 103, 196))
            button.BorderThickness = New Thickness(1)
        End Sub

        Private Shared Sub StyleSecondaryButton(button As WpfButton)
            If button Is Nothing Then Return

            button.Height = 26
            button.Padding = New Thickness(10, 0, 10, 0)
            button.FontSize = 10.5
            button.FontWeight = FontWeights.Medium
            button.Foreground = New SolidColorBrush(WpfColor.FromRgb(72, 83, 99))
            button.Background = New SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            button.BorderBrush = New SolidColorBrush(WpfColor.FromRgb(208, 216, 228))
            button.BorderThickness = New Thickness(1)
            button.VerticalAlignment = VerticalAlignment.Center
        End Sub

        Private Sub OnAlwaysOnTopChanged(sender As Object, e As RoutedEventArgs)
            Topmost = _alwaysOnTopCheck.IsChecked.GetValueOrDefault(False)
        End Sub

        Private Sub OnHideWindowClicked(sender As Object, e As RoutedEventArgs)
            DocumentVisualAidService.HideLegendWindowPreservingTabs()
        End Sub

        Private Sub OnHideNavigatorChanged(sender As Object, e As RoutedEventArgs)
            UpdateNavigatorBodyVisibility()
        End Sub

        Private Sub UpdateNavigatorBodyVisibility()
            _navigatorBodyHost.Visibility =
                If(_hideNavigatorCheck.IsChecked.GetValueOrDefault(False),
                   System.Windows.Visibility.Collapsed,
                   System.Windows.Visibility.Visible)
        End Sub

        Private Sub OnSelectedDocumentChanged(sender As Object, e As SelectionChangedEventArgs)
            If _isUpdatingNavigatorControls Then Return

            Dim selectedDocument = TryCast(_documentCombo.SelectedItem, DocumentViewNavigatorDocument)
            If selectedDocument Is Nothing Then
                ClearManualNavigatorSelection()
                _viewTypeCombo.ItemsSource = Nothing
                _viewTypeCombo.IsEnabled = False
                _viewCombo.ItemsSource = Nothing
                _viewCombo.IsEnabled = False
                _goButton.IsEnabled = False
                Return
            End If

            RememberManualNavigatorSelection(selectedDocument, Nothing, Nothing)

            Dim previousUpdating = _isUpdatingNavigatorControls
            _isUpdatingNavigatorControls = True
            Try
                UpdateCategorySelection(selectedDocument,
                                        preferredCategoryKey:=selectedDocument.DefaultCategoryKey,
                                        preferredViewId:=selectedDocument.ActiveViewId)
            Finally
                _isUpdatingNavigatorControls = previousUpdating
            End Try
        End Sub

        Private Sub OnViewTypeChanged(sender As Object, e As SelectionChangedEventArgs)
            If _isUpdatingNavigatorControls Then Return

            Dim selectedDocument = TryCast(_documentCombo.SelectedItem, DocumentViewNavigatorDocument)
            Dim selectedCategory = TryCast(_viewTypeCombo.SelectedItem, DocumentViewCategory)

            RememberManualNavigatorSelection(selectedDocument, selectedCategory, Nothing)

            Dim previousUpdating = _isUpdatingNavigatorControls
            _isUpdatingNavigatorControls = True
            Try
                UpdateViewSelection(-1)
            Finally
                _isUpdatingNavigatorControls = previousUpdating
            End Try
        End Sub

        Private Sub OnSelectedViewChanged(sender As Object, e As SelectionChangedEventArgs)
            If _isUpdatingNavigatorControls Then Return

            Dim selectedDocument = TryCast(_documentCombo.SelectedItem, DocumentViewNavigatorDocument)
            Dim selectedCategory = TryCast(_viewTypeCombo.SelectedItem, DocumentViewCategory)
            Dim selectedView = TryCast(_viewCombo.SelectedItem, DocumentViewOption)

            RememberManualNavigatorSelection(selectedDocument, selectedCategory, selectedView)
            _goButton.IsEnabled = TypeOf _viewCombo.SelectedItem Is DocumentViewOption
        End Sub

        Private Sub UpdateCategorySelection(selectedDocument As DocumentViewNavigatorDocument,
                                            preferredCategoryKey As String,
                                            preferredViewId As Integer)
            If selectedDocument Is Nothing OrElse selectedDocument.Categories Is Nothing OrElse selectedDocument.Categories.Count = 0 Then
                _viewTypeCombo.ItemsSource = Nothing
                _viewTypeCombo.IsEnabled = False
                _viewCombo.ItemsSource = Nothing
                _viewCombo.IsEnabled = False
                _goButton.IsEnabled = False
                Return
            End If

            _viewTypeCombo.ItemsSource = selectedDocument.Categories
            _viewTypeCombo.IsEnabled = True

            Dim defaultCategory = selectedDocument.Categories.FirstOrDefault(
                Function(category) String.Equals(category.Key, preferredCategoryKey, StringComparison.OrdinalIgnoreCase))
            If defaultCategory Is Nothing Then
                defaultCategory = selectedDocument.Categories.FirstOrDefault(
                    Function(category) String.Equals(category.Key, selectedDocument.DefaultCategoryKey, StringComparison.OrdinalIgnoreCase))
            End If
            If defaultCategory Is Nothing Then
                defaultCategory = selectedDocument.Categories.FirstOrDefault()
            End If

            _viewTypeCombo.SelectedItem = defaultCategory
            UpdateViewSelection(preferredViewId)
        End Sub

        Private Sub UpdateViewSelection(preferredViewId As Integer)
            Dim category = TryCast(_viewTypeCombo.SelectedItem, DocumentViewCategory)
            If category Is Nothing OrElse category.Views Is Nothing OrElse category.Views.Count = 0 Then
                _viewCombo.ItemsSource = Nothing
                _viewCombo.IsEnabled = False
                _goButton.IsEnabled = False
                Return
            End If

            _viewCombo.ItemsSource = category.Views
            _viewCombo.IsEnabled = True

            Dim preferredItem = category.Views.FirstOrDefault(Function(viewOption) viewOption.ViewId = preferredViewId)
            If preferredItem Is Nothing Then
                preferredItem = category.Views.FirstOrDefault()
            End If

            _viewCombo.SelectedItem = preferredItem
            _goButton.IsEnabled = preferredItem IsNot Nothing
        End Sub

        Private Sub OnGoClicked(sender As Object, e As RoutedEventArgs)
            Dim selectedView = TryCast(_viewCombo.SelectedItem, DocumentViewOption)
            If selectedView Is Nothing OrElse _navigateAction Is Nothing Then Return

            RememberManualNavigatorSelection(TryCast(_documentCombo.SelectedItem, DocumentViewNavigatorDocument),
                                             TryCast(_viewTypeCombo.SelectedItem, DocumentViewCategory),
                                             selectedView)
            _navigateAction.Invoke(_selectedSessionProcessId, selectedView)
        End Sub

        Private Sub RememberManualNavigatorSelection(selectedDocument As DocumentViewNavigatorDocument,
                                                     selectedCategory As DocumentViewCategory,
                                                     selectedView As DocumentViewOption)
            _manualNavigatorSessionProcessId = _selectedSessionProcessId
            _manualNavigatorDocumentKey = If(selectedDocument?.DocumentKey, String.Empty)
            _manualNavigatorCategoryKey = If(selectedCategory?.Key, String.Empty)
            _manualNavigatorViewId = If(selectedView Is Nothing, -1, selectedView.ViewId)
        End Sub

        Private Sub ClearManualNavigatorSelection()
            _manualNavigatorSessionProcessId = 0
            _manualNavigatorDocumentKey = String.Empty
            _manualNavigatorCategoryKey = String.Empty
            _manualNavigatorViewId = -1
        End Sub

        Private Function CreateEntryRow(entry As DocumentColorEntry) As UIElement
            Dim chip As New Border With {
                .Width = 14,
                .Height = 14,
                .CornerRadius = New CornerRadius(3),
                .Background = entry.ChipBrush,
                .BorderBrush = entry.BorderBrush,
                .BorderThickness = New Thickness(1),
                .Margin = New Thickness(0, 2, 10, 0),
                .VerticalAlignment = VerticalAlignment.Top
            }

            Dim titleText As New TextBlock With {
                .Text = entry.DisplayName,
                .FontSize = 12.5,
                .FontWeight = If(entry.IsActive, FontWeights.SemiBold, FontWeights.Normal),
                .Foreground = New SolidColorBrush(WpfColor.FromRgb(34, 43, 58)),
                .TextTrimming = TextTrimming.CharacterEllipsis
            }

            Dim detailText As New TextBlock With {
                .Text = If(entry.IsActive, "Active document", "Open document"),
                .FontSize = 10.5,
                .Foreground = New SolidColorBrush(WpfColor.FromRgb(108, 119, 136)),
                .Margin = New Thickness(0, 2, 0, 0)
            }

            Dim textStack As New StackPanel()
            textStack.Children.Add(titleText)
            textStack.Children.Add(detailText)

            Dim row As New DockPanel With {
                .LastChildFill = True,
                .ToolTip = If(String.IsNullOrWhiteSpace(entry.FullPath), entry.DisplayName, entry.FullPath)
            }

            DockPanel.SetDock(chip, Dock.Left)
            row.Children.Add(chip)
            row.Children.Add(textStack)

            Dim container As New Border With {
                .Background = If(entry.IsActive,
                                 New SolidColorBrush(WpfColor.FromRgb(244, 248, 255)),
                                 New SolidColorBrush(WpfColor.FromRgb(255, 255, 255))),
                .BorderBrush = If(entry.IsActive,
                                  entry.BorderBrush,
                                  New SolidColorBrush(WpfColor.FromRgb(226, 231, 239))),
                .BorderThickness = New Thickness(1),
                .CornerRadius = New CornerRadius(6),
                .Padding = New Thickness(10, 8, 10, 8),
                .Margin = New Thickness(0, 0, 0, 8),
                .ToolTip = row.ToolTip,
                .Cursor = System.Windows.Input.Cursors.Hand,
                .Focusable = True,
                .Child = row
            }
            AddHandler container.MouseLeftButtonUp,
                Sub(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
                    If _documentActivateAction IsNot Nothing Then
                        _documentActivateAction.Invoke(_selectedSessionProcessId, entry.Key)
                    End If
                End Sub
            AddHandler container.PreviewKeyDown,
                Sub(sender As Object, e As System.Windows.Input.KeyEventArgs)
                    If e.Key = System.Windows.Input.Key.Enter OrElse
                       e.Key = System.Windows.Input.Key.Space Then
                        e.Handled = True
                        If _documentActivateAction IsNot Nothing Then
                            _documentActivateAction.Invoke(_selectedSessionProcessId, entry.Key)
                        End If
                    End If
                End Sub

            Return container
        End Function

    End Class

    Friend NotInheritable Class DocumentColorEntry

        Friend Property Key As String
        Friend Property DisplayName As String
        Friend Property FullPath As String
        Friend Property ColorSlot As Integer
        Friend Property ChipBrush As SolidColorBrush
        Friend Property TabFillBrush As SolidColorBrush
        Friend Property ActiveTabFillBrush As SolidColorBrush
        Friend Property BorderBrush As SolidColorBrush
        Friend Property AccentBrush As SolidColorBrush
        Friend Property IsActive As Boolean
        Friend Property MatchTokens As IList(Of String)

    End Class

    Friend NotInheritable Class DocumentViewNavigatorSnapshot

        Friend Property SelectedDocumentKey As String
        Friend Property Documents As IList(Of DocumentViewNavigatorDocument)

        Public Overrides Function ToString() As String
            If Documents Is Nothing OrElse Documents.Count = 0 Then Return String.Empty

            Dim selectedDocument = Documents.FirstOrDefault(
                Function(item) String.Equals(item.DocumentKey, SelectedDocumentKey, StringComparison.OrdinalIgnoreCase))
            If selectedDocument Is Nothing Then
                selectedDocument = Documents(0)
            End If

            Return If(selectedDocument.DocumentName, String.Empty)
        End Function

    End Class

    Friend NotInheritable Class DocumentViewNavigatorDocument

        Friend Property DocumentKey As String
        Friend Property DocumentName As String
        Friend Property ActiveViewId As Integer
        Friend Property DefaultCategoryKey As String
        Friend Property Categories As IList(Of DocumentViewCategory)

        Public Overrides Function ToString() As String
            Return If(DocumentName, String.Empty)
        End Function

    End Class

    Friend NotInheritable Class DocumentViewCategory

        Friend Property Key As String
        Friend Property DisplayName As String
        Friend Property Views As IList(Of DocumentViewOption)

        Public Overrides Function ToString() As String
            Return If(DisplayName, String.Empty)
        End Function

    End Class

    Friend NotInheritable Class DocumentViewOption

        Friend Property DocumentKey As String
        Friend Property ViewId As Integer
        Friend Property CategoryKey As String
        Friend Property DisplayName As String

        Public Overrides Function ToString() As String
            Return If(DisplayName, String.Empty)
        End Function

    End Class

End Namespace
