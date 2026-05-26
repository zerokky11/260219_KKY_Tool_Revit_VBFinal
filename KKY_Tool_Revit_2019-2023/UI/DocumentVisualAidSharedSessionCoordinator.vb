Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Web.Script.Serialization
Imports Autodesk.Revit.DB
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application
Imports WpfColor = System.Windows.Media.Color
Imports WpfSolidColorBrush = System.Windows.Media.SolidColorBrush

Namespace UI

    Friend NotInheritable Class DocumentVisualAidSharedSessionCoordinator

        Private Shared ReadOnly RootDirectoryPath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "KKY_Tool_Revit",
                         "document-visual-aid-shared")
        Private Shared ReadOnly SessionsDirectoryPath As String = Path.Combine(RootDirectoryPath, "sessions")
        Private Shared ReadOnly CommandsDirectoryPath As String = Path.Combine(RootDirectoryPath, "commands")
        Private Shared ReadOnly OwnerFilePath As String = Path.Combine(RootDirectoryPath, "owner.json")
        Private Shared ReadOnly Serializer As New JavaScriptSerializer() With {
            .MaxJsonLength = Integer.MaxValue
        }
        Private Shared ReadOnly LocalProcessId As Integer = Process.GetCurrentProcess().Id
        Private Shared ReadOnly LocalStartedUtcTicks As Long = DateTime.UtcNow.Ticks
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

        Friend Shared Function BuildLocalState(app As RevitApp,
                                               entries As IReadOnlyList(Of DocumentColorEntry),
                                               navigator As DocumentViewNavigatorSnapshot) As DocumentVisualAidSessionState
            Return New DocumentVisualAidSessionState With {
                .ProcessId = LocalProcessId,
                .VersionLabel = ResolveVersionLabel(app),
                .StartedUtcTicks = LocalStartedUtcTicks,
                .UpdatedUtcTicks = DateTime.UtcNow.Ticks,
                .Entries = If(entries, New List(Of DocumentColorEntry)()).ToList(),
                .Navigator = navigator,
                .IsLocalProcess = True
            }
        End Function

        Friend Shared Sub PublishLocalState(state As DocumentVisualAidSessionState)
            If state Is Nothing Then
                ClearLocalState()
                Return
            End If

            EnsureDirectories()

            Dim model As New DocumentVisualAidSessionFileModel With {
                .ProcessId = state.ProcessId,
                .VersionLabel = state.VersionLabel,
                .StartedUtcTicks = state.StartedUtcTicks,
                .UpdatedUtcTicks = state.UpdatedUtcTicks,
                .Documents = ConvertEntriesToFileModel(state.Entries),
                .Navigator = ConvertNavigatorToFileModel(state.Navigator)
            }

            WriteJson(GetSessionFilePath(state.ProcessId), model)
        End Sub

        Friend Shared Sub ClearLocalState()
            SafeDeleteFile(GetSessionFilePath(LocalProcessId))
            SafeDeleteDirectory(GetCommandDirectoryPath(LocalProcessId))
            ReleaseOwnerIfLocal()
        End Sub

        Friend Shared Function TryBecomeOwner() As Boolean
            EnsureDirectories()

            Dim currentOwner = TryReadJson(Of DocumentVisualAidOwnerFileModel)(OwnerFilePath)
            If currentOwner IsNot Nothing Then
                If currentOwner.ProcessId = LocalProcessId AndAlso currentOwner.StartedUtcTicks = LocalStartedUtcTicks Then
                    Return True
                End If

                If IsOwnerValid(currentOwner) Then
                    Return False
                End If
            End If

            Dim newOwner As New DocumentVisualAidOwnerFileModel With {
                .ProcessId = LocalProcessId,
                .StartedUtcTicks = LocalStartedUtcTicks
            }
            WriteJson(OwnerFilePath, newOwner)

            Dim verifiedOwner = TryReadJson(Of DocumentVisualAidOwnerFileModel)(OwnerFilePath)
            Return verifiedOwner IsNot Nothing AndAlso
                   verifiedOwner.ProcessId = LocalProcessId AndAlso
                   verifiedOwner.StartedUtcTicks = LocalStartedUtcTicks
        End Function

        Friend Shared Sub ReleaseOwnerIfLocal()
            Dim currentOwner = TryReadJson(Of DocumentVisualAidOwnerFileModel)(OwnerFilePath)
            If currentOwner Is Nothing Then Return

            If currentOwner.ProcessId = LocalProcessId AndAlso currentOwner.StartedUtcTicks = LocalStartedUtcTicks Then
                SafeDeleteFile(OwnerFilePath)
            End If
        End Sub

        Friend Shared Sub ClaimOwnerForLocal()
            EnsureDirectories()

            Dim newOwner As New DocumentVisualAidOwnerFileModel With {
                .ProcessId = LocalProcessId,
                .StartedUtcTicks = LocalStartedUtcTicks
            }
            WriteJson(OwnerFilePath, newOwner)
        End Sub

        Friend Shared Function BuildAggregateState(localState As DocumentVisualAidSessionState) As DocumentVisualAidAggregateState
            EnsureDirectories()

            Dim sessionMap As New Dictionary(Of Integer, DocumentVisualAidSessionState)()
            For Each session In LoadPersistedSessions()
                If session Is Nothing Then Continue For
                sessionMap(session.ProcessId) = session
            Next

            If localState IsNot Nothing Then
                sessionMap(localState.ProcessId) = localState
            End If

            Dim sessions = sessionMap.Values.
                OrderBy(Function(item) item.StartedUtcTicks).
                ThenBy(Function(item) item.ProcessId).
                ToList()

            ApplyGlobalColorSlots(sessions)
            ApplyTabLabels(sessions)

            Return New DocumentVisualAidAggregateState With {
                .Sessions = sessions
            }
        End Function

        Friend Shared Function TryGetForegroundSessionProcessId(processIds As IEnumerable(Of Integer)) As Integer
            If processIds Is Nothing Then Return 0

            Dim candidates As New HashSet(Of Integer)()
            For Each processId In processIds
                If processId > 0 Then
                    candidates.Add(processId)
                End If
            Next

            If candidates.Count = 0 Then Return 0

            Dim foregroundWindow = GetForegroundWindow()
            If foregroundWindow = IntPtr.Zero Then Return 0

            Dim ownerPid As UInteger
            GetWindowThreadProcessId(foregroundWindow, ownerPid)
            Dim resolvedProcessId As Integer = CInt(ownerPid)
            If resolvedProcessId <= 0 Then Return 0

            If candidates.Contains(resolvedProcessId) Then
                Return resolvedProcessId
            End If

            Return 0
        End Function

        Friend Shared Sub EnqueueDocumentActivation(targetProcessId As Integer, documentKey As String)
            If targetProcessId <= 0 OrElse String.IsNullOrWhiteSpace(documentKey) Then Return

            Dim command As New DocumentVisualAidCommandFileModel With {
                .CommandName = "activate-document",
                .DocumentKey = documentKey
            }
            WriteCommand(targetProcessId, command)
        End Sub

        Friend Shared Sub EnqueueViewNavigation(targetProcessId As Integer, target As DocumentViewOption)
            If targetProcessId <= 0 OrElse target Is Nothing Then Return

            Dim command As New DocumentVisualAidCommandFileModel With {
                .CommandName = "navigate-view",
                .DocumentKey = target.DocumentKey,
                .ViewId = target.ViewId
            }
            WriteCommand(targetProcessId, command)
        End Sub

        Friend Shared Sub ProcessLocalCommands(documentActivationAction As Action(Of String),
                                               navigateAction As Action(Of DocumentViewOption))
            EnsureDirectories()

            Dim commandDirectory = GetCommandDirectoryPath(LocalProcessId)
            If Not Directory.Exists(commandDirectory) Then Return

            Dim files As String() = Array.Empty(Of String)()
            Try
                files = Directory.GetFiles(commandDirectory, "*.json")
            Catch
                files = Array.Empty(Of String)()
            End Try

            For Each filePath In files.OrderBy(Function(path) path, StringComparer.OrdinalIgnoreCase)
                Dim command = TryReadJson(Of DocumentVisualAidCommandFileModel)(filePath)
                SafeDeleteFile(filePath)
                If command Is Nothing Then Continue For

                Select Case command.CommandName
                    Case "activate-document"
                        If documentActivationAction IsNot Nothing AndAlso
                           Not String.IsNullOrWhiteSpace(command.DocumentKey) Then
                            documentActivationAction.Invoke(command.DocumentKey)
                        End If

                    Case "navigate-view"
                        If navigateAction Is Nothing OrElse
                           String.IsNullOrWhiteSpace(command.DocumentKey) OrElse
                           command.ViewId <= 0 Then
                            Continue For
                        End If

                        navigateAction.Invoke(New DocumentViewOption With {
                            .DocumentKey = command.DocumentKey,
                            .ViewId = command.ViewId
                        })
                End Select
            Next
        End Sub

        Friend Shared Sub FocusProcessWindow(processId As Integer)
            If processId <= 0 Then Return

            Dim hWnd = ResolveMainWindowHandle(processId)
            If hWnd = IntPtr.Zero Then Return

            If IsIconic(hWnd) Then
                ShowWindow(hWnd, SW_RESTORE)
            End If
            BringWindowToTop(hWnd)
            SetForegroundWindow(hWnd)
        End Sub

        Private Shared Sub ApplyTabLabels(sessions As IList(Of DocumentVisualAidSessionState))
            If sessions Is Nothing Then Return

            Dim versionCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each session In sessions
                If session Is Nothing Then Continue For

                Dim versionKey = If(session.VersionLabel, String.Empty).Trim()
                If versionKey.Length = 0 Then versionKey = "Revit"

                Dim nextIndex As Integer = 1
                If versionCounts.ContainsKey(versionKey) Then
                    nextIndex = versionCounts(versionKey) + 1
                End If
                versionCounts(versionKey) = nextIndex

                session.TabLabel = If(nextIndex = 1, versionKey, $"{versionKey} #{nextIndex}")
                session.TabToolTip = versionKey
            Next
        End Sub

        Private Shared Sub ApplyGlobalColorSlots(sessions As IList(Of DocumentVisualAidSessionState))
            If sessions Is Nothing Then Return

            Dim documentKeys As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each session In sessions
                If session Is Nothing OrElse session.Entries Is Nothing Then Continue For

                For Each entry In session.Entries
                    If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.Key) Then Continue For

                    Dim normalizedKey = entry.Key.Trim()
                    If Not documentKeys.ContainsKey(normalizedKey) Then
                        documentKeys(normalizedKey) = normalizedKey
                    End If
                Next
            Next

            If documentKeys.Count = 0 Then Return

            Dim orderedKeys = documentKeys.Values.
                OrderBy(Function(key) key, StringComparer.OrdinalIgnoreCase).
                ToList()
            Dim slotByKey As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For index = 0 To orderedKeys.Count - 1
                slotByKey(orderedKeys(index)) = index
            Next

            For Each session In sessions
                If session Is Nothing OrElse session.Entries Is Nothing Then Continue For

                For Each entry In session.Entries
                    If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.Key) Then Continue For

                    Dim slot As Integer = 0
                    If slotByKey.TryGetValue(entry.Key.Trim(), slot) Then
                        ApplyColorSlot(entry, slot)
                    End If
                Next
            Next
        End Sub

        Private Shared Function LoadPersistedSessions() As List(Of DocumentVisualAidSessionState)
            Dim results As New List(Of DocumentVisualAidSessionState)()
            If Not Directory.Exists(SessionsDirectoryPath) Then Return results

            Dim filePaths As String() = Array.Empty(Of String)()
            Try
                filePaths = Directory.GetFiles(SessionsDirectoryPath, "*.json")
            Catch
                filePaths = Array.Empty(Of String)()
            End Try

            For Each filePath In filePaths
                Dim model = TryReadJson(Of DocumentVisualAidSessionFileModel)(filePath)
                If model Is Nothing Then Continue For

                If model.ProcessId <= 0 OrElse Not IsSessionAlive(model) Then
                    SafeDeleteFile(filePath)
                    Continue For
                End If

                results.Add(ConvertToRuntimeState(model))
            Next

            Return results
        End Function

        Private Shared Function IsOwnerValid(owner As DocumentVisualAidOwnerFileModel) As Boolean
            If owner Is Nothing OrElse owner.ProcessId <= 0 Then Return False
            If Not IsProcessAlive(owner.ProcessId) Then Return False

            Dim session = TryReadJson(Of DocumentVisualAidSessionFileModel)(GetSessionFilePath(owner.ProcessId))
            If session Is Nothing Then Return False

            Return session.ProcessId = owner.ProcessId AndAlso
                   session.StartedUtcTicks = owner.StartedUtcTicks AndAlso
                   IsSessionAlive(session)
        End Function

        Private Shared Function IsSessionAlive(model As DocumentVisualAidSessionFileModel) As Boolean
            If model Is Nothing OrElse model.ProcessId <= 0 Then Return False

            If IsProcessAlive(model.ProcessId) Then Return True

            Dim lastUpdateUtc = New DateTime(Math.Max(0L, model.UpdatedUtcTicks), DateTimeKind.Utc)
            Return (DateTime.UtcNow - lastUpdateUtc) <= TimeSpan.FromSeconds(5)
        End Function

        Private Shared Function IsProcessAlive(processId As Integer) As Boolean
            Try
                Dim proc As Process = Process.GetProcessById(processId)
                Return proc IsNot Nothing AndAlso Not proc.HasExited
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ConvertToRuntimeState(model As DocumentVisualAidSessionFileModel) As DocumentVisualAidSessionState
            Return New DocumentVisualAidSessionState With {
                .ProcessId = model.ProcessId,
                .VersionLabel = If(model.VersionLabel, String.Empty),
                .StartedUtcTicks = model.StartedUtcTicks,
                .UpdatedUtcTicks = model.UpdatedUtcTicks,
                .Entries = ConvertEntriesToRuntime(model.Documents),
                .Navigator = ConvertNavigatorToRuntime(model.Navigator),
                .IsLocalProcess = model.ProcessId = LocalProcessId
            }
        End Function

        Private Shared Function ConvertEntriesToFileModel(entries As IEnumerable(Of DocumentColorEntry)) As List(Of DocumentVisualAidEntryFileModel)
            Dim results As New List(Of DocumentVisualAidEntryFileModel)()
            If entries Is Nothing Then Return results

            For Each entry In entries
                If entry Is Nothing Then Continue For

                results.Add(New DocumentVisualAidEntryFileModel With {
                    .Key = entry.Key,
                    .DisplayName = entry.DisplayName,
                    .FullPath = entry.FullPath,
                    .ColorSlot = entry.ColorSlot,
                    .IsActive = entry.IsActive
                })
            Next

            Return results
        End Function

        Private Shared Function ConvertEntriesToRuntime(entries As IEnumerable(Of DocumentVisualAidEntryFileModel)) As List(Of DocumentColorEntry)
            Dim results As New List(Of DocumentColorEntry)()
            If entries Is Nothing Then Return results

            For Each item In entries
                If item Is Nothing Then Continue For

                Dim entry = New DocumentColorEntry With {
                    .Key = item.Key,
                    .DisplayName = item.DisplayName,
                    .FullPath = item.FullPath,
                    .ColorSlot = Math.Max(0, item.ColorSlot),
                    .IsActive = item.IsActive,
                    .MatchTokens = New List(Of String)()
                }
                ApplyColorSlot(entry, entry.ColorSlot)
                results.Add(entry)
            Next

            Return results
        End Function

        Private Shared Function ConvertNavigatorToFileModel(navigator As DocumentViewNavigatorSnapshot) As DocumentVisualAidNavigatorFileModel
            If navigator Is Nothing Then Return Nothing

            Dim documents As New List(Of DocumentVisualAidNavigatorDocumentFileModel)()
            If navigator.Documents IsNot Nothing Then
                For Each documentItem In navigator.Documents
                    If documentItem Is Nothing Then Continue For

                    Dim categories As New List(Of DocumentVisualAidNavigatorCategoryFileModel)()
                    If documentItem.Categories IsNot Nothing Then
                        For Each category In documentItem.Categories
                            If category Is Nothing Then Continue For

                            Dim views As New List(Of DocumentVisualAidNavigatorViewFileModel)()
                            If category.Views IsNot Nothing Then
                                For Each viewItem In category.Views
                                    If viewItem Is Nothing Then Continue For

                                    views.Add(New DocumentVisualAidNavigatorViewFileModel With {
                                        .DocumentKey = viewItem.DocumentKey,
                                        .ViewId = viewItem.ViewId,
                                        .CategoryKey = viewItem.CategoryKey,
                                        .DisplayName = viewItem.DisplayName
                                    })
                                Next
                            End If

                            categories.Add(New DocumentVisualAidNavigatorCategoryFileModel With {
                                .Key = category.Key,
                                .DisplayName = category.DisplayName,
                                .Views = views
                            })
                        Next
                    End If

                    documents.Add(New DocumentVisualAidNavigatorDocumentFileModel With {
                        .DocumentKey = documentItem.DocumentKey,
                        .DocumentName = documentItem.DocumentName,
                        .ActiveViewId = documentItem.ActiveViewId,
                        .DefaultCategoryKey = documentItem.DefaultCategoryKey,
                        .Categories = categories
                    })
                Next
            End If

            Return New DocumentVisualAidNavigatorFileModel With {
                .SelectedDocumentKey = navigator.SelectedDocumentKey,
                .Documents = documents
            }
        End Function

        Private Shared Function ConvertNavigatorToRuntime(navigator As DocumentVisualAidNavigatorFileModel) As DocumentViewNavigatorSnapshot
            If navigator Is Nothing Then Return Nothing

            Dim documents As New List(Of DocumentViewNavigatorDocument)()
            If navigator.Documents IsNot Nothing Then
                For Each documentItem In navigator.Documents
                    If documentItem Is Nothing Then Continue For

                    Dim categories As New List(Of DocumentViewCategory)()
                    If documentItem.Categories IsNot Nothing Then
                        For Each category In documentItem.Categories
                            If category Is Nothing Then Continue For

                            Dim views As New List(Of DocumentViewOption)()
                            If category.Views IsNot Nothing Then
                                For Each viewItem In category.Views
                                    If viewItem Is Nothing Then Continue For

                                    views.Add(New DocumentViewOption With {
                                        .DocumentKey = viewItem.DocumentKey,
                                        .ViewId = viewItem.ViewId,
                                        .CategoryKey = viewItem.CategoryKey,
                                        .DisplayName = viewItem.DisplayName
                                    })
                                Next
                            End If

                            categories.Add(New DocumentViewCategory With {
                                .Key = category.Key,
                                .DisplayName = category.DisplayName,
                                .Views = views
                            })
                        Next
                    End If

                    documents.Add(New DocumentViewNavigatorDocument With {
                        .DocumentKey = documentItem.DocumentKey,
                        .DocumentName = documentItem.DocumentName,
                        .ActiveViewId = documentItem.ActiveViewId,
                        .DefaultCategoryKey = documentItem.DefaultCategoryKey,
                        .Categories = categories
                    })
                Next
            End If

            Return New DocumentViewNavigatorSnapshot With {
                .SelectedDocumentKey = navigator.SelectedDocumentKey,
                .Documents = documents
            }
        End Function

        Private Shared Function ResolveVersionLabel(app As RevitApp) As String
            If app IsNot Nothing Then
                Try
                    Dim raw = If(app.VersionNumber, String.Empty).Trim()
                    If raw.Length > 0 Then Return raw
                Catch
                End Try
            End If

#If REVIT2027 Then
            Return "2027"
#ElseIf REVIT2025 Then
            Return "2025"
#ElseIf REVIT2023 Then
            Return "2023"
#ElseIf REVIT2021 Then
            Return "2021"
#ElseIf REVIT2019 Then
            Return "2019"
#Else
            Return "Revit"
#End If
        End Function

        Private Shared Sub WriteCommand(targetProcessId As Integer, command As DocumentVisualAidCommandFileModel)
            If command Is Nothing Then Return

            EnsureDirectories()

            Dim commandDirectory = GetCommandDirectoryPath(targetProcessId)
            Try
                Directory.CreateDirectory(commandDirectory)
            Catch
            End Try

            Dim fileName = $"cmd-{DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture)}-{Guid.NewGuid():N}.json"
            WriteJson(Path.Combine(commandDirectory, fileName), command)
        End Sub

        Private Shared Function GetSessionFilePath(processId As Integer) As String
            Return Path.Combine(SessionsDirectoryPath, $"session-{processId.ToString(CultureInfo.InvariantCulture)}.json")
        End Function

        Private Shared Function GetCommandDirectoryPath(processId As Integer) As String
            Return Path.Combine(CommandsDirectoryPath, processId.ToString(CultureInfo.InvariantCulture))
        End Function

        Private Shared Sub EnsureDirectories()
            Try
                Directory.CreateDirectory(RootDirectoryPath)
                Directory.CreateDirectory(SessionsDirectoryPath)
                Directory.CreateDirectory(CommandsDirectoryPath)
            Catch
            End Try
        End Sub

        Private Shared Sub WriteJson(filePath As String, value As Object)
            If String.IsNullOrWhiteSpace(filePath) Then Return

            Try
                Dim directoryPath = IO.Path.GetDirectoryName(filePath)
                If Not String.IsNullOrWhiteSpace(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If

                Dim tempPath = filePath & "." & LocalProcessId.ToString(CultureInfo.InvariantCulture) & ".tmp"
                File.WriteAllText(tempPath, Serializer.Serialize(value))
                File.Copy(tempPath, filePath, True)
                SafeDeleteFile(tempPath)
            Catch
            End Try
        End Sub

        Private Shared Function TryReadJson(Of T)(path As String) As T
            Try
                If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                    Return Nothing
                End If

                Dim raw = File.ReadAllText(path)
                If String.IsNullOrWhiteSpace(raw) Then
                    Return Nothing
                End If

                Return Serializer.Deserialize(Of T)(raw)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub SafeDeleteFile(path As String)
            Try
                If Not String.IsNullOrWhiteSpace(path) AndAlso File.Exists(path) Then
                    File.Delete(path)
                End If
            Catch
            End Try
        End Sub

        Private Shared Sub SafeDeleteDirectory(path As String)
            Try
                If Not String.IsNullOrWhiteSpace(path) AndAlso Directory.Exists(path) Then
                    Directory.Delete(path, True)
                End If
            Catch
            End Try
        End Sub

        Private Shared Function CreateFrozenBrush(color As WpfColor) As WpfSolidColorBrush
            Dim brush As New WpfSolidColorBrush(color)
            brush.Freeze()
            Return brush
        End Function

        Private Shared Sub ApplyColorSlot(entry As DocumentColorEntry, slot As Integer)
            If entry Is Nothing Then Return

            Dim normalizedSlot = Math.Max(0, slot)
            Dim baseColor = ResolveBaseColor(normalizedSlot)
            Dim chipColor = BlendWithWhite(baseColor, 0.28R)
            Dim tabColor = BlendWithWhite(baseColor, 0.72R)
            Dim activeTabColor = BlendWithWhite(baseColor, 0.64R)
            Dim borderColor = BlendWithWhite(baseColor, 0.55R)
            Dim accentColor = BlendWithWhite(baseColor, 0.48R)

            entry.ColorSlot = normalizedSlot
            entry.ChipBrush = CreateFrozenBrush(WpfColor.FromArgb(230, chipColor.R, chipColor.G, chipColor.B))
            entry.TabFillBrush = CreateFrozenBrush(WpfColor.FromArgb(34, tabColor.R, tabColor.G, tabColor.B))
            entry.ActiveTabFillBrush = CreateFrozenBrush(WpfColor.FromArgb(52, activeTabColor.R, activeTabColor.G, activeTabColor.B))
            entry.BorderBrush = CreateFrozenBrush(WpfColor.FromArgb(145, borderColor.R, borderColor.G, borderColor.B))
            entry.AccentBrush = CreateFrozenBrush(WpfColor.FromArgb(160, accentColor.R, accentColor.G, accentColor.B))
        End Sub

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

        Private Shared Function ResolveMainWindowHandle(processId As Integer) As IntPtr
            Try
                Dim proc As Process = Process.GetProcessById(processId)
                If proc IsNot Nothing AndAlso Not proc.HasExited Then
                    If proc.MainWindowHandle <> IntPtr.Zero Then
                        Return proc.MainWindowHandle
                    End If
                End If
            Catch
            End Try

            Dim found As IntPtr = IntPtr.Zero
            EnumWindows(Function(hWnd, lParam)
                            If found <> IntPtr.Zero Then Return False
                            If Not IsWindowVisible(hWnd) Then Return True

                            Dim ownerPid As UInteger
                            GetWindowThreadProcessId(hWnd, ownerPid)
                            If ownerPid <> CUInt(processId) Then Return True

                            If GetWindowTextLength(hWnd) <= 0 Then Return True

                            found = hWnd
                            Return False
                        End Function,
                        IntPtr.Zero)
            Return found
        End Function

        Private Const SW_RESTORE As Integer = 9

        <DllImport("user32.dll")>
        Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

        <DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef processId As UInteger) As UInteger
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function BringWindowToTop(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

    End Class

    Friend NotInheritable Class DocumentVisualAidAggregateState
        Friend Property Sessions As IList(Of DocumentVisualAidSessionState)
    End Class

    Friend NotInheritable Class DocumentVisualAidSessionState
        Friend Property ProcessId As Integer
        Friend Property VersionLabel As String
        Friend Property TabLabel As String
        Friend Property TabToolTip As String
        Friend Property StartedUtcTicks As Long
        Friend Property UpdatedUtcTicks As Long
        Friend Property Entries As IList(Of DocumentColorEntry)
        Friend Property Navigator As DocumentViewNavigatorSnapshot
        Friend Property IsLocalProcess As Boolean
    End Class

    Public Class DocumentVisualAidOwnerFileModel
        Public Property ProcessId As Integer
        Public Property StartedUtcTicks As Long
    End Class

    Public Class DocumentVisualAidCommandFileModel
        Public Property CommandName As String
        Public Property DocumentKey As String
        Public Property ViewId As Integer
    End Class

    Public Class DocumentVisualAidSessionFileModel
        Public Property ProcessId As Integer
        Public Property VersionLabel As String
        Public Property StartedUtcTicks As Long
        Public Property UpdatedUtcTicks As Long
        Public Property Documents As List(Of DocumentVisualAidEntryFileModel)
        Public Property Navigator As DocumentVisualAidNavigatorFileModel
    End Class

    Public Class DocumentVisualAidEntryFileModel
        Public Property Key As String
        Public Property DisplayName As String
        Public Property FullPath As String
        Public Property ColorSlot As Integer
        Public Property IsActive As Boolean
    End Class

    Public Class DocumentVisualAidNavigatorFileModel
        Public Property SelectedDocumentKey As String
        Public Property Documents As List(Of DocumentVisualAidNavigatorDocumentFileModel)
    End Class

    Public Class DocumentVisualAidNavigatorDocumentFileModel
        Public Property DocumentKey As String
        Public Property DocumentName As String
        Public Property ActiveViewId As Integer
        Public Property DefaultCategoryKey As String
        Public Property Categories As List(Of DocumentVisualAidNavigatorCategoryFileModel)
    End Class

    Public Class DocumentVisualAidNavigatorCategoryFileModel
        Public Property Key As String
        Public Property DisplayName As String
        Public Property Views As List(Of DocumentVisualAidNavigatorViewFileModel)
    End Class

    Public Class DocumentVisualAidNavigatorViewFileModel
        Public Property DocumentKey As String
        Public Property ViewId As Integer
        Public Property CategoryKey As String
        Public Property DisplayName As String
    End Class

End Namespace
