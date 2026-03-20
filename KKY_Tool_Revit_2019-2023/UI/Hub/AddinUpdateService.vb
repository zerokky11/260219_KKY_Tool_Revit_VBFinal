Option Explicit On
Option Strict On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization

Namespace UI.Hub

    Friend NotInheritable Class AddinUpdateService

        Private Shared ReadOnly _serializer As New JavaScriptSerializer()

        Friend Class UpdateConfig
            Public Property FeedUrl As String
            Public Property FeedUrls As List(Of String)
            Public Property DownloadDirectory As String
            Public Property ConfigPath As String
        End Class

        Friend Class UpdateInfo
            Public Property CurrentVersion As String
            Public Property CurrentVersionDisplay As String
            Public Property LatestVersion As String
            Public Property HasUpdate As Boolean
            Public Property CanInstall As Boolean
            Public Property IsConfigured As Boolean
            Public Property FeedUrl As String
            Public Property ConfigPath As String
            Public Property DownloadUrl As String
            Public Property Notes As String
            Public Property PublishedAt As String
            Public Property Message As String
        End Class

        Private Class UpdateManifest
            Public Property Version As String
            Public Property Url As String
            Public Property Sha256 As String
            Public Property Notes As String
            Public Property PublishedAt As String
            Public Property FileName As String
        End Class

        Private Sub New()
        End Sub

        Friend Shared Function CreateInitialInfo() As UpdateInfo
            Dim cfg = LoadConfig()
            Return New UpdateInfo With {
                .CurrentVersion = GetCurrentVersionDisplay(),
                .CurrentVersionDisplay = GetCurrentVersionDisplay(),
                .LatestVersion = String.Empty,
                .HasUpdate = False,
                .CanInstall = False,
                .IsConfigured = GetConfiguredFeedUrls(cfg).Count > 0,
                .FeedUrl = GetPrimaryFeedUrl(cfg),
                .ConfigPath = GetPreferredConfigPath(),
                .Message = String.Empty
            }
        End Function

        Friend Shared Function CheckForUpdates() As UpdateInfo
            Dim info = CreateInitialInfo()
            Dim cfg = LoadConfig()
            Dim feedUrls = GetConfiguredFeedUrls(cfg)

            If feedUrls.Count = 0 Then
                info.IsConfigured = False
                info.Message = "업데이트 피드가 설정되지 않았습니다. Resources\update-config.json 파일에 feedUrl 또는 feedUrls 를 입력해 주세요."
                Return info
            End If

            Dim chosenFeedUrl As String = String.Empty
            Dim manifest As UpdateManifest = Nothing
            Dim lastError As Exception = Nothing

            For Each feedUrl In feedUrls
                Try
                    manifest = ParseManifest(ReadText(feedUrl))
                    chosenFeedUrl = feedUrl
                    Exit For
                Catch ex As Exception
                    lastError = ex
                End Try
            Next

            If manifest Is Nothing Then
                Dim reason = If(lastError Is Nothing, "알 수 없는 오류", lastError.Message)
                Throw New InvalidOperationException("모든 업데이트 피드에 연결하지 못했습니다. " & reason)
            End If

            If String.IsNullOrWhiteSpace(manifest.Version) Then
                Throw New InvalidDataException("업데이트 피드에 version 값이 없습니다.")
            End If

            info.IsConfigured = True
            info.FeedUrl = chosenFeedUrl
            info.ConfigPath = cfg.ConfigPath
            info.LatestVersion = CleanVersionDisplay(manifest.Version)
            info.DownloadUrl = ResolveLocation(chosenFeedUrl, manifest.Url)
            info.Notes = manifest.Notes
            info.PublishedAt = manifest.PublishedAt
            info.HasUpdate = CompareVersions(info.LatestVersion, info.CurrentVersionDisplay) > 0
            info.CanInstall = info.HasUpdate AndAlso Not String.IsNullOrWhiteSpace(info.DownloadUrl)

            If info.HasUpdate Then
                info.Message = String.Format("새 버전 {0} 이(가) 있습니다. 현재 버전은 {1} 입니다.", info.LatestVersion, info.CurrentVersionDisplay)
            Else
                info.Message = String.Format("현재 최신 버전({0})을 사용 중입니다.", info.CurrentVersionDisplay)
            End If

            Return info
        End Function

        Friend Shared Function PrepareInstaller(info As UpdateInfo) As String
            If info Is Nothing Then Throw New ArgumentNullException(NameOf(info))
            If String.IsNullOrWhiteSpace(info.DownloadUrl) Then
                Throw New InvalidOperationException("다운로드할 설치파일 주소가 없습니다.")
            End If

            Dim cfg = LoadConfig()
            Dim downloadRoot = ResolveDownloadRoot(cfg)
            Dim versionFolder = Path.Combine(downloadRoot, SanitizePathSegment(If(info.LatestVersion, "latest")))
            Directory.CreateDirectory(versionFolder)

            Dim source = info.DownloadUrl.Trim()
            Dim fileName = DetermineFileName(source, info.LatestVersion)
            Dim localPath = Path.Combine(versionFolder, fileName)

            If IsHttpUrl(source) Then
                Using wc As New WebClient()
                    wc.Headers(HttpRequestHeader.UserAgent) = "KKY_Tool_Revit/" & GetCurrentVersionDisplay()
                    wc.DownloadFile(source, localPath)
                End Using
            Else
                Dim sourcePath = ResolveLocalPath(source)
                If Not File.Exists(sourcePath) Then
                    Throw New FileNotFoundException("업데이트 설치파일을 찾을 수 없습니다.", sourcePath)
                End If

                If Not String.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(localPath), StringComparison.OrdinalIgnoreCase) Then
                    File.Copy(sourcePath, localPath, True)
                Else
                    localPath = sourcePath
                End If
            End If

            Return localPath
        End Function

        Friend Shared Function QueueInstallerAfterProcessExit(installerPath As String) As String
            If String.IsNullOrWhiteSpace(installerPath) Then Throw New ArgumentNullException(NameOf(installerPath))
            If Not File.Exists(installerPath) Then Throw New FileNotFoundException("설치파일을 찾을 수 없습니다.", installerPath)

            Dim scriptDir = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "UpdateQueue")
            Directory.CreateDirectory(scriptDir)

            Dim stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            Dim scriptPath = Path.Combine(scriptDir, "run_update_" & stamp & ".cmd")
            Dim pidText = Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
            Dim command = BuildInstallerLaunchCommand(installerPath)

            Dim lines As New List(Of String) From {
                "@echo off",
                "setlocal",
                "set PID=" & pidText,
                ":wait_loop",
                "tasklist /FI ""PID eq %PID%"" 2>NUL | find ""%PID%"" >NUL",
                "if not errorlevel 1 (",
                "  timeout /t 2 /nobreak >NUL",
                "  goto wait_loop",
                ")",
                "timeout /t 1 /nobreak >NUL",
                command,
                "endlocal"
            }

            File.WriteAllLines(scriptPath, lines.ToArray(), Encoding.ASCII)

            Dim psi As New ProcessStartInfo()
            psi.FileName = "cmd.exe"
            psi.Arguments = "/c start """" """ & scriptPath & """"
            psi.UseShellExecute = False
            psi.CreateNoWindow = True
            psi.WindowStyle = ProcessWindowStyle.Hidden

            Process.Start(psi)
            Return scriptPath
        End Function

        Friend Shared Function GetCurrentVersionDisplay() As String
            Dim asm = Assembly.GetExecutingAssembly()

            Try
                For Each attrObj As Object In asm.GetCustomAttributes(GetType(AssemblyInformationalVersionAttribute), False)
                    Dim attr = TryCast(attrObj, AssemblyInformationalVersionAttribute)
                    If attr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(attr.InformationalVersion) Then
                        Return CleanVersionDisplay(attr.InformationalVersion)
                    End If
                Next
            Catch
            End Try

            Try
                Dim fv = FileVersionInfo.GetVersionInfo(asm.Location)
                If fv IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(fv.FileVersion) Then
                    Return CleanVersionDisplay(fv.FileVersion)
                End If
            Catch
            End Try

            Return "0.0"
        End Function

        Private Shared Function LoadConfig() As UpdateConfig
            Dim candidates As New List(Of String)()
            Dim envPath = Environment.GetEnvironmentVariable("KKY_TOOL_UPDATE_CONFIG")
            If Not String.IsNullOrWhiteSpace(envPath) Then candidates.Add(envPath)

            candidates.Add(GetPreferredConfigPath())

            Try
                Dim asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                If Not String.IsNullOrWhiteSpace(asmDir) Then
                    candidates.Add(Path.Combine(asmDir, "update-config.json"))
                End If
            Catch
            End Try

            For Each rawCandidate In candidates
                If String.IsNullOrWhiteSpace(rawCandidate) Then Continue For

                Dim candidate = ResolveLocalPath(rawCandidate)
                If Not File.Exists(candidate) Then Continue For

                Dim root = DeserializeObject(File.ReadAllText(candidate, Encoding.UTF8))
                Dim cfg As New UpdateConfig()
                cfg.FeedUrl = ReadString(root, "feedUrl")
                cfg.FeedUrls = ReadStringList(root, "feedUrls")
                cfg.DownloadDirectory = ReadString(root, "downloadDirectory")
                cfg.ConfigPath = candidate
                Return cfg
            Next

            Return New UpdateConfig With {
                .FeedUrls = New List(Of String)(),
                .ConfigPath = GetPreferredConfigPath()
            }
        End Function

        Private Shared Function ParseManifest(text As String) As UpdateManifest
            Dim root = DeserializeObject(text)
            Dim manifest As New UpdateManifest()
            manifest.Version = ReadString(root, "version")
            manifest.Url = ReadString(root, "url")
            manifest.Sha256 = ReadString(root, "sha256")
            manifest.Notes = ReadString(root, "notes")
            manifest.PublishedAt = ReadString(root, "publishedAt")
            manifest.FileName = ReadString(root, "fileName")
            Return manifest
        End Function

        Private Shared Function DeserializeObject(text As String) As Dictionary(Of String, Object)
            Dim dict = _serializer.Deserialize(Of Dictionary(Of String, Object))(text)
            If dict Is Nothing Then
                Throw New InvalidDataException("JSON 형식이 올바르지 않습니다.")
            End If
            Return dict
        End Function

        Private Shared Function ReadString(root As Dictionary(Of String, Object), key As String) As String
            If root Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return String.Empty
            If Not root.ContainsKey(key) OrElse root(key) Is Nothing Then Return String.Empty
            Return Convert.ToString(root(key), CultureInfo.InvariantCulture).Trim()
        End Function

        Private Shared Function ReadStringList(root As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim results As New List(Of String)()
            If root Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return results
            If Not root.ContainsKey(key) OrElse root(key) Is Nothing Then Return results

            Dim items = TryCast(root(key), ArrayList)
            If items Is Nothing Then Return results

            For Each rawItem In items
                Dim value = Convert.ToString(rawItem, CultureInfo.InvariantCulture).Trim()
                If Not String.IsNullOrWhiteSpace(value) AndAlso Not results.Contains(value) Then
                    results.Add(value)
                End If
            Next

            Return results
        End Function

        Private Shared Function ReadText(location As String) As String
            Dim candidate = (If(location, String.Empty)).Trim()
            If String.IsNullOrWhiteSpace(candidate) Then
                Throw New InvalidOperationException("업데이트 피드 주소가 비어 있습니다.")
            End If

            If IsHttpUrl(candidate) Then
                Using wc As New WebClient()
                    wc.Headers(HttpRequestHeader.UserAgent) = "KKY_Tool_Revit/" & GetCurrentVersionDisplay()
                    Return wc.DownloadString(candidate)
                End Using
            End If

            Dim localPath = ResolveLocalPath(candidate)
            If Not File.Exists(localPath) Then
                Throw New FileNotFoundException("업데이트 피드 파일을 찾을 수 없습니다.", localPath)
            End If

            Return File.ReadAllText(localPath, Encoding.UTF8)
        End Function

        Private Shared Function ResolveLocation(baseLocation As String, candidate As String) As String
            Dim value = (If(candidate, String.Empty)).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            If IsHttpUrl(value) OrElse value.StartsWith("\\", StringComparison.Ordinal) Then
                Return Environment.ExpandEnvironmentVariables(value)
            End If

            Dim absoluteUri As Uri = Nothing
            If Uri.TryCreate(value, UriKind.Absolute, absoluteUri) Then
                If absoluteUri.IsFile Then Return absoluteUri.LocalPath
                Return absoluteUri.ToString()
            End If

            Dim expandedValue = Environment.ExpandEnvironmentVariables(value)
            If Path.IsPathRooted(expandedValue) Then Return expandedValue

            If IsHttpUrl(baseLocation) Then
                Dim baseUri As New Uri(baseLocation)
                Return New Uri(baseUri, value).ToString()
            End If

            Dim basePath = ResolveLocalPath(baseLocation)
            Dim baseDir = If(File.Exists(basePath), Path.GetDirectoryName(basePath), basePath)
            If String.IsNullOrWhiteSpace(baseDir) Then Return expandedValue

            Return Path.GetFullPath(Path.Combine(baseDir, expandedValue))
        End Function

        Private Shared Function ResolveLocalPath(pathOrUri As String) As String
            Dim value = Environment.ExpandEnvironmentVariables(If(pathOrUri, String.Empty)).Trim()
            Dim absoluteUri As Uri = Nothing

            If Uri.TryCreate(value, UriKind.Absolute, absoluteUri) AndAlso absoluteUri.IsFile Then
                Return absoluteUri.LocalPath
            End If

            Return value
        End Function

        Private Shared Function ResolveDownloadRoot(cfg As UpdateConfig) As String
            Dim configured = If(cfg Is Nothing, String.Empty, cfg.DownloadDirectory)
            If Not String.IsNullOrWhiteSpace(configured) Then
                Return ResolveLocalPath(configured)
            End If

            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY_Tool_Revit", "Updates")
        End Function

        Private Shared Function GetConfiguredFeedUrls(cfg As UpdateConfig) As List(Of String)
            Dim results As New List(Of String)()
            If cfg Is Nothing Then Return results

            If cfg.FeedUrls IsNot Nothing Then
                For Each rawUrl In cfg.FeedUrls
                    Dim value = If(rawUrl, String.Empty).Trim()
                    If Not String.IsNullOrWhiteSpace(value) AndAlso Not ContainsIgnoreCase(results, value) Then
                        results.Add(value)
                    End If
                Next
            End If

            If Not String.IsNullOrWhiteSpace(cfg.FeedUrl) AndAlso Not ContainsIgnoreCase(results, cfg.FeedUrl.Trim()) Then
                results.Add(cfg.FeedUrl.Trim())
            End If

            Return results
        End Function

        Private Shared Function GetPrimaryFeedUrl(cfg As UpdateConfig) As String
            Dim feeds = GetConfiguredFeedUrls(cfg)
            If feeds.Count = 0 Then Return String.Empty
            Return feeds(0)
        End Function

        Private Shared Function ContainsIgnoreCase(values As List(Of String), target As String) As Boolean
            For Each item In values
                If String.Equals(item, target, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Function BuildInstallerLaunchCommand(installerPath As String) As String
            Dim ext = Path.GetExtension(installerPath)
            If String.Equals(ext, ".msi", StringComparison.OrdinalIgnoreCase) Then
                Return "start """" msiexec /i """ & installerPath & """"
            End If

            Return "start """" """ & installerPath & """"
        End Function

        Private Shared Function DetermineFileName(source As String, latestVersion As String) As String
            Dim candidateName As String = Nothing

            If IsHttpUrl(source) Then
                Dim uri As New Uri(source)
                candidateName = Path.GetFileName(uri.LocalPath)
            Else
                candidateName = Path.GetFileName(ResolveLocalPath(source))
            End If

            If String.IsNullOrWhiteSpace(candidateName) Then
                candidateName = "KKY_Tool_Revit_" & SanitizePathSegment(latestVersion) & ".exe"
            End If

            Return candidateName
        End Function

        Private Shared Function CleanVersionDisplay(raw As String) As String
            Dim value = (If(raw, String.Empty)).Trim()
            If value.StartsWith("v", StringComparison.OrdinalIgnoreCase) Then
                value = value.Substring(1)
            End If

            Dim dottedParts = value.Split("."c)
            If dottedParts.Length > 0 Then
                Dim allNumeric = True
                For Each part In dottedParts
                    If String.IsNullOrWhiteSpace(part) OrElse Not Regex.IsMatch(part, "^\d+$") Then
                        allNumeric = False
                        Exit For
                    End If
                Next

                If allNumeric Then
                    Dim visibleCount = dottedParts.Length
                    While visibleCount > 2 AndAlso String.Equals(dottedParts(visibleCount - 1), "0", StringComparison.Ordinal)
                        visibleCount -= 1
                    End While

                    Return String.Join(".", dottedParts, 0, visibleCount)
                End If
            End If

            Dim comparable = ParseVersion(value)
            If comparable Is Nothing Then Return value

            Dim parts As New List(Of String) From {
                comparable.Major.ToString(CultureInfo.InvariantCulture),
                comparable.Minor.ToString(CultureInfo.InvariantCulture)
            }

            If comparable.Build >= 0 AndAlso comparable.Build <> 0 Then
                parts.Add(comparable.Build.ToString(CultureInfo.InvariantCulture))
            End If

            If comparable.Revision >= 0 AndAlso comparable.Revision <> 0 Then
                parts.Add(comparable.Revision.ToString(CultureInfo.InvariantCulture))
            End If

            Return String.Join(".", parts.ToArray())
        End Function

        Private Shared Function CompareVersions(leftValue As String, rightValue As String) As Integer
            Dim leftVersion = ParseVersion(leftValue)
            Dim rightVersion = ParseVersion(rightValue)

            If leftVersion Is Nothing AndAlso rightVersion Is Nothing Then Return 0
            If leftVersion Is Nothing Then Return -1
            If rightVersion Is Nothing Then Return 1
            Return leftVersion.CompareTo(rightVersion)
        End Function

        Private Shared Function ParseVersion(raw As String) As Version
            Dim value = (If(raw, String.Empty)).Trim()
            If String.IsNullOrWhiteSpace(value) Then Return Nothing

            Dim matches = Regex.Matches(value, "\d+")
            If matches.Count = 0 Then Return Nothing

            Dim parts As New List(Of Integer)()
            For Each match As Match In matches
                If parts.Count = 4 Then Exit For
                Dim parsed As Integer = 0
                If Integer.TryParse(match.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    parts.Add(parsed)
                End If
            Next

            While parts.Count < 4
                parts.Add(0)
            End While

            Return New Version(parts(0), parts(1), parts(2), parts(3))
        End Function

        Private Shared Function SanitizePathSegment(value As String) As String
            Dim raw = If(value, "latest")
            For Each invalidChar In Path.GetInvalidFileNameChars()
                raw = raw.Replace(invalidChar, "_"c)
            Next
            Return raw.Replace(" ", "_").Trim("_"c)
        End Function

        Private Shared Function GetPreferredConfigPath() As String
            Try
                Dim asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                If String.IsNullOrWhiteSpace(asmDir) Then Return "update-config.json"
                Return Path.Combine(asmDir, "Resources", "update-config.json")
            Catch
                Return "update-config.json"
            End Try
        End Function

        Private Shared Function IsHttpUrl(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then Return False
            Return value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                   value.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
        End Function

    End Class

End Namespace
