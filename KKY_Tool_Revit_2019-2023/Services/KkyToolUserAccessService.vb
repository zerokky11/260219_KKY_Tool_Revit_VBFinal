Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web.Script.Serialization
Imports Autodesk.Revit.UI

Namespace Services
    Public NotInheritable Class KkyToolUserAccessService
        Private Const DefaultKeyword As String = "KCIM"
        Private Const DefaultAdminPassword As String = "KKYTOOL"
        Private Const DefaultRemoteConfigUrl As String = "https://update.zerokky.com/kky-tool/user-access.json"

        Public Const DefaultBlockMessage As String = "외부 사용자는 사용할 수 없습니다."

        Private Shared ReadOnly Serializer As New JavaScriptSerializer()

        Private Sub New()
        End Sub

        Public Shared Function Evaluate(uiapp As UIApplication) As KkyToolUserAccessEvaluation
            Dim config = LoadEffectiveConfig()
            Dim userName = ResolveRevitUserName(uiapp)
            Dim allowed = IsUserAllowed(config, userName)
            Return New KkyToolUserAccessEvaluation With {
                .Allowed = allowed,
                .UserName = userName,
                .Message = If(String.IsNullOrWhiteSpace(config.BlockMessage), DefaultBlockMessage, config.BlockMessage),
                .Config = config
            }
        End Function

        Public Shared Function ResolveRevitUserName(uiapp As UIApplication) As String
            Try
                If uiapp IsNot Nothing AndAlso uiapp.Application IsNot Nothing Then
                    Dim value = If(uiapp.Application.Username, String.Empty).Trim()
                    If Not String.IsNullOrWhiteSpace(value) Then Return value
                End If
            Catch
            End Try

            Try
                Dim value = If(Environment.UserName, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(value) Then Return value
            Catch
            End Try

            Return String.Empty
        End Function

        Public Shared Function LoadEffectiveConfig() As KkyToolUserAccessConfig
            Dim remote = TryLoadRemoteConfig()
            If remote IsNot Nothing Then Return remote

            Dim cached = TryLoadConfigFile(GetCachePath(), False)
            If cached IsNot Nothing Then
                cached.Source = "cache"
                cached.SourceUrl = GetRemoteConfigUrl()
                Return NormalizeConfig(cached)
            End If

            Dim local = TryLoadConfigFile(GetConfigPath(), False)
            If local IsNot Nothing Then
                local.Source = "local"
                local.SourceUrl = GetRemoteConfigUrl()
                Return NormalizeConfig(local)
            End If

            Dim fallback = CreateDefaultConfig()
            fallback.Source = "default"
            fallback.SourceUrl = GetRemoteConfigUrl()
            Return fallback
        End Function

        Public Shared Function LoadOrCreate() As KkyToolUserAccessConfig
            Dim existing = TryLoadConfigFile(GetConfigPath(), False)
            If existing IsNot Nothing Then Return existing

            Dim created = CreateDefaultConfig()
            SaveConfigFile(GetConfigPath(), created)
            Return created
        End Function

        Public Shared Function GetPublicSnapshot(uiapp As UIApplication, Optional authenticated As Boolean = False) As Object
            Dim eval = Evaluate(uiapp)
            Dim config = NormalizeConfig(eval.Config)
            Return New With {
                .enabled = config.Enabled,
                .currentUser = eval.UserName,
                .allowed = eval.Allowed,
                .message = eval.Message,
                .authenticated = authenticated,
                .requirePasswordChange = config.RequirePasswordChange,
                .configPath = GetRemoteConfigUrl(),
                .cachePath = GetCachePath(),
                .source = config.Source,
                .sourceUrl = config.SourceUrl,
                .allowedProfileKeywords = config.AllowedProfileKeywords.ToArray(),
                .allowedUsers = config.AllowedUsers.ToArray()
            }
        End Function

        Public Shared Function VerifyAdminPassword(password As String) As Boolean
            Dim config = LoadOrCreate()
            Return VerifyPassword(password, config.AdminPasswordSalt, config.AdminPasswordHash)
        End Function

        Public Shared Function SaveFromRequest(enabled As Boolean,
                                               keywords As IEnumerable(Of String),
                                               users As IEnumerable(Of String),
                                               blockMessage As String,
                                               currentPassword As String,
                                               newPassword As String) As KkyToolUserAccessConfig
            Dim config = LoadOrCreate()
            If Not VerifyPassword(currentPassword, config.AdminPasswordSalt, config.AdminPasswordHash) Then
                Throw New UnauthorizedAccessException("관리자 비밀번호가 맞지 않습니다.")
            End If

            config.Enabled = enabled
            config.AllowedProfileKeywords = NormalizeList(keywords)
            config.AllowedUsers = NormalizeList(users)
            config.BlockMessage = If(String.IsNullOrWhiteSpace(blockMessage), DefaultBlockMessage, blockMessage.Trim())

            If config.AllowedProfileKeywords.Count = 0 AndAlso config.AllowedUsers.Count = 0 Then
                config.AllowedProfileKeywords.Add(DefaultKeyword)
            End If

            If Not String.IsNullOrWhiteSpace(newPassword) Then
                SetPassword(config, newPassword.Trim())
                config.RequirePasswordChange = False
            End If

            config.UpdatedAtUtc = DateTime.UtcNow.ToString("o")
            SaveConfigFile(GetConfigPath(), config)
            Return config
        End Function

        Private Shared Function IsUserAllowed(config As KkyToolUserAccessConfig, userName As String) As Boolean
            config = NormalizeConfig(config)
            If Not config.Enabled Then Return True

            Dim user = If(userName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(user) Then Return False

            For Each allowedUser In config.AllowedUsers
                If String.Equals(user, allowedUser, StringComparison.OrdinalIgnoreCase) Then Return True
            Next

            For Each keyword In config.AllowedProfileKeywords
                If user.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next

            Return False
        End Function

        Private Shared Function CreateDefaultConfig() As KkyToolUserAccessConfig
            Dim config As New KkyToolUserAccessConfig With {
                .Enabled = True,
                .AllowedProfileKeywords = New List(Of String) From {DefaultKeyword},
                .AllowedUsers = New List(Of String)(),
                .BlockMessage = DefaultBlockMessage,
                .RequirePasswordChange = True,
                .SourceUrl = GetRemoteConfigUrl(),
                .UpdatedAtUtc = DateTime.UtcNow.ToString("o")
            }
            SetPassword(config, DefaultAdminPassword)
            Return config
        End Function

        Private Shared Function NormalizeConfig(config As KkyToolUserAccessConfig) As KkyToolUserAccessConfig
            If config Is Nothing Then config = CreateDefaultConfig()
            config.AllowedProfileKeywords = NormalizeList(config.AllowedProfileKeywords)
            config.AllowedUsers = NormalizeList(config.AllowedUsers)
            If config.AllowedProfileKeywords.Count = 0 AndAlso config.AllowedUsers.Count = 0 Then
                config.AllowedProfileKeywords.Add(DefaultKeyword)
            End If
            If String.IsNullOrWhiteSpace(config.BlockMessage) Then config.BlockMessage = DefaultBlockMessage
            If String.IsNullOrWhiteSpace(config.SourceUrl) Then config.SourceUrl = GetRemoteConfigUrl()
            If String.IsNullOrWhiteSpace(config.AdminPasswordSalt) OrElse String.IsNullOrWhiteSpace(config.AdminPasswordHash) Then
                SetPassword(config, DefaultAdminPassword)
                config.RequirePasswordChange = True
            End If
            Return config
        End Function

        Private Shared Function NormalizeList(values As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            If values Is Nothing Then Return result

            For Each raw In values
                Dim value = If(raw, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(value) Then Continue For
                If result.Any(Function(x) String.Equals(x, value, StringComparison.OrdinalIgnoreCase)) Then Continue For
                result.Add(value)
            Next

            Return result
        End Function

        Private Shared Sub SetPassword(config As KkyToolUserAccessConfig, password As String)
            Dim salt(15) As Byte
            Using rng = RandomNumberGenerator.Create()
                rng.GetBytes(salt)
            End Using

            config.AdminPasswordSalt = Convert.ToBase64String(salt)
            config.AdminPasswordHash = HashPassword(password, salt)
        End Sub

        Private Shared Function VerifyPassword(password As String, saltText As String, hashText As String) As Boolean
            If String.IsNullOrWhiteSpace(password) OrElse String.IsNullOrWhiteSpace(saltText) OrElse String.IsNullOrWhiteSpace(hashText) Then
                Return False
            End If

            Try
                Dim salt = Convert.FromBase64String(saltText)
                Dim actual = HashPassword(password, salt)
                Return FixedTimeEquals(actual, hashText)
            Catch
                Return False
            End Try
        End Function

        Private Shared Function HashPassword(password As String, salt As Byte()) As String
            Using pbkdf2 As New Rfc2898DeriveBytes(If(password, String.Empty), salt, 100000, HashAlgorithmName.SHA256)
                Return Convert.ToBase64String(pbkdf2.GetBytes(32))
            End Using
        End Function

        Private Shared Function FixedTimeEquals(left As String, right As String) As Boolean
            Dim a = Encoding.UTF8.GetBytes(If(left, String.Empty))
            Dim b = Encoding.UTF8.GetBytes(If(right, String.Empty))
            Dim diff As Integer = a.Length Xor b.Length
            Dim length = Math.Min(a.Length, b.Length)
            For i As Integer = 0 To length - 1
                diff = diff Or (a(i) Xor b(i))
            Next
            Return diff = 0
        End Function

        Private Shared Function TryLoadRemoteConfig() As KkyToolUserAccessConfig
            Dim url = GetRemoteConfigUrl()
            If String.IsNullOrWhiteSpace(url) Then Return Nothing

            Try
                Using wc As New TimeoutWebClient(5000)
                    wc.Encoding = Encoding.UTF8
                    wc.Headers(HttpRequestHeader.UserAgent) = "KKY_Tool_Revit/UserAccess"
                    wc.Headers(HttpRequestHeader.CacheControl) = "no-cache"
                    Dim raw = wc.DownloadString(url)
                    Dim config = Serializer.Deserialize(Of KkyToolUserAccessConfig)(raw)
                    config = NormalizeConfig(config)
                    config.Source = "remote"
                    config.SourceUrl = url
                    SaveCache(config)
                    Return config
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function TryLoadConfigFile(path As String, createIfMissing As Boolean) As KkyToolUserAccessConfig
            Try
                If Not File.Exists(path) Then
                    If Not createIfMissing Then Return Nothing
                    Dim created = CreateDefaultConfig()
                    SaveConfigFile(path, created)
                    Return created
                End If

                Dim raw = File.ReadAllText(path, Encoding.UTF8)
                Dim config = Serializer.Deserialize(Of KkyToolUserAccessConfig)(raw)
                Return NormalizeConfig(config)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub SaveCache(config As KkyToolUserAccessConfig)
            Try
                config.CachedAtUtc = DateTime.UtcNow.ToString("o")
                SaveConfigFile(GetCachePath(), config)
            Catch
            End Try
        End Sub

        Private Shared Sub SaveConfigFile(path As String, config As KkyToolUserAccessConfig)
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
            Dim json = Serializer.Serialize(NormalizeConfig(config))
            File.WriteAllText(path, json, New UTF8Encoding(False))
        End Sub

        Private Shared Function GetConfigPath() As String
            Dim root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Return Path.Combine(root, "KKY_Tool_Revit", "Security", "user-access.json")
        End Function

        Private Shared Function GetCachePath() As String
            Dim root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Return Path.Combine(root, "KKY_Tool_Revit", "Security", "user-access-cache.json")
        End Function

        Private Shared Function GetRemoteConfigUrl() As String
            Try
                Dim env = Environment.GetEnvironmentVariable("KKY_TOOL_USER_ACCESS_URL")
                If Not String.IsNullOrWhiteSpace(env) Then Return env.Trim()
            Catch
            End Try
            Return DefaultRemoteConfigUrl
        End Function

        Private NotInheritable Class TimeoutWebClient
            Inherits WebClient

            Private ReadOnly _timeoutMilliseconds As Integer

            Public Sub New(timeoutMilliseconds As Integer)
                _timeoutMilliseconds = timeoutMilliseconds
            End Sub

            Protected Overrides Function GetWebRequest(address As Uri) As WebRequest
                Dim request = MyBase.GetWebRequest(address)
                If request IsNot Nothing Then request.Timeout = _timeoutMilliseconds
                Return request
            End Function
        End Class
    End Class

    Public Class KkyToolUserAccessEvaluation
        Public Property Allowed As Boolean
        Public Property UserName As String
        Public Property Message As String
        Public Property Config As KkyToolUserAccessConfig
    End Class

    Public Class KkyToolUserAccessConfig
        Public Property Enabled As Boolean = True
        Public Property AllowedProfileKeywords As List(Of String) = New List(Of String)()
        Public Property AllowedUsers As List(Of String) = New List(Of String)()
        Public Property BlockMessage As String = KkyToolUserAccessService.DefaultBlockMessage
        Public Property AdminPasswordSalt As String = String.Empty
        Public Property AdminPasswordHash As String = String.Empty
        Public Property RequirePasswordChange As Boolean = True
        Public Property UpdatedAtUtc As String = String.Empty
        Public Property CachedAtUtc As String = String.Empty
        Public Property Source As String = String.Empty
        Public Property SourceUrl As String = String.Empty
    End Class
End Namespace
