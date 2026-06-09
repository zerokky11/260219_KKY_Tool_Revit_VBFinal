Option Explicit On
Option Strict On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Services

Namespace UI.Hub
    Partial Public Class UiBridgeExternalEvent
        Private Sub HandleUserManagementInit(app As UIApplication, payload As Object)
            SendToWeb("usermanagement:state", KkyToolUserAccessService.GetPublicSnapshot(app, False))
        End Sub

        Private Sub HandleUserManagementLogin(app As UIApplication, payload As Object)
            Dim password = ReadText(payload, "password")
            Dim ok = KkyToolUserAccessService.VerifyAdminPassword(password)
            If ok Then
                SendToWeb("usermanagement:login-result", KkyToolUserAccessService.GetPublicSnapshot(app, True))
            Else
                SendToWeb("usermanagement:login-result", New With {
                    .authenticated = False,
                    .message = "관리자 비밀번호가 맞지 않습니다."
                })
            End If
        End Sub

        Private Sub HandleUserManagementSave(app As UIApplication, payload As Object)
            Try
                Dim enabled = ReadBool(payload, "enabled", True)
                Dim keywords = ReadStringList(payload, "allowedProfileKeywords")
                Dim users = ReadStringList(payload, "allowedUsers")
                Dim blockMessage = ReadText(payload, "blockMessage")
                Dim password = ReadText(payload, "password")
                Dim newPassword = ReadText(payload, "newPassword")

                KkyToolUserAccessService.SaveFromRequest(enabled, keywords, users, blockMessage, password, newPassword)
                SendToWeb("usermanagement:saved", KkyToolUserAccessService.GetPublicSnapshot(app, True))
            Catch ex As Exception
                SendToWeb("usermanagement:error", New With {.message = ex.Message})
            End Try
        End Sub

        Private Shared Function ReadText(payload As Object, prop As String) As String
            Try
                Dim raw = GetProp(payload, prop)
                If raw Is Nothing Then Return String.Empty
                Return Convert.ToString(raw).Trim()
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Function ReadBool(payload As Object, prop As String, fallback As Boolean) As Boolean
            Try
                Dim raw = GetProp(payload, prop)
                If raw Is Nothing Then Return fallback
                Return Convert.ToBoolean(raw)
            Catch
                Return fallback
            End Try
        End Function

        Private Shared Function ReadStringList(payload As Object, prop As String) As List(Of String)
            Dim result As New List(Of String)()

            Try
                Dim raw = GetProp(payload, prop)
                If raw Is Nothing Then Return result

                Dim text = TryCast(raw, String)
                If text IsNot Nothing Then
                    AddDelimited(result, text)
                    Return result
                End If

                Dim enumerable = TryCast(raw, IEnumerable)
                If enumerable Is Nothing Then Return result

                For Each item In enumerable
                    AddDelimited(result, Convert.ToString(item))
                Next
            Catch
            End Try

            Return result
        End Function

        Private Shared Sub AddDelimited(target As List(Of String), value As String)
            For Each part In If(value, String.Empty).Split(New Char() {","c, ";"c, vbCr(0), vbLf(0), vbTab(0)}, StringSplitOptions.RemoveEmptyEntries)
                Dim normalized = part.Trim()
                If normalized.Length = 0 Then Continue For
                If target.Exists(Function(x) String.Equals(x, normalized, StringComparison.OrdinalIgnoreCase)) Then Continue For
                target.Add(normalized)
            Next
        End Sub
    End Class
End Namespace
