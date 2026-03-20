Option Explicit On
Option Strict On

Imports System
Imports System.Threading

Namespace UI.Hub

    Partial Public Class UiBridgeExternalEvent

        Private Shared _updateBusy As Integer = 0
        Private Shared _lastUpdateInfo As AddinUpdateService.UpdateInfo = Nothing

        Private Sub HandleUpdateQuery()
            Dim info = AddinUpdateService.CreateInitialInfo()
            _lastUpdateInfo = info
            SendToWeb("host:update-info", info)
        End Sub

        Private Sub HandleUpdateCheck()
            If Interlocked.CompareExchange(_updateBusy, 1, 0) <> 0 Then
                SendToWeb("host:update-state", New With {
                    .busy = True,
                    .showToast = True,
                    .kind = "info",
                    .message = "이미 업데이트 확인이 진행 중입니다."
                })
                Return
            End If

            SendToWeb("host:update-state", New With {
                .busy = True,
                .showToast = False,
                .kind = "info",
                .message = "업데이트를 확인하는 중입니다."
            })

            ThreadPool.QueueUserWorkItem(
                Sub(state)
                    Try
                        Dim info = AddinUpdateService.CheckForUpdates()
                        _lastUpdateInfo = info
                        SendToWeb("host:update-info", info)
                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = If(info.HasUpdate, "ok", "info"),
                            .message = info.Message
                        })
                    Catch ex As Exception
                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = "err",
                            .message = "업데이트 확인에 실패했습니다: " & ex.Message
                        })
                    Finally
                        Interlocked.Exchange(_updateBusy, 0)
                    End Try
                End Sub)
        End Sub

        Private Sub HandleUpdateInstall(payload As Object)
            If Interlocked.CompareExchange(_updateBusy, 1, 0) <> 0 Then
                SendToWeb("host:update-state", New With {
                    .busy = True,
                    .showToast = True,
                    .kind = "info",
                    .message = "이미 다른 업데이트 작업이 진행 중입니다."
                })
                Return
            End If

            SendToWeb("host:update-state", New With {
                .busy = True,
                .showToast = False,
                .kind = "info",
                .message = "업데이트 설치를 준비하는 중입니다."
            })

            ThreadPool.QueueUserWorkItem(
                Sub(state)
                    Try
                        Dim info = _lastUpdateInfo
                        If info Is Nothing OrElse Not info.HasUpdate Then
                            info = AddinUpdateService.CheckForUpdates()
                            _lastUpdateInfo = info
                        End If

                        If info Is Nothing OrElse Not info.HasUpdate Then
                            Throw New InvalidOperationException("설치할 새 버전이 없습니다.")
                        End If

                        If Not info.CanInstall Then
                            Throw New InvalidOperationException("새 버전 정보는 확인됐지만 설치파일 주소가 없습니다.")
                        End If

                        Dim installerPath = AddinUpdateService.PrepareInstaller(info)
                        Dim scriptPath = AddinUpdateService.QueueInstallerAfterProcessExit(installerPath)

                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = "ok",
                            .message = "설치파일 준비가 완료되었습니다. Revit을 종료하면 업데이트가 자동으로 시작됩니다.",
                            .installerPath = installerPath,
                            .scriptPath = scriptPath
                        })
                    Catch ex As Exception
                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = "err",
                            .message = "업데이트 설치 준비에 실패했습니다: " & ex.Message
                        })
                    Finally
                        Interlocked.Exchange(_updateBusy, 0)
                    End Try
                End Sub)
        End Sub

    End Class

End Namespace
