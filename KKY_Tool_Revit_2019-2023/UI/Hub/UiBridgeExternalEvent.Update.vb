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

        Private Sub HandleUpdateCheckLegacy()
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
                            .hasUpdate = info.HasUpdate,
                            .latestVersion = info.LatestVersion,
                            .currentVersionDisplay = info.CurrentVersionDisplay,
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

        Private Sub HandleUpdateCheck(payload As Object)
            Dim silent As Boolean = False
            Dim startup As Boolean = False

            Dim silentRaw = GetProp(payload, "silent")
            If silentRaw IsNot Nothing Then
                silent = String.Equals(Convert.ToString(silentRaw), "true", StringComparison.OrdinalIgnoreCase) OrElse
                         String.Equals(Convert.ToString(silentRaw), "1", StringComparison.OrdinalIgnoreCase)
            End If

            Dim startupRaw = GetProp(payload, "startup")
            If startupRaw IsNot Nothing Then
                startup = String.Equals(Convert.ToString(startupRaw), "true", StringComparison.OrdinalIgnoreCase) OrElse
                          String.Equals(Convert.ToString(startupRaw), "1", StringComparison.OrdinalIgnoreCase)
            End If

            If Interlocked.CompareExchange(_updateBusy, 1, 0) <> 0 Then
                SendToWeb("host:update-state", New With {
                    .busy = True,
                    .showToast = Not silent,
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
                            .showToast = Not silent,
                            .hasUpdate = info.HasUpdate,
                            .latestVersion = info.LatestVersion,
                            .currentVersionDisplay = info.CurrentVersionDisplay,
                            .startupNotice = silent AndAlso startup AndAlso info.HasUpdate,
                            .kind = If(info.HasUpdate, "ok", "info"),
                            .message = info.Message
                        })
                    Catch ex As Exception
                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = Not silent,
                            .kind = "err",
                            .message = "업데이트 확인에 실패했습니다: " & ex.Message
                        })
                    Finally
                        Interlocked.Exchange(_updateBusy, 0)
                    End Try
                End Sub)
        End Sub

        Private Sub SendUpdateDownloadProgress(percent As Integer, message As String)
            SendToWeb("host:update-state", New With {
                .busy = True,
                .showToast = True,
                .kind = "info",
                .phase = "download",
                .progressPercent = percent,
                .progressMessage = message
            })
        End Sub

        Private Sub HandleUpdateInstallLegacy(payload As Object)
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
                            Throw New InvalidOperationException("새 버전 정보는 확인됐지만 업데이트 파일 주소가 없습니다.")
                        End If

                        Dim installerPath = AddinUpdateService.PrepareInstaller(info)
                        Dim scriptPath = AddinUpdateService.QueueInstallerAfterProcessExit(installerPath)

                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = "ok",
                            .message = "업데이트 파일 준비가 완료되었습니다. Revit을 종료하면 업데이트가 자동으로 시작됩니다.",
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
                .phase = "download",
                .progressPercent = 0,
                .progressMessage = "최신 업데이트 패키지를 준비하는 중입니다."
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
                            Throw New InvalidOperationException("새 버전 정보는 확인됐지만 업데이트 파일 주소가 없습니다.")
                        End If

                        Dim installerPath = AddinUpdateService.PrepareInstaller(
                            info,
                            Sub(percent, progressMessage)
                                SendUpdateDownloadProgress(percent, progressMessage)
                            End Sub)

                        Dim scriptPath = AddinUpdateService.QueueInstallerAfterProcessExit(installerPath)

                        SendToWeb("host:update-state", New With {
                            .busy = False,
                            .showToast = True,
                            .kind = "ok",
                            .phase = "ready",
                            .message = "업데이트 패키지 다운로드가 완료되었습니다. 현재 실행 중인 Revit을 모두 종료하면 업데이트가 자동 실행됩니다. Revit을 다시 실행해 주세요.",
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
