using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace KKY_Tool_Revit.Services
{
    public static class LinkWorksetUiPrimeService
    {
        private enum PrimeState
        {
            Idle,
            OpenManageLinks,
            WaitForWindow,
            SelectRow,
            ClickReload,
            WaitAfterReload,
            CloseWindow,
            WaitForWindowClose,
            Completed,
            Faulted
        }

        private static readonly object SyncRoot = new object();
        private static UIApplication _uiapp;
        private static string _documentPath;
        private static string _documentTitle;
        private static int _targetLinkCount;
        private static int _currentLinkIndex;
        private static DateTime _nextActionUtc = DateTime.MinValue;
        private static DateTime _stateEnteredUtc = DateTime.MinValue;
        private static PrimeState _state = PrimeState.Idle;
        private static bool _running;
        private static bool _subscribed;
        private static string _lastMessage = string.Empty;
        private static Action<string> _log;
        private static Action<bool, string> _completed;
        private static RevitCommandId _manageLinksCommandId;

        private const int DialogAppearTimeoutSeconds = 15;
        private const int DialogCloseTimeoutSeconds = 8;
        private const int ReloadWaitMilliseconds = 4500;
        private const int AfterClickDelayMilliseconds = 500;
        private const int BetweenRowsDelayMilliseconds = 650;

        // Tuned from journaled UI interactions in Revit 2025.4.
        private const int BrowserFirstRowX = 164;
        private const int BrowserFirstRowY = 142;
        private const int BrowserRowStepY = 34;
        private const int BrowserReloadButtonX = 104;
        private const int BrowserReloadButtonY = 52;
        private const int BrowserCloseButtonX = 1219;
        private const int BrowserCloseButtonY = 2;

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const int SW_RESTORE = 9;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        public static bool IsRunning
        {
            get { lock (SyncRoot) { return _running; } }
        }

        public static bool Start(
            UIApplication uiapp,
            Document doc,
            string documentPath,
            string documentTitle,
            int targetLinkCount,
            Action<string> log,
            Action<bool, string> completed)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            lock (SyncRoot)
            {
                if (_running)
                {
                    log?.Invoke("링크 UI 프라이밍이 이미 실행 중입니다.");
                    return false;
                }

                _manageLinksCommandId = ResolveManageLinksCommandId();
                if (_manageLinksCommandId == null)
                {
                    log?.Invoke("Manage Links 명령을 찾지 못했습니다.");
                    return false;
                }

                _uiapp = uiapp;
                _documentPath = documentPath ?? string.Empty;
                _documentTitle = string.IsNullOrWhiteSpace(documentTitle) ? SafeTitle(doc) : documentTitle;
                _targetLinkCount = Math.Max(0, targetLinkCount);
                _currentLinkIndex = 0;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(250);
                _stateEnteredUtc = DateTime.UtcNow;
                _state = _targetLinkCount > 0 ? PrimeState.OpenManageLinks : PrimeState.Completed;
                _running = true;
                _lastMessage = "링크 UI 프라이밍 준비";
                _log = log;
                _completed = completed;

                EnsureSubscribedLocked();
                WriteLogLocked("링크 UI 프라이밍 시작 / 대상 링크 수 " + _targetLinkCount);
            }

            return true;
        }

        public static void Stop()
        {
            lock (SyncRoot)
            {
                StopLocked(false, _lastMessage);
            }
        }

        private static void EnsureSubscribedLocked()
        {
            if (_subscribed || _uiapp == null) return;
            _uiapp.Idling += OnIdling;
            _subscribed = true;
        }

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            lock (SyncRoot)
            {
                if (!_running) return;

                try
                {
                    e?.SetRaiseWithoutDelay();
                }
                catch
                {
                }

                if (DateTime.UtcNow < _nextActionUtc) return;

                try
                {
                    switch (_state)
                    {
                        case PrimeState.OpenManageLinks:
                            OpenManageLinksLocked();
                            break;
                        case PrimeState.WaitForWindow:
                            WaitForManageLinksWindowLocked();
                            break;
                        case PrimeState.SelectRow:
                            SelectCurrentRowLocked();
                            break;
                        case PrimeState.ClickReload:
                            ClickReloadLocked();
                            break;
                        case PrimeState.WaitAfterReload:
                            AdvanceAfterReloadLocked();
                            break;
                        case PrimeState.CloseWindow:
                            CloseManageLinksLocked();
                            break;
                        case PrimeState.WaitForWindowClose:
                            WaitForWindowCloseLocked();
                            break;
                        case PrimeState.Completed:
                            StopLocked(true, "링크 UI 프라이밍 완료");
                            break;
                        case PrimeState.Faulted:
                            StopLocked(false, _lastMessage);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    SetFaultedLocked("링크 UI 프라이밍 상태 처리 실패: " + ex.Message);
                }
            }
        }

        private static void OpenManageLinksLocked()
        {
            if (!IsTargetDocumentActiveLocked())
            {
                SetFaultedLocked("활성 문서가 변경되어 Manage Links를 열 수 없습니다.");
                return;
            }

            _uiapp.PostCommand(_manageLinksCommandId);
            WriteLogLocked("Manage Links 열기");
            TransitionLocked(PrimeState.WaitForWindow, 800);
        }

        private static void WaitForManageLinksWindowLocked()
        {
            IntPtr handle = FindManageLinksWindow();
            if (handle != IntPtr.Zero)
            {
                WriteLogLocked("Manage Links 창 확인");
                TransitionLocked(PrimeState.SelectRow, 500);
                return;
            }

            if (DateTime.UtcNow - _stateEnteredUtc > TimeSpan.FromSeconds(DialogAppearTimeoutSeconds))
            {
                SetFaultedLocked("Manage Links 창이 나타나지 않았습니다.");
                return;
            }

            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(250);
        }

        private static void SelectCurrentRowLocked()
        {
            IntPtr handle = FindManageLinksWindow();
            if (handle == IntPtr.Zero)
            {
                SetFaultedLocked("Manage Links 창을 찾지 못했습니다.");
                return;
            }

            FocusWindow(handle);
            int rowY = BrowserFirstRowY + (_currentLinkIndex * BrowserRowStepY);
            ClickClientPoint(handle, BrowserFirstRowX, rowY);
            WriteLogLocked("Manage Links 행 선택: " + (_currentLinkIndex + 1) + "/" + _targetLinkCount);
            TransitionLocked(PrimeState.ClickReload, AfterClickDelayMilliseconds);
        }

        private static void ClickReloadLocked()
        {
            IntPtr handle = FindManageLinksWindow();
            if (handle == IntPtr.Zero)
            {
                SetFaultedLocked("Reload 전에 Manage Links 창을 찾지 못했습니다.");
                return;
            }

            FocusWindow(handle);
            ClickClientPoint(handle, BrowserReloadButtonX, BrowserReloadButtonY);
            WriteLogLocked("Manage Links Reload 클릭: " + (_currentLinkIndex + 1) + "/" + _targetLinkCount);
            TransitionLocked(PrimeState.WaitAfterReload, ReloadWaitMilliseconds);
        }

        private static void AdvanceAfterReloadLocked()
        {
            _currentLinkIndex++;
            if (_currentLinkIndex >= _targetLinkCount)
            {
                TransitionLocked(PrimeState.CloseWindow, 350);
                return;
            }

            TransitionLocked(PrimeState.SelectRow, BetweenRowsDelayMilliseconds);
        }

        private static void CloseManageLinksLocked()
        {
            IntPtr handle = FindManageLinksWindow();
            if (handle == IntPtr.Zero)
            {
                TransitionLocked(PrimeState.Completed, 150);
                return;
            }

            FocusWindow(handle);
            TryCloseWindow(handle);
            ClickClientPoint(handle, BrowserCloseButtonX, BrowserCloseButtonY);
            WriteLogLocked("Manage Links 닫기");
            TransitionLocked(PrimeState.WaitForWindowClose, 500);
        }

        private static void WaitForWindowCloseLocked()
        {
            IntPtr handle = FindManageLinksWindow();
            if (handle == IntPtr.Zero)
            {
                TransitionLocked(PrimeState.Completed, 100);
                return;
            }

            if (DateTime.UtcNow - _stateEnteredUtc > TimeSpan.FromSeconds(DialogCloseTimeoutSeconds))
            {
                SetFaultedLocked("Manage Links 창이 닫히지 않았습니다.");
                return;
            }

            TryCloseWindow(handle);
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
        }

        private static void TransitionLocked(PrimeState nextState, int delayMilliseconds)
        {
            _state = nextState;
            _stateEnteredUtc = DateTime.UtcNow;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(delayMilliseconds);
        }

        private static void SetFaultedLocked(string message)
        {
            _lastMessage = message ?? string.Empty;
            _state = PrimeState.Faulted;
            _stateEnteredUtc = DateTime.UtcNow;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(100);
            WriteLogLocked(_lastMessage);
        }

        private static void StopLocked(bool succeeded, string message)
        {
            Action<bool, string> completed = null;

            if (_subscribed && _uiapp != null)
            {
                _uiapp.Idling -= OnIdling;
            }

            _subscribed = false;
            _running = false;
            _state = PrimeState.Idle;
            _nextActionUtc = DateTime.MinValue;
            _stateEnteredUtc = DateTime.MinValue;
            _currentLinkIndex = 0;
            _targetLinkCount = 0;
            _documentPath = string.Empty;
            _documentTitle = string.Empty;
            _manageLinksCommandId = null;
            _uiapp = null;
            _lastMessage = message ?? string.Empty;

            completed = _completed;
            _completed = null;
            _log = null;

            try
            {
                completed?.Invoke(succeeded, message ?? string.Empty);
            }
            catch
            {
            }
        }

        private static bool IsTargetDocumentActiveLocked()
        {
            if (_uiapp == null) return false;

            Document activeDoc = null;
            try
            {
                activeDoc = _uiapp.ActiveUIDocument != null ? _uiapp.ActiveUIDocument.Document : null;
            }
            catch
            {
                activeDoc = null;
            }

            if (activeDoc == null) return false;

            string activePath = SafePath(activeDoc);
            if (!string.IsNullOrWhiteSpace(_documentPath) &&
                !string.IsNullOrWhiteSpace(activePath) &&
                string.Equals(activePath, _documentPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string activeTitle = SafeTitle(activeDoc);
            return !string.IsNullOrWhiteSpace(_documentTitle) &&
                   string.Equals(activeTitle, _documentTitle, StringComparison.OrdinalIgnoreCase);
        }

        private static string SafePath(Document doc)
        {
            try { return doc != null ? doc.PathName : string.Empty; } catch { return string.Empty; }
        }

        private static string SafeTitle(Document doc)
        {
            try { return doc != null ? doc.Title : string.Empty; } catch { return string.Empty; }
        }

        private static RevitCommandId ResolveManageLinksCommandId()
        {
            try
            {
                return RevitCommandId.LookupCommandId("ID_LINKED_DWG");
            }
            catch
            {
                return null;
            }
        }

        private static void WriteLogLocked(string message)
        {
            _lastMessage = message ?? string.Empty;
            _log?.Invoke(_lastMessage);
        }

        private static IntPtr FindManageLinksWindow()
        {
            IntPtr found = IntPtr.Zero;
            int currentProcessId = Process.GetCurrentProcess().Id;

            EnumWindows((hWnd, lParam) =>
            {
                if (found != IntPtr.Zero) return false;
                if (!IsWindowVisible(hWnd)) return true;

                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);
                if (processId != currentProcessId) return true;

                string title = GetWindowTextSafe(hWnd);
                if (string.IsNullOrWhiteSpace(title)) return true;
                if (title.IndexOf("Manage Links", StringComparison.OrdinalIgnoreCase) < 0) return true;

                found = hWnd;
                return false;
            }, IntPtr.Zero);

            return found;
        }

        private static void FocusWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            ShowWindow(hWnd, SW_RESTORE);
            BringWindowToTop(hWnd);
            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            SetForegroundWindow(hWnd);
        }

        private static void TryCloseWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            try
            {
                PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE);
            }
            catch
            {
            }
        }

        private static void ClickClientPoint(IntPtr hWnd, int clientX, int clientY)
        {
            if (hWnd == IntPtr.Zero) return;

            RECT rect;
            if (!GetWindowRect(hWnd, out rect)) return;

            int screenX = rect.Left + clientX;
            int screenY = rect.Top + clientY;
            SetCursorPos(screenX, screenY);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        private static string GetWindowTextSafe(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length <= 0) return string.Empty;
            var buffer = new System.Text.StringBuilder(length + 1);
            GetWindowText(hWnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        private const int WM_CLOSE = 0x0010;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    }
}
