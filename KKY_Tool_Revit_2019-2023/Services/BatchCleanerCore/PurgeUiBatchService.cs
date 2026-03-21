using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class PurgeUiBatchService
    {
        private enum PurgeBatchState
        {
            Idle,
            OpenNextDocument,
            ActivateCurrentDocument,
            WaitBeforePostingCommand,
            WaitingForDialog,
            WaitingForDialogToClose,
            SaveCurrentDocument,
            ReturnToAnchor,
            CloseCurrentDocument,
            Completed,
            Faulted
        }

        private sealed class PurgeExternalEventHandler : IExternalEventHandler
        {
            public string PendingOpenPath { get; set; }

            public void Execute(UIApplication app)
            {
                string path = PendingOpenPath;
                PendingOpenPath = null;
                if (string.IsNullOrWhiteSpace(path)) return;
                PurgeUiBatchService.HandleOpenAndActivateFromExternalEvent(app, path);
            }

            public string GetName()
            {
                return "KKY RVT Cleaner Purge External Event";
            }
        }

        private static readonly object SyncRoot = new object();
        private static UIApplication _uiapp;
        private static BatchPrepareSession _session;
        private static List<string> _remainingPaths = new List<string>();
        private static string _currentPath;
        private static Document _currentDocument;
        private static string _anchorPath;
        private static string _anchorTitle;
        private static int _initialFileCount;
        private static int _maxIterations;
        private static int _currentIteration;
        private static int _processedFileCount;
        private static int _docSwitchAttempts;
        private static DateTime _nextActionUtc = DateTime.MinValue;
        private static DateTime _commandPostedUtc = DateTime.MinValue;
        private static DateTime _lastEnterUtc = DateTime.MinValue;
        private static int _dialogEnterAttempts;
        private static Timer _dialogWatcherTimer;
        private static bool _subscribedToIdling;
        private static bool _running;
        private static bool _completedSuccessfully;
        private static bool _faulted;
        private static PurgeBatchState _state = PurgeBatchState.Idle;
        private static string _lastStatusMessage = string.Empty;
        private static IntPtr _mainWindowHandle = IntPtr.Zero;
        private static bool _mainWindowTopMost;
        private static DateTime _lastRevitRefocusUtc = DateTime.MinValue;
        private static RevitCommandId _purgeCommandId;
        private static ExternalEvent _externalEvent;
        private static PurgeExternalEventHandler _externalEventHandler;

        private const int DefaultIterations = 5;
        private const int DialogAppearTimeoutSeconds = 15;
        private const int DialogCloseTimeoutSeconds = 20;
        private const int MaxEnterAttemptsPerIteration = 3;
        private const int MaxDocSwitchAttempts = 12;
        private const int RevitRefocusCooldownMilliseconds = 400;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const byte VK_RETURN = 0x0D;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_TAB = 0x09;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;

        public static bool IsRunning
        {
            get { lock (SyncRoot) { return _running; } }
        }

        public static bool Start(UIApplication uiapp, BatchPrepareSession session, int iterations, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            if (session == null) throw new ArgumentNullException(nameof(session));

            lock (SyncRoot)
            {
                if (_running)
                {
                    log?.Invoke("Purge 일괄처리가 이미 실행 중입니다.");
                    return false;
                }

                List<string> targetPaths = (session.CleanedOutputPaths ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (targetPaths.Count == 0)
                {
                    log?.Invoke("Purge 대상 저장 파일이 없습니다.");
                    return false;
                }

                Document activeDoc = GetActiveDocument(uiapp);
                if (activeDoc == null)
                {
                    log?.Invoke("Purge를 시작하려면 기준이 되는 활성 Revit 문서가 하나 열려 있어야 합니다.");
                    return false;
                }

                EnsureExternalEventLocked();

                _uiapp = uiapp;
                _session = session;
                _remainingPaths = targetPaths;
                _initialFileCount = targetPaths.Count;
                _maxIterations = iterations > 0 ? iterations : DefaultIterations;
                _currentIteration = 1;
                _processedFileCount = 0;
                _currentPath = null;
                _currentDocument = null;
                _anchorPath = SafePath(activeDoc);
                _anchorTitle = SafeTitle(activeDoc);
                _state = PurgeBatchState.OpenNextDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
                _commandPostedUtc = DateTime.MinValue;
                _lastEnterUtc = DateTime.MinValue;
                _dialogEnterAttempts = 0;
                _docSwitchAttempts = 0;
                _lastStatusMessage = "Purge 시작 준비";
                _completedSuccessfully = false;
                _faulted = false;
                _mainWindowHandle = GetMainWindowHandle(uiapp);
                _mainWindowTopMost = false;
                _lastRevitRefocusUtc = DateTime.MinValue;
                _purgeCommandId = ResolvePurgeCommandId();
                _running = true;
                if (_session != null)
                {
                    _session.PurgeCountComparisons = new List<ModelObjectCountComparison>();
                    _session.PurgeCountComparisonXlsxPath = null;
                }

                if (_purgeCommandId == null)
                {
                    WriteLog("Purge Unused 명령 ID를 찾지 못했습니다. 자동 퍼지를 시작할 수 없습니다.");
                    StopLocked(false);
                    return false;
                }

                EnsureSubscribedToIdlingLocked();
                StartDialogWatcherLocked();
            }

            log?.Invoke("Purge 일괄처리 시작 - 저장된 정리 결과 파일을 하나씩 다시 열어 순차 실행합니다.");
            WriteLog("Purge batch start / target files " + GetRemainingFileCount() + " / iterations " + (_maxIterations > 0 ? _maxIterations : DefaultIterations));
            return true;
        }

        public static void Stop()
        {
            lock (SyncRoot)
            {
                StopLocked(false);
            }
        }

        public static PurgeBatchProgressSnapshot GetProgressSnapshot()
        {
            lock (SyncRoot)
            {
                return new PurgeBatchProgressSnapshot
                {
                    IsRunning = _running,
                    IsCompleted = _completedSuccessfully,
                    IsFaulted = _faulted,
                    TotalFiles = _initialFileCount,
                    CurrentFileIndex = _currentPath != null ? Math.Min(_processedFileCount + 1, _initialFileCount) : Math.Min(_processedFileCount, _initialFileCount),
                    CurrentIteration = _currentIteration,
                    TotalIterations = _maxIterations > 0 ? _maxIterations : DefaultIterations,
                    CurrentFileName = _currentPath != null ? Path.GetFileName(_currentPath) : string.Empty,
                    StateName = _state.ToString(),
                    Message = _lastStatusMessage
                };
            }
        }

        internal static void HandleOpenAndActivateFromExternalEvent(UIApplication app, string targetPath)
        {
            if (app == null || string.IsNullOrWhiteSpace(targetPath)) return;

            try
            {
                Document activeDoc = GetActiveDocument(app);
                if (IsMatchingDocument(activeDoc, targetPath, Path.GetFileNameWithoutExtension(targetPath)))
                {
                    lock (SyncRoot)
                    {
                        _currentDocument = activeDoc;
                        _lastStatusMessage = "퍼지 대상 활성 열기 완료";
                    }
                    WriteLog("퍼지 대상 활성 열기 완료(이미 활성): " + targetPath);
                    return;
                }

                UIDocument activated = app.OpenAndActivateDocument(targetPath);
                lock (SyncRoot)
                {
                    _currentDocument = activated != null ? activated.Document : null;
                    _lastStatusMessage = "퍼지 대상 활성 열기 완료";
                }
                WriteLog("퍼지 대상 활성 열기 완료: " + targetPath);
            }
            catch (Exception ex)
            {
                lock (SyncRoot)
                {
                    _lastStatusMessage = "퍼지 대상 활성 열기 실패: " + ex.Message;
                    _faulted = true;
                    _state = PurgeBatchState.Faulted;
                    _nextActionUtc = DateTime.UtcNow.AddMilliseconds(100);
                }
                WriteLog("퍼지 대상 활성 열기 실패: " + targetPath + " / " + ex.Message);
            }
        }

        private static void EnsureExternalEventLocked()
        {
            if (_externalEventHandler != null && _externalEvent != null) return;
            _externalEventHandler = new PurgeExternalEventHandler();
            _externalEvent = ExternalEvent.Create(_externalEventHandler);
        }

        private static int GetRemainingFileCount()
        {
            lock (SyncRoot)
            {
                return _remainingPaths != null ? _remainingPaths.Count : 0;
            }
        }

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            lock (SyncRoot)
            {
                if (!_running) return;
                try
                {
                    if (e != null)
                    {
                        e.SetRaiseWithoutDelay();
                    }
                }
                catch
                {
                }
                if (DateTime.UtcNow < _nextActionUtc) return;

                try
                {
                    EnsureRevitTopMostLocked();
                    switch (_state)
                    {
                        case PurgeBatchState.OpenNextDocument:
                            OpenNextDocumentLocked();
                            break;
                        case PurgeBatchState.ActivateCurrentDocument:
                            ActivateCurrentDocumentLocked();
                            break;
                        case PurgeBatchState.WaitBeforePostingCommand:
                            PostPurgeCommandLocked();
                            break;
                        case PurgeBatchState.SaveCurrentDocument:
                            SaveCurrentDocumentLocked();
                            break;
                        case PurgeBatchState.ReturnToAnchor:
                            ReturnToAnchorLocked();
                            break;
                        case PurgeBatchState.CloseCurrentDocument:
                            CloseCurrentDocumentLocked();
                            break;
                        case PurgeBatchState.Completed:
                            WriteLog("Purge 일괄처리 완료");
                            StopLocked(true);
                            break;
                        case PurgeBatchState.Faulted:
                            StopLocked(false);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("Purge ?곹깭 泥섎━ ?ㅻ쪟: " + ex.Message);
                    SetFaultedLocked();
                }
            }
        }

        private static void OpenNextDocumentLocked()
        {
            if (_remainingPaths == null || _remainingPaths.Count == 0)
            {
                _state = PurgeBatchState.Completed;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
                return;
            }

            _currentPath = _remainingPaths[0];
            _currentDocument = FindOpenDocumentByPath(_currentPath);
            _currentIteration = 1;
            _docSwitchAttempts = 0;
            RefocusRevitForPurgeLocked("Prepare next purge file");

            if (IsCurrentDocumentActiveLocked())
            {
                EnsureCurrentFileBeforeCountLocked();
                WriteLog("퍼지 대상 활성 문서 확인: " + Path.GetFileName(_currentPath));
                _state = PurgeBatchState.WaitBeforePostingCommand;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
                return;
            }

            if (_currentDocument != null && _currentDocument.IsValidObject)
            {
                SendCtrlTab();
                WriteLog("?대? ?대┛ ?쇱? ???臾몄꽌 ?쒖꽦???쒕룄(Ctrl+Tab): " + Path.GetFileName(_currentPath));
                _state = PurgeBatchState.ActivateCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(700);
                return;
            }

            if (_externalEvent == null || _externalEventHandler == null)
            {
                WriteLog("?쇱? ?몃? ?대깽?멸? 以鍮꾨릺吏 ?딆븯?듬땲??");
                SetFaultedLocked();
                return;
            }

            _externalEventHandler.PendingOpenPath = _currentPath;
            RefocusRevitForPurgeLocked("퍼지 대상 활성 열기 요청");
            _externalEvent.Raise();
            WriteLog("?쇱? ????뚯씪 ?쒖꽦 ?닿린 ?붿껌: " + _currentPath);
            _state = PurgeBatchState.ActivateCurrentDocument;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(900);
        }

        private static void ActivateCurrentDocumentLocked()
        {
            if (IsCurrentDocumentActiveLocked())
            {
                EnsureCurrentFileBeforeCountLocked();
                WriteLog("퍼지 대상 활성 열기 완료: " + Path.GetFileName(_currentPath));
                _state = PurgeBatchState.WaitBeforePostingCommand;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
                return;
            }

            _docSwitchAttempts++;
            if (_docSwitchAttempts % 3 == 0 && _externalEvent != null && _externalEventHandler != null && !string.IsNullOrWhiteSpace(_currentPath))
            {
                _externalEventHandler.PendingOpenPath = _currentPath;
                RefocusRevitForPurgeLocked("Request reopen current purge file");
                _externalEvent.Raise();
            }

            if (_docSwitchAttempts >= MaxDocSwitchAttempts)
            {
                WriteLog("?쇱? ????뚯씪???쒖꽦?뷀븯吏 紐삵빐 嫄대꼫?곷땲?? " + (_currentPath ?? string.Empty));
                SkipCurrentDocumentLocked(false);
                return;
            }

            RefocusRevitForPurgeLocked("Retry activate current purge file");
            SendCtrlTab();
            WriteLog("퍼지 대상 활성 재시도(Ctrl+Tab): " + _docSwitchAttempts + "/" + MaxDocSwitchAttempts);
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(650);
        }

        private static void PostPurgeCommandLocked()
        {
            if (_purgeCommandId == null)
            {
                SetFaultedLocked();
                return;
            }

            if (!IsCurrentDocumentActiveLocked())
            {
                _state = PurgeBatchState.ActivateCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
                return;
            }

            RefocusRevitForPurgeLocked("Purge 紐낅졊 ?ㅽ뻾");
            WriteLog("Purge ?ㅽ뻾: " + Path.GetFileName(_currentPath) + " / 諛섎났 " + _currentIteration + "/" + _maxIterations);
            _uiapp.PostCommand(_purgeCommandId);
            _commandPostedUtc = DateTime.UtcNow;
            _dialogEnterAttempts = 0;
            _lastEnterUtc = DateTime.MinValue;
            _state = PurgeBatchState.WaitingForDialog;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(250);
        }

        private static void AdvanceIterationLocked()
        {
            if (_currentIteration < _maxIterations)
            {
                _currentIteration++;
                _state = PurgeBatchState.WaitBeforePostingCommand;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
                return;
            }

            _state = PurgeBatchState.SaveCurrentDocument;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
        }

        private static void SaveCurrentDocumentLocked()
        {
            Document doc = GetCurrentDocumentLocked();
            if (doc == null)
            {
                WriteLog("??????臾몄꽌瑜?李얠? 紐삵뻽?듬땲?? " + (_currentPath ?? string.Empty));
                SkipCurrentDocumentLocked(false);
                return;
            }

            WriteLog("Purge ??????쒖옉: " + _currentPath);
            doc.Save();
            DeleteBackupFiles(_currentPath);
            CaptureCurrentFileAfterCountLocked(doc);
            WriteLog("Purge ??????꾨즺: " + _currentPath);

            RefocusRevitForPurgeLocked("Return from purge save");

            if (IsAnchorActiveLocked())
            {
                _state = PurgeBatchState.CloseCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(400);
                return;
            }

            _docSwitchAttempts = 0;
            RefocusRevitForPurgeLocked("湲곗? 臾몄꽌 蹂듦?");
            SendCtrlTab();
            WriteLog("湲곗? 臾몄꽌濡?蹂듦? ?쒕룄(Ctrl+Tab)");
            _state = PurgeBatchState.ReturnToAnchor;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(650);
        }

        private static void ReturnToAnchorLocked()
        {
            RefocusRevitForPurgeLocked("Return to anchor after purge");

            if (IsAnchorActiveLocked())
            {
                WriteLog("湲곗? 臾몄꽌 蹂듦? ?꾨즺");
                _state = PurgeBatchState.CloseCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(400);
                return;
            }

            _docSwitchAttempts++;
            if (_docSwitchAttempts % 3 == 0 && _externalEvent != null && _externalEventHandler != null && !string.IsNullOrWhiteSpace(_currentPath))
            {
                _externalEventHandler.PendingOpenPath = _currentPath;
                RefocusRevitForPurgeLocked("Request reopen current purge file");
                _externalEvent.Raise();
            }

            if (_docSwitchAttempts >= MaxDocSwitchAttempts)
            {
                WriteLog("湲곗? 臾몄꽌濡?蹂듦??섏? 紐삵빐 ?먮룞 ?쇱?瑜?以묐떒?⑸땲??");
                SetFaultedLocked();
                return;
            }

            RefocusRevitForPurgeLocked("Retry return to anchor document");
            SendCtrlTab();
            WriteLog("湲곗? 臾몄꽌 蹂듦? ?ъ떆??Ctrl+Tab): " + _docSwitchAttempts + "/" + MaxDocSwitchAttempts);
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(650);
        }

        private static void CloseCurrentDocumentLocked()
        {
            Document doc = GetCurrentDocumentLocked();
            if (doc != null && doc.IsValidObject)
            {
                if (IsCurrentDocumentActiveLocked())
                {
                    WriteLog("?쇱? ???臾몄꽌媛 ?꾩쭅 ?쒖꽦 ?곹깭?ъ꽌 ?レ? 紐삵뻽?듬땲??");
                    SetFaultedLocked();
                    return;
                }

                doc.Close(false);
                WriteLog("Purge ?꾨즺 臾몄꽌 ?リ린: " + (_currentPath ?? string.Empty));
            }

            if (_remainingPaths != null && _remainingPaths.Count > 0)
            {
                _remainingPaths.RemoveAt(0);
            }
            _processedFileCount++;
            _currentPath = null;
            _currentDocument = null;
            _currentIteration = 1;
            _state = PurgeBatchState.OpenNextDocument;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(350);
        }

        private static void SkipCurrentDocumentLocked(bool closeIfPossible)
        {
            MarkCurrentFileCountFailedLocked(_lastStatusMessage);
            try
            {
                Document doc = GetCurrentDocumentLocked();
                if (closeIfPossible && doc != null && doc.IsValidObject && !IsCurrentDocumentActiveLocked())
                {
                    doc.Close(false);
                }
            }
            catch
            {
            }

            if (_remainingPaths != null && _remainingPaths.Count > 0)
            {
                _remainingPaths.RemoveAt(0);
            }
            _processedFileCount++;
            _currentPath = null;
            _currentDocument = null;
            _currentIteration = 1;
            _state = PurgeBatchState.OpenNextDocument;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
        }

        private static void EnsureCurrentFileBeforeCountLocked()
        {
            if (_session == null || string.IsNullOrWhiteSpace(_currentPath)) return;
            if (_session.PurgeCountComparisons == null)
            {
                _session.PurgeCountComparisons = new List<ModelObjectCountComparison>();
            }

            ModelObjectCountComparison existing = _session.PurgeCountComparisons
                .FirstOrDefault(x => x != null && string.Equals(x.SourcePath, _currentPath, StringComparison.OrdinalIgnoreCase));
            if (existing != null && existing.BeforeCount.HasValue) return;

            Document doc = GetCurrentDocumentLocked();
            if (doc == null || !doc.IsValidObject) return;

            int beforeCount = 0;
            try
            {
                beforeCount = ModelParameterExtractionService.CountExtractableElements(doc);
            }
            catch (Exception ex)
            {
                WriteLog("Purge 객체수 사전 집계 실패: " + Path.GetFileName(_currentPath) + " / " + ex.Message);
                return;
            }

            if (existing == null)
            {
                existing = new ModelObjectCountComparison
                {
                    FileName = Path.GetFileName(_currentPath),
                    SourcePath = _currentPath,
                    OutputPath = _currentPath,
                    Status = string.Empty,
                    Note = string.Empty
                };
                _session.PurgeCountComparisons.Add(existing);
            }

            existing.BeforeCount = beforeCount;
        }

        private static void CaptureCurrentFileAfterCountLocked(Document doc)
        {
            if (_session == null || doc == null || !doc.IsValidObject || string.IsNullOrWhiteSpace(_currentPath)) return;
            EnsureCurrentFileBeforeCountLocked();

            if (_session.PurgeCountComparisons == null)
            {
                _session.PurgeCountComparisons = new List<ModelObjectCountComparison>();
            }

            ModelObjectCountComparison existing = _session.PurgeCountComparisons
                .FirstOrDefault(x => x != null && string.Equals(x.SourcePath, _currentPath, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                existing = new ModelObjectCountComparison
                {
                    FileName = Path.GetFileName(_currentPath),
                    SourcePath = _currentPath,
                    OutputPath = _currentPath
                };
                _session.PurgeCountComparisons.Add(existing);
            }

            try
            {
                existing.AfterCount = ModelParameterExtractionService.CountExtractableElements(doc);
                existing.Status = "O";
                existing.Note = "Purge 완료";
                WriteLog("Purge 객체수 비교: " + existing.FileName + " / 전 " + (existing.BeforeCount.HasValue ? existing.BeforeCount.Value.ToString() : "-") + " / 후 " + existing.AfterCount.Value);
            }
            catch (Exception ex)
            {
                existing.Status = "X";
                existing.Note = ex.Message;
                WriteLog("Purge 객체수 사후 집계 실패: " + existing.FileName + " / " + ex.Message);
            }
        }

        private static void MarkCurrentFileCountFailedLocked(string note)
        {
            if (_session == null || string.IsNullOrWhiteSpace(_currentPath)) return;
            if (_session.PurgeCountComparisons == null)
            {
                _session.PurgeCountComparisons = new List<ModelObjectCountComparison>();
            }

            ModelObjectCountComparison existing = _session.PurgeCountComparisons
                .FirstOrDefault(x => x != null && string.Equals(x.SourcePath, _currentPath, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                existing = new ModelObjectCountComparison
                {
                    FileName = Path.GetFileName(_currentPath),
                    SourcePath = _currentPath,
                    OutputPath = _currentPath
                };
                _session.PurgeCountComparisons.Add(existing);
            }

            existing.Status = "X";
            if (string.IsNullOrWhiteSpace(existing.Note))
            {
                existing.Note = note ?? string.Empty;
            }
        }

        private static void OnDialogWatcherTick(object state)
        {
            lock (SyncRoot)
            {
                if (!_running) return;
                if (_state == PurgeBatchState.WaitingForDialog)
                {
                    HandleWaitingForDialogLocked();
                }
                else if (_state == PurgeBatchState.WaitingForDialogToClose)
                {
                    HandleWaitingForDialogToCloseLocked();
                }
            }
        }

        private static void HandleWaitingForDialogLocked()
        {
            if (!IsCurrentDocumentActiveLocked())
            {
                WriteLog("Active document changed while waiting for purge dialog; restoring target document.");
                _state = PurgeBatchState.ActivateCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
                return;
            }

            IntPtr dialogHandle = FindRevitModalDialog(_mainWindowHandle);
            if (dialogHandle != IntPtr.Zero)
            {
                RefocusRevitForPurgeLocked("Purge ??붿긽???뺤씤");
                EnsureWindowTopMost(dialogHandle);
                FocusWindow(dialogHandle);
                SendEnterToWindow(dialogHandle);
                _dialogEnterAttempts = 1;
                _lastEnterUtc = DateTime.UtcNow;
                _state = PurgeBatchState.WaitingForDialogToClose;
                WriteLog("Purge ??붿긽???뺤씤 / Enter ?꾩넚");
                return;
            }

            RefocusRevitForPurgeLocked("Wait for purge dialog");

            if (_commandPostedUtc != DateTime.MinValue && DateTime.UtcNow - _commandPostedUtc > TimeSpan.FromSeconds(DialogAppearTimeoutSeconds))
            {
                WriteLog("Purge ??붿긽?먭? ?섑??섏? ?딆븘 ?ㅼ쓬 ?④퀎濡??대룞?⑸땲??");
                AdvanceIterationLocked();
            }
        }

        private static void HandleWaitingForDialogToCloseLocked()
        {
            if (!IsCurrentDocumentActiveLocked())
            {
                WriteLog("Active document changed during purge; restoring target document.");
                _state = PurgeBatchState.ActivateCurrentDocument;
                _nextActionUtc = DateTime.UtcNow.AddMilliseconds(300);
                return;
            }

            IntPtr dialogHandle = FindRevitModalDialog(_mainWindowHandle);
            if (dialogHandle == IntPtr.Zero)
            {
                WriteLog("Purge ??붿긽???ロ옒 ?뺤씤");
                AdvanceIterationLocked();
                return;
            }

            TimeSpan sinceLastEnter = DateTime.UtcNow - _lastEnterUtc;
            if (sinceLastEnter > TimeSpan.FromSeconds(2) && _dialogEnterAttempts < MaxEnterAttemptsPerIteration)
            {
                RefocusRevitForPurgeLocked("Purge ??붿긽???뺤씤");
                EnsureWindowTopMost(dialogHandle);
                FocusWindow(dialogHandle);
                SendEnterToWindow(dialogHandle);
                _dialogEnterAttempts++;
                _lastEnterUtc = DateTime.UtcNow;
                WriteLog("Purge ??붿긽???ы솗??/ Enter ?ъ쟾??" + _dialogEnterAttempts + "/" + MaxEnterAttemptsPerIteration);
                return;
            }

            if (_commandPostedUtc != DateTime.MinValue && DateTime.UtcNow - _commandPostedUtc > TimeSpan.FromSeconds(DialogCloseTimeoutSeconds))
            {
                WriteLog("Purge ??붿긽??醫낅즺 ?湲??쒓컙??珥덇낵?섏뼱 ?ㅼ쓬 ?④퀎濡??대룞?⑸땲??");
                AdvanceIterationLocked();
            }
        }

        private static void EnsureSubscribedToIdlingLocked()
        {
            if (_subscribedToIdling) return;
            _uiapp.Idling += OnIdling;
            _subscribedToIdling = true;
        }

        private static void StartDialogWatcherLocked()
        {
            if (_dialogWatcherTimer != null)
            {
                _dialogWatcherTimer.Dispose();
            }
            _dialogWatcherTimer = new Timer(OnDialogWatcherTick, null, 200, 200);
        }

        private static void StopLocked(bool showCompletionDialog)
        {
            if (_subscribedToIdling && _uiapp != null)
            {
                _uiapp.Idling -= OnIdling;
            }
            _subscribedToIdling = false;

            if (_dialogWatcherTimer != null)
            {
                _dialogWatcherTimer.Dispose();
                _dialogWatcherTimer = null;
            }

            ReleaseRevitTopMostLocked();
            try { PurgeProgressWindowHost.CloseWindow(); } catch { }

            bool wasRunning = _running;
            if (showCompletionDialog) _completedSuccessfully = true;
            _running = false;
            _state = PurgeBatchState.Idle;
            _remainingPaths = new List<string>();
            _currentPath = null;
            _currentDocument = null;
            _session = null;
            _anchorPath = null;
            _anchorTitle = null;
            _initialFileCount = 0;
            _currentIteration = 1;
            _processedFileCount = 0;
            _commandPostedUtc = DateTime.MinValue;
            _nextActionUtc = DateTime.MinValue;
            _lastEnterUtc = DateTime.MinValue;
            _dialogEnterAttempts = 0;
            _docSwitchAttempts = 0;
            _purgeCommandId = null;
            _mainWindowTopMost = false;
            _lastRevitRefocusUtc = DateTime.MinValue;
            if (_externalEventHandler != null) _externalEventHandler.PendingOpenPath = null;

            if (showCompletionDialog && wasRunning)
            {
                try
                {
                    TaskDialog.Show("KKY RVT Cleaner", "Purge 일괄처리가 종료되었습니다. 각 파일 저장과 백업 정리가 완료되었습니다." + Environment.NewLine + "다시 로드하면 직전 설정과 결과 목록을 복원합니다.");
                }
                catch
                {
                }
            }
        }

        private static void SetFaultedLocked()
        {
            _faulted = true;
            _state = PurgeBatchState.Faulted;
            _nextActionUtc = DateTime.UtcNow.AddMilliseconds(100);
        }

        private static Document GetCurrentDocumentLocked()
        {
            if (_currentDocument != null && _currentDocument.IsValidObject)
            {
                return _currentDocument;
            }

            _currentDocument = FindOpenDocumentByPath(_currentPath);
            return _currentDocument;
        }

        private static Document GetActiveDocument(UIApplication uiapp)
        {
            try
            {
                return uiapp != null && uiapp.ActiveUIDocument != null ? uiapp.ActiveUIDocument.Document : null;
            }
            catch
            {
                return null;
            }
        }

        private static Document FindOpenDocumentByPath(string path)
        {
            if (_uiapp == null || string.IsNullOrWhiteSpace(path)) return null;
            try
            {
                foreach (Document doc in _uiapp.Application.Documents)
                {
                    if (IsMatchingDocument(doc, path, null))
                    {
                        return doc;
                    }
                }
            }
            catch
            {
            }
            return null;
        }

        private static bool IsCurrentDocumentActiveLocked()
        {
            Document activeDoc = GetActiveDocument(_uiapp);
            return IsMatchingDocument(activeDoc, _currentPath, Path.GetFileNameWithoutExtension(_currentPath ?? string.Empty));
        }

        private static bool IsAnchorActiveLocked()
        {
            Document activeDoc = GetActiveDocument(_uiapp);
            return IsMatchingDocument(activeDoc, _anchorPath, _anchorTitle);
        }

        private static bool IsMatchingDocument(Document doc, string path, string fallbackTitle)
        {
            if (doc == null) return false;
            string docPath = SafePath(doc);
            if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(docPath) && string.Equals(docPath, path, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string title = SafeTitle(doc);
            if (!string.IsNullOrWhiteSpace(fallbackTitle) && !string.IsNullOrWhiteSpace(title) && string.Equals(title, fallbackTitle, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(title))
            {
                string fileName = Path.GetFileNameWithoutExtension(path);
                return string.Equals(title, fileName, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static string SafePath(Document doc)
        {
            try { return doc != null ? doc.PathName : string.Empty; } catch { return string.Empty; }
        }

        private static string SafeTitle(Document doc)
        {
            try { return doc != null ? doc.Title : string.Empty; } catch { return string.Empty; }
        }

        private static void DeleteBackupFiles(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return;
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);
                string pattern = "^" + Regex.Escape(fileName) + @"\.\d{4}" + Regex.Escape(extension) + "$";
                foreach (string candidate in Directory.GetFiles(directory, fileName + ".*" + extension))
                {
                    string name = Path.GetFileName(candidate);
                    if (Regex.IsMatch(name, pattern, RegexOptions.IgnoreCase))
                    {
                        try
                        {
                            File.Delete(candidate);
                            WriteLog("諛깆뾽 ?뚯씪 ??젣: " + candidate);
                        }
                        catch (Exception ex)
                        {
                            WriteLog("諛깆뾽 ?뚯씪 ??젣 ?ㅽ뙣(臾댁떆): " + candidate + " / " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("諛깆뾽 ?뚯씪 ?뺣━ ?ㅽ뙣(臾댁떆): " + ex.Message);
            }
        }

        private static IntPtr GetMainWindowHandle(UIApplication uiapp)
        {
            try
            {
                if (uiapp != null)
                {
                    IntPtr handle = uiapp.MainWindowHandle;
                    if (handle != IntPtr.Zero)
                    {
                        return handle;
                    }
                }
            }
            catch
            {
            }

            try
            {
                return Process.GetCurrentProcess().MainWindowHandle;
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        private static void EnsureRevitTopMostLocked()
        {
            if (_mainWindowHandle == IntPtr.Zero) return;
            if (_mainWindowTopMost) return;

            try
            {
                SetWindowPos(_mainWindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                _mainWindowTopMost = true;
                _lastStatusMessage = "Revit 李???긽 ???좎?";
            }
            catch
            {
            }
        }

        private static void EnsureWindowTopMost(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero) return;
            try
            {
                SetWindowPos(windowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            }
            catch
            {
            }
        }

        private static void ReleaseRevitTopMostLocked()
        {
            if (_mainWindowHandle == IntPtr.Zero) return;
            try
            {
                SetWindowPos(_mainWindowHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE);
            }
            catch
            {
            }
        }

        private static void RefocusRevitForPurgeLocked(string reason)
        {
            if (_mainWindowHandle == IntPtr.Zero) return;
            if (DateTime.UtcNow - _lastRevitRefocusUtc < TimeSpan.FromMilliseconds(RevitRefocusCooldownMilliseconds)) return;

            try
            {
                EnsureRevitTopMostLocked();
                ShowWindow(_mainWindowHandle, 9);
                BringWindowToTop(_mainWindowHandle);
                SetForegroundWindow(_mainWindowHandle);
                _lastRevitRefocusUtc = DateTime.UtcNow;
                _lastStatusMessage = string.IsNullOrWhiteSpace(reason) ? "Revit ?ъ빱??蹂듦뎄" : reason;
            }
            catch
            {
            }
        }

        private static void WriteLog(string message)
        {
            _lastStatusMessage = message ?? string.Empty;
            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message;
            App.AppendSharedLog(line);
        }

        private static RevitCommandId ResolvePurgeCommandId()
        {
            try
            {
                RevitCommandId commandId = RevitCommandId.LookupCommandId("ID_PURGE_UNUSED");
                if (commandId != null) return commandId;
            }
            catch { }

            try
            {
                Type postableType = typeof(PostableCommand);
                object purgeEnum = Enum.Parse(postableType, "PurgeUnused");
                if (purgeEnum != null)
                {
                    return RevitCommandId.LookupPostableCommandId((PostableCommand)purgeEnum);
                }
            }
            catch { }

            return null;
        }

        private static IntPtr FindRevitModalDialog(IntPtr mainWindowHandle)
        {
            IntPtr found = IntPtr.Zero;
            int currentProcessId = Process.GetCurrentProcess().Id;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (found != IntPtr.Zero) return false;
                if (!IsWindowVisible(hWnd)) return true;

                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);
                if (processId != currentProcessId) return true;
                if (hWnd == mainWindowHandle) return true;
                if (GetWindow(hWnd, 4) == IntPtr.Zero) return true;

                string className = GetWindowClassName(hWnd);
                string title = GetWindowTextSafe(hWnd);
                if (className == "#32770" && !string.IsNullOrWhiteSpace(title))
                {
                    found = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static string GetWindowClassName(IntPtr hWnd)
        {
            var buffer = new System.Text.StringBuilder(256);
            GetClassName(hWnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        private static string GetWindowTextSafe(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length <= 0) return string.Empty;
            var buffer = new System.Text.StringBuilder(length + 1);
            GetWindowText(hWnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        private static void FocusWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            ShowWindow(hWnd, 9);
            BringWindowToTop(hWnd);
            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            SetForegroundWindow(hWnd);
        }

        private static void SendCtrlTab()
        {
            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_TAB, 0, 0, UIntPtr.Zero);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        private static void SendEnterToWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            PostMessage(hWnd, WM_KEYDOWN, new IntPtr(VK_RETURN), IntPtr.Zero);
            PostMessage(hWnd, WM_KEYUP, new IntPtr(VK_RETURN), IntPtr.Zero);
        }

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

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
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    }
}

