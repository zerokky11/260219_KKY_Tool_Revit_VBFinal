using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public partial class FamilyBrowserDashboardHtmlForm
{
    private readonly Dictionary<string, FamilyBrowserStandardRevisionState> _standardRevisionStates = new Dictionary<string, FamilyBrowserStandardRevisionState>(StringComparer.OrdinalIgnoreCase);
    private int _standardRevisionProbeGeneration;
    private DateTime _lastStandardRevisionProbeUtc = DateTime.MinValue;
    private bool _activeStandardRevisionBlocked;
    private FamilyBrowserStandardRevisionState _activeStandardRevisionState;

    private void ApplyPreparedStandardRevisionStates(FamilyBrowserStartupPreloadResult preload)
    {
        _standardRevisionStates.Clear();
        if (preload == null || preload.Slots == null)
        {
            return;
        }
        foreach (FamilyBrowserPreparedSlotData prepared in preload.Slots)
        {
            if (prepared == null || prepared.StandardRevisionState == null)
            {
                continue;
            }
            StoreStandardRevisionState(prepared.Registration, prepared.SlotKey, prepared.StandardRevisionState);
        }
    }

    private void StoreStandardRevisionState(StandardLibraryRegistrationRecord registration, string slotKey, FamilyBrowserStandardRevisionState state)
    {
        if (state == null)
        {
            return;
        }
        string key = StandardRevisionKey(registration, slotKey);
        if (!string.IsNullOrWhiteSpace(key))
        {
            _standardRevisionStates[key] = state;
        }
        if (registration != null && !string.IsNullOrWhiteSpace(registration.SourceId))
        {
            _standardRevisionStates["source:" + registration.SourceId.Trim()] = state;
        }
    }

    private FamilyBrowserStandardRevisionState ResolveStandardRevisionState(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration, bool probeWhenMissing)
    {
        string slotKey = slot == null ? string.Empty : ResolveSlotKey(slot);
        string key = StandardRevisionKey(registration, slotKey);
        FamilyBrowserStandardRevisionState state;
        if (!string.IsNullOrWhiteSpace(key) && _standardRevisionStates.TryGetValue(key, out state))
        {
            return state;
        }
        if (registration != null && !string.IsNullOrWhiteSpace(registration.SourceId) && _standardRevisionStates.TryGetValue("source:" + registration.SourceId.Trim(), out state))
        {
            return state;
        }
        if (!probeWhenMissing || registration == null)
        {
            return null;
        }
        state = FamilyBrowserStandardRevisionService.Probe(_workspaceRoot, registration, false);
        StoreStandardRevisionState(registration, slotKey, state);
        return state;
    }

    private bool ApplyStandardRevisionBlockIfNeeded(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration, bool forceProbe = false)
    {
        if (forceProbe && registration != null)
        {
            _activeStandardRevisionState = FamilyBrowserStandardRevisionService.Probe(_workspaceRoot, registration, true);
            StoreStandardRevisionState(registration, slot == null ? string.Empty : ResolveSlotKey(slot), _activeStandardRevisionState);
        }
        else
        {
            _activeStandardRevisionState = ResolveStandardRevisionState(slot, registration, true);
        }
        _activeStandardRevisionBlocked = _activeStandardRevisionState != null && _activeStandardRevisionState.BlocksStandardUse;
        if (!_activeStandardRevisionBlocked)
        {
            return false;
        }
        _activeStandardScanNeeded = true;
        _loadableRows = new List<BrowserRow>();
        _systemRows = new List<SystemRow>();
        _systemComparisonRows = new List<SystemRow>();
        _loadText = T("Load Available: blocked", "로드 가능: 차단");
        _updateText = T("Update Available: blocked", "업데이트 가능: 차단");
        _trackingText = T("Standard source: rescan required", "표준 원본: 재스캔 필요");
        _trackingTone = "bad";
        _permissionText = T("Permission: standard revision blocked", "권한: 표준 revision 차단");
        _statusMessage = BuildStandardRevisionBlockingMessage(_activeStandardRevisionState);
        _projectScanText = _statusMessage;
        _systemSummary = _statusMessage;
        _nextWorkflowText = _statusMessage;
        return true;
    }

    private bool ValidateStandardRevisionAfterOperation(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration, FamilyBrowserStandardRevisionState before)
    {
        FamilyBrowserStandardRevisionState after = FamilyBrowserStandardRevisionService.Probe(_workspaceRoot, registration, true);
        StoreStandardRevisionState(registration, slot == null ? string.Empty : ResolveSlotKey(slot), after);
        if (FamilyBrowserStandardRevisionService.IsSameCurrentRevision(before, after))
        {
            _activeStandardRevisionState = after;
            _activeStandardRevisionBlocked = false;
            return true;
        }
        _activeStandardRevisionState = after;
        _activeStandardRevisionBlocked = true;
        _activeStandardScanNeeded = true;
        _loadableRows = new List<BrowserRow>();
        _systemRows = new List<SystemRow>();
        _systemComparisonRows = new List<SystemRow>();
        _loadText = T("Load Available: blocked", "로드 가능: 차단");
        _updateText = T("Update Available: blocked", "업데이트 가능: 차단");
        _trackingText = T("Standard source changed during check", "검사 중 표준 원본 변경됨");
        _trackingTone = "bad";
        _permissionText = T("Permission: rerun required", "권한: 재검사 필요");
        _statusMessage = after != null && after.BlocksStandardUse
            ? BuildStandardRevisionBlockingMessage(after)
            : T("The Standard RVT revision changed while Current Model Check was running. No shared result was published. Run the check again.", "현재 모델 검사 중 표준 RVT revision이 바뀌어 공용 결과를 저장하지 않았습니다. 검사를 다시 실행하세요.");
        _projectScanText = _statusMessage;
        _systemSummary = _statusMessage;
        _nextWorkflowText = _statusMessage;
        return false;
    }

    private void ResetActiveStandardRevisionBlock()
    {
        _activeStandardRevisionBlocked = false;
        _activeStandardRevisionState = null;
    }

    private void RecordStandardRevisionBaseline(StandardLibraryRegistrationRecord registration, string slotKey)
    {
        if (registration == null)
        {
            return;
        }
        FamilyBrowserStandardRevisionState state = FamilyBrowserStandardRevisionService.RecordBaseline(_workspaceRoot, registration, Environment.UserName);
        StoreStandardRevisionState(registration, slotKey, state);
        _activeStandardRevisionState = state;
        _activeStandardRevisionBlocked = state != null && state.BlocksStandardUse;
    }

    private void QueueStandardRevisionProbe(bool computeRevisionHash, bool force = false)
    {
        if (!force && (DateTime.UtcNow - _lastStandardRevisionProbeUtc).TotalSeconds < 55.0)
        {
            return;
        }
        _lastStandardRevisionProbeUtc = DateTime.UtcNow;
        FamilyBrowserStandardPolicy policy = _standardPolicy ?? LoadStandardPolicy();
        FamilyBrowserStandardLibrarySlot activeSlot = ResolveBrowseSlot(policy);
        List<Tuple<string, StandardLibraryRegistrationRecord>> targets = new List<Tuple<string, StandardLibraryRegistrationRecord>>();
        foreach (FamilyBrowserStandardLibrarySlot slot in GetBrowserStandardSlots(policy).Where(delegate(FamilyBrowserStandardLibrarySlot x) { return x != null && x.Enabled; }))
        {
            StandardLibraryRegistrationRecord registration = ResolveStandardRevisionRegistration(policy, slot, activeSlot);
            if (registration != null)
            {
                targets.Add(Tuple.Create(ResolveSlotKey(slot), registration));
            }
        }
        int generation = Interlocked.Increment(ref _standardRevisionProbeGeneration);
        Task.Factory.StartNew(delegate
        {
            List<Tuple<string, StandardLibraryRegistrationRecord, FamilyBrowserStandardRevisionState>> states = new List<Tuple<string, StandardLibraryRegistrationRecord, FamilyBrowserStandardRevisionState>>();
            foreach (Tuple<string, StandardLibraryRegistrationRecord> target in targets)
            {
                states.Add(Tuple.Create(target.Item1, target.Item2, FamilyBrowserStandardRevisionService.Probe(_workspaceRoot, target.Item2, computeRevisionHash)));
            }
            return states;
        }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default).ContinueWith(delegate(Task<List<Tuple<string, StandardLibraryRegistrationRecord, FamilyBrowserStandardRevisionState>>> task)
        {
            if (task.IsFaulted || task.IsCanceled || generation != _standardRevisionProbeGeneration || IsDisposed)
            {
                return;
            }
            try
            {
                BeginInvoke(new Action(delegate
                {
                    if (generation != _standardRevisionProbeGeneration || IsDisposed)
                    {
                        return;
                    }
                    bool changed = false;
                    foreach (Tuple<string, StandardLibraryRegistrationRecord, FamilyBrowserStandardRevisionState> item in task.Result)
                    {
                        FamilyBrowserStandardRevisionState previous = ResolveStandardRevisionState(null, item.Item2, false);
                        changed = changed || !StandardRevisionStatesEquivalent(previous, item.Item3);
                        StoreStandardRevisionState(item.Item2, item.Item1, item.Item3);
                    }
                    if (changed)
                    {
						FamilyBrowserStandardLibrarySlot activeSlot = ResolveBrowseSlot(_standardPolicy ?? LoadStandardPolicy());
						StandardLibraryRegistrationRecord activeRegistration = activeSlot == null ? _registration : TryLoadRegistrationForSlot(_standardPolicy ?? LoadStandardPolicy(), activeSlot);
						if (!ApplyStandardRevisionBlockIfNeeded(activeSlot, activeRegistration))
						{
							ResetActiveStandardRevisionBlock();
						}
                        RefreshDocumentShellOnly();
                    }
                }));
            }
            catch
            {
            }
        }, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
    }

    private static bool StandardRevisionStatesEquivalent(FamilyBrowserStandardRevisionState left, FamilyBrowserStandardRevisionState right)
    {
        if (left == null || right == null)
        {
            return left == right;
        }
        return string.Equals(left.StateCode, right.StateCode, StringComparison.OrdinalIgnoreCase)
            && left.Changed == right.Changed
            && left.Unavailable == right.Unavailable
            && left.BaselineMissing == right.BaselineMissing
            && left.CurrentLength == right.CurrentLength
            && string.Equals(left.CurrentLastWriteUtc, right.CurrentLastWriteUtc, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.CurrentRevisionHash, right.CurrentRevisionHash, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.FileIdentity, right.FileIdentity, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.SnapshotPath, right.SnapshotPath, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.SnapshotAtUtc, right.SnapshotAtUtc, StringComparison.OrdinalIgnoreCase);
    }

    private void AppendStandardRevisionPill(StringBuilder sb)
    {
        FamilyBrowserStandardRevisionState state = ResolveStandardRevisionState(ResolveBrowseSlot(_standardPolicy), _registration, false);
        if (_registration == null && state == null)
        {
            return;
        }
        sb.AppendLine(Pill(T("Standard Source: ", "표준 원본: ") + StandardRevisionStatusTitle(state), StandardRevisionTone(state)));
    }

    private void AppendHomeStandardRevisionBoard(StringBuilder sb)
    {
        FamilyBrowserStandardPolicy policy = _standardPolicy ?? LoadStandardPolicy();
        FamilyBrowserStandardLibrarySlot activeSlot = ResolveBrowseSlot(policy);
        List<Tuple<string, FamilyBrowserStandardRevisionState>> rows = new List<Tuple<string, FamilyBrowserStandardRevisionState>>();
        foreach (FamilyBrowserStandardLibrarySlot slot in GetBrowserStandardSlots(policy).Where(delegate(FamilyBrowserStandardLibrarySlot x) { return x != null && x.Enabled; }))
        {
            StandardLibraryRegistrationRecord registration = ResolveStandardRevisionRegistration(policy, slot, activeSlot);
            if (registration != null)
            {
                rows.Add(Tuple.Create(ResolveSlotLabel(slot), ResolveStandardRevisionState(slot, registration, false)));
            }
        }
        if (rows.Count == 0)
        {
            return;
        }
        bool blocked = rows.Any(delegate(Tuple<string, FamilyBrowserStandardRevisionState> x) { return x.Item2 != null && x.Item2.BlocksStandardUse; });
        sb.AppendLine("<div class=\"home-board standard-revision-board " + (blocked ? "bad" : "good") + "\">");
        sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Standard RVT Source Status", "표준 RVT 원본 상태")) + "</strong><span>" + Html(T("The source file is checked automatically against the last accepted scan.", "마지막 승인 스캔과 원본 파일을 자동으로 대조합니다.")) + "</span></div>");
        sb.AppendLine("<div class=\"standard-revision-grid\">");
        foreach (Tuple<string, FamilyBrowserStandardRevisionState> row in rows)
        {
            sb.AppendLine("<div class=\"standard-revision-item " + Attr(StandardRevisionTone(row.Item2)) + "\"><strong>" + Html(row.Item1) + "</strong><span>" + Html(StandardRevisionStatusTitle(row.Item2)) + "</span><em title=\"" + Attr(StandardRevisionReason(row.Item2)) + "\">" + Html(SafeShortName(StandardRevisionReason(row.Item2), 120)) + "</em></div>");
        }
        sb.AppendLine("</div>");
        if (blocked)
        {
            sb.AppendLine("<a class=\"tool primary standard-revision-action\" href=\"" + Attr(DashboardTabHref("admin")) + "\">" + Html(T("Open Standards and Rescan", "표준 관리에서 재스캔")) + "</a>");
        }
        sb.AppendLine("</div>");
    }

    private StandardLibraryRegistrationRecord ResolveStandardRevisionRegistration(
        FamilyBrowserStandardPolicy policy,
        FamilyBrowserStandardLibrarySlot slot,
        FamilyBrowserStandardLibrarySlot activeSlot)
    {
        StandardLibraryRegistrationRecord registration = slot == null ? null : TryLoadRegistrationForSlot(policy, slot);
        if (registration != null || _registration == null || slot == null)
        {
            return registration;
        }

        string slotKey = ResolveSlotKey(slot);
        string activeSlotKey = activeSlot == null ? _browseDisciplineKey : ResolveSlotKey(activeSlot);
        bool isActiveSlot = object.ReferenceEquals(slot, activeSlot)
            || string.Equals(Normalize(slotKey), Normalize(activeSlotKey), StringComparison.Ordinal);
        return isActiveSlot ? _registration : null;
    }

    private void AppendHomeStandardRevisionRecentItems(StringBuilder sb, FamilyBrowserStandardPolicy policy)
    {
        FamilyBrowserStandardLibrarySlot slot = ResolveBrowseSlot(policy);
        StandardLibraryRegistrationRecord registration = slot == null ? null : TryLoadRegistrationForSlot(policy, slot);
		if (registration == null && _registration != null)
		{
			registration = _registration;
		}
        FamilyBrowserStandardRevisionState state = ResolveStandardRevisionState(slot, registration, false);
        if (registration == null)
        {
            return;
        }
        AppendHomeRecentItem(sb, T("Standard RVT Source", "표준 RVT 원본"), StandardRevisionStatusTitle(state), StandardRevisionReason(state), StandardRevisionTone(state));
    }

    private void AppendStandardRevisionManagerDetails(StringBuilder sb, FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
    {
        FamilyBrowserStandardRevisionState state = ResolveStandardRevisionState(slot, registration, true);
        sb.AppendLine("<div class=\"standard-revision-manager " + Attr(StandardRevisionTone(state)) + "\">");
        sb.AppendLine("<div class=\"section\">" + Html(T("Automatic Source Revision Check", "원본 revision 자동 확인")) + "</div>");
        sb.AppendLine("<div class=\"standard-revision-summary\"><strong>" + Html(StandardRevisionStatusTitle(state)) + "</strong><span>" + Html(StandardRevisionReason(state)) + "</span></div>");
        if (state != null)
        {
            AppendAdminRow(sb, T("Last checked", "마지막 확인"), FormatRevisionDate(state.CheckedAtUtc));
            AppendAdminRow(sb, T("Last accepted source", "마지막 승인 원본"), FormatRevisionFileStamp(state.RecordedLastWriteUtc, state.RecordedLength));
            AppendAdminRow(sb, T("Current source", "현재 원본"), FormatRevisionFileStamp(state.CurrentLastWriteUtc, state.CurrentLength));
            AppendAdminRow(sb, T("Path identity", "경로 식별"), state.PathAliasMatched ? T("Same file through another path alias", "다른 경로 표기지만 동일 파일") : SafeShortName(state.FileIdentity, 48));
        }
        sb.AppendLine("</div>");
        AppendStandardChangeHistoryTable(sb, registration, 50);
    }

    private void AppendStandardChangeHistoryTable(StringBuilder sb, StandardLibraryRegistrationRecord registration, int limit)
    {
        if (registration == null)
        {
            return;
        }
        List<StandardRvtChangeCandidateEntry> entries = StandardRvtChangeCandidateService.LoadRecent(_workspaceRoot, registration.SourceId, limit);
        sb.AppendLine("<div class=\"standard-change-history\"><div class=\"section\">" + Html(T("Confirmed Standard RVT Change History", "확정 표준 RVT 변경 이력")) + "</div>");
        if (entries.Count == 0)
        {
            sb.AppendLine("<div class=\"empty compact\">" + Html(T("No confirmed changes have been recorded yet.", "아직 저장이 확인된 변경 이력이 없습니다.")) + "</div></div>");
            return;
        }
        sb.AppendLine("<div class=\"table-scroll\"><table class=\"standard-change-table\"><tr><th>" + Html(T("Saved", "저장 시각")) + "</th><th>" + Html(T("User", "사용자")) + "</th><th>" + Html(T("Change", "변경")) + "</th><th>" + Html(T("Item", "항목")) + "</th><th>" + Html(T("Before / After", "이전 / 이후")) + "</th><th>" + Html(T("Commit", "확정")) + "</th></tr>");
        foreach (StandardRvtChangeCandidateEntry entry in entries)
        {
            string user = string.IsNullOrWhiteSpace(entry.RevitUserName) ? entry.UserName : entry.RevitUserName + " / " + entry.UserName;
            string item = BuildStandardHistoryItem(entry);
            string beforeAfter = (string.IsNullOrWhiteSpace(entry.BeforeFingerprint) ? "-" : entry.BeforeFingerprint) + " -> " + (string.IsNullOrWhiteSpace(entry.AfterFingerprint) ? "-" : entry.AfterFingerprint);
            string commit = (entry.CommitKind ?? string.Empty) + (string.IsNullOrWhiteSpace(entry.MachineName) ? string.Empty : " / " + entry.MachineName);
            sb.AppendLine("<tr><td>" + Html(FormatRevisionDate(string.IsNullOrWhiteSpace(entry.CommittedAtUtc) ? entry.RecordedAtUtc : entry.CommittedAtUtc)) + "</td><td title=\"" + Attr(user) + "\">" + Html(SafeShortName(user, 42)) + "</td><td>" + Html(StandardHistoryChangeLabel(entry)) + "</td><td title=\"" + Attr(item) + "\">" + Html(SafeShortName(item, 70)) + "</td><td title=\"" + Attr(beforeAfter) + "\">" + Html(SafeShortName(beforeAfter, 84)) + "</td><td>" + Html(SafeShortName(commit, 40)) + "</td></tr>");
        }
        sb.AppendLine("</table></div></div>");
    }

    private string BuildStandardRevisionBlockingMessage(FamilyBrowserStandardRevisionState state)
    {
        return T("The Standard RVT source changed or cannot be verified. The existing snapshot is read-only and load/apply/model-check actions are blocked until the standard is scanned again. ", "표준 RVT 원본이 변경되었거나 확인할 수 없습니다. 기존 스냅샷은 읽기 전용으로 취급하며 표준을 다시 스캔할 때까지 로드/적용/모델 검사를 차단합니다. ") + StandardRevisionReason(state);
    }

	private void RefreshStandardRevisionLocalizedState()
	{
		if (!_activeStandardRevisionBlocked || _activeStandardRevisionState == null)
		{
			return;
		}
		_trackingText = T("Standard source: rescan required", "표준 원본: 재스캔 필요");
		_trackingTone = "bad";
		_permissionText = T("Permission: standard revision blocked", "권한: 표준 revision 차단");
		_projectScanText = BuildStandardRevisionBlockingMessage(_activeStandardRevisionState);
		_systemSummary = _projectScanText;
		_nextWorkflowText = _projectScanText;
	}

    private string StandardRevisionStatusTitle(FamilyBrowserStandardRevisionState state)
    {
        if (state == null) return T("checking", "확인 중");
        if (state.Changed) return T("changed - rescan required", "변경됨 - 재스캔 필요");
        if (state.Unavailable) return T("source unavailable", "원본 연결 불가");
        if (state.BaselineMissing) return T("baseline missing", "기준선 없음");
        if (!string.IsNullOrWhiteSpace(state.ErrorMessage)) return T("check failed", "확인 실패");
        if (string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase)) return T("current", "최신");
        return T("not checked", "미확인");
    }

    private string StandardRevisionReason(FamilyBrowserStandardRevisionState state)
    {
        if (state == null) return T("Automatic source verification is still running.", "원본 자동 확인이 아직 실행 중입니다.");
		if (state.Unavailable) return T("The registered Standard RVT source cannot be found.", "등록된 표준 RVT 원본을 찾을 수 없습니다.");
		if (state.Changed)
		{
			if (!string.IsNullOrWhiteSpace(state.Reason) && state.Reason.IndexOf("identity", StringComparison.OrdinalIgnoreCase) >= 0)
				return T("The Standard RVT file identity changed after the last accepted scan.", "마지막 승인 스캔 이후 표준 RVT 파일 식별 정보가 달라졌습니다.");
			if (!string.IsNullOrWhiteSpace(state.Reason) && state.Reason.IndexOf("content revision", StringComparison.OrdinalIgnoreCase) >= 0)
				return T("The Standard RVT contents changed even though the file time and size are the same.", "파일 시각과 크기는 같지만 표준 RVT 내용이 변경되었습니다.");
			return T("The Standard RVT was modified after the last accepted scan.", "마지막 승인 스캔 이후 표준 RVT가 수정되었습니다.");
		}
		if (state.BaselineMissing) return T("No accepted scan revision is available for this Standard RVT.", "이 표준 RVT에 승인된 스캔 revision이 없습니다.");
		if (!string.IsNullOrWhiteSpace(state.ErrorMessage)) return T("The source revision check failed: ", "원본 revision 확인 실패: ") + state.ErrorMessage;
		if (state.PathAliasMatched) return T("A different path expression resolves to the same Standard RVT file.", "경로 표기는 다르지만 동일한 표준 RVT 파일로 확인되었습니다.");
		if (string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase)) return T("The Standard RVT matches the last accepted scan.", "표준 RVT가 마지막 승인 스캔과 일치합니다.");
		if (!string.IsNullOrWhiteSpace(state.Reason)) return state.Reason;
        return T("No revision details are available.", "revision 상세 정보가 없습니다.");
    }

    private static string StandardRevisionTone(FamilyBrowserStandardRevisionState state)
    {
        if (state == null) return "info";
        if (state.BlocksStandardUse) return "bad";
        return string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase) ? "good" : "warn";
    }

    private static string StandardRevisionKey(StandardLibraryRegistrationRecord registration, string slotKey)
    {
        if (registration != null && !string.IsNullOrWhiteSpace(registration.SourceId)) return "source:" + registration.SourceId.Trim();
        if (!string.IsNullOrWhiteSpace(slotKey)) return "slot:" + slotKey.Trim();
        return string.Empty;
    }

    private static string FormatRevisionDate(string value)
    {
        DateTime date;
        if (DateTime.TryParse(value ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out date))
        {
            return date.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }
        return string.IsNullOrWhiteSpace(value) ? "-" : value;
    }

    private static string FormatRevisionFileStamp(string date, long length)
    {
        string stamp = FormatRevisionDate(date);
        if (length <= 0L) return stamp;
        return stamp + " / " + (length / 1024d / 1024d).ToString("0.0", CultureInfo.InvariantCulture) + " MB";
    }

    private static string BuildStandardHistoryItem(StandardRvtChangeCandidateEntry entry)
    {
        if (entry == null) return string.Empty;
        List<string> parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(entry.CategoryName)) parts.Add(entry.CategoryName);
        if (!string.IsNullOrWhiteSpace(entry.FamilyName)) parts.Add(entry.FamilyName);
        if (!string.IsNullOrWhiteSpace(entry.TypeName)) parts.Add(entry.TypeName);
        if (parts.Count == 0 && !string.IsNullOrWhiteSpace(entry.SystemFamilyKind)) parts.Add(entry.SystemFamilyKind);
        return string.Join(" / ", parts);
    }

    private string StandardHistoryChangeLabel(StandardRvtChangeCandidateEntry entry)
    {
        string kind = entry == null ? string.Empty : entry.ChangeKind ?? string.Empty;
        if (string.Equals(kind, "Added", StringComparison.OrdinalIgnoreCase)) return T("Added", "추가");
        if (string.Equals(kind, "Deleted", StringComparison.OrdinalIgnoreCase)) return T("Deleted", "삭제");
        if (string.Equals(kind, "Modified", StringComparison.OrdinalIgnoreCase)) return T("Modified", "수정");
        return kind;
    }
}
