using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;

public partial class FamilyBrowserDashboardHtmlForm
{
    private FamilyBrowserProjectCatalogState _projectCatalogState;
    private bool _projectCatalogObservationQueued;
    private DateTime _lastProjectCatalogObservationUtc = DateTime.MinValue;
    private string _pendingProjectCatalogTrigger = string.Empty;
    private string _lastProjectCatalogWarningToken = string.Empty;
    private bool _projectElementHistoryLoadPending;

    private sealed class ProjectElementHistoryLoadBundle
    {
        public FamilyBrowserElementChangeHistoryLoadResult History { get; set; }
        public FamilyBrowserElementChangeHistoryLoadResult UploadPendingHistory { get; set; }
        public FamilyBrowserElementSessionCheckpointHistoryLoadResult CheckpointHistory { get; set; }
        public FamilyBrowserElementSessionCheckpointCountResult LocalSyncPendingStatus { get; set; }
        public int InvalidLocalSyncPending { get; set; }
        public int MismatchedLocalSyncPending { get; set; }
        public FamilyBrowserOperationLogEntry LatestPolicyChange { get; set; }

        public ProjectElementHistoryLoadBundle()
        {
            History = new FamilyBrowserElementChangeHistoryLoadResult();
            UploadPendingHistory = new FamilyBrowserElementChangeHistoryLoadResult();
            CheckpointHistory = new FamilyBrowserElementSessionCheckpointHistoryLoadResult();
            LocalSyncPendingStatus = new FamilyBrowserElementSessionCheckpointCountResult();
        }
    }

    private void QueueProjectCatalogObservation(string trigger, bool force)
    {
        if (IsDisposed || _modelessActionDispatcher == null || _projectCatalogObservationQueued)
        {
            return;
        }
        if (!force && (DateTime.UtcNow - _lastProjectCatalogObservationUtc).TotalSeconds < 55.0)
        {
            return;
        }
        _pendingProjectCatalogTrigger = string.IsNullOrWhiteSpace(trigger) ? "Automatic" : trigger.Trim();
        _projectCatalogObservationQueued = true;
        if (!_modelessActionDispatcher("project-catalog-observe-auto"))
        {
            _projectCatalogObservationQueued = false;
            _pendingProjectCatalogTrigger = string.Empty;
            return;
        }
        _lastProjectCatalogObservationUtc = DateTime.UtcNow;
    }

    private bool TryHandleProjectCatalogAction(string actionKey)
    {
        if (string.Equals(actionKey, "project-catalog-observe-auto", StringComparison.OrdinalIgnoreCase))
        {
            _projectCatalogObservationQueued = false;
            ObserveCurrentProjectCatalog(_pendingProjectCatalogTrigger, false);
            return true;
        }
        if (string.Equals(actionKey, "project-catalog-check", StringComparison.OrdinalIgnoreCase))
        {
            _projectCatalogObservationQueued = false;
            ObserveCurrentProjectCatalog("ManualCheck", true);
            return true;
        }
        if (string.Equals(actionKey, "project-catalog-accept", StringComparison.OrdinalIgnoreCase))
        {
            AcceptCurrentProjectCatalog();
            return true;
        }
        return false;
    }

    private void ObserveCurrentProjectCatalog(string trigger, bool showResult)
    {
        Document doc = GetActiveDocument();
        _projectCatalogState = FamilyBrowserProjectCatalogService.Observe(_workspaceRoot, doc, trigger);
        WriteDashboardRuntimeDiagnostic(
            "project-catalog-observed:" + (_projectCatalogState == null ? "null" : _projectCatalogState.StateCode)
                + ":families=" + ((_projectCatalogState == null) ? 0 : _projectCatalogState.FamilyCount).ToString(CultureInfo.InvariantCulture)
                + ":familyTypes=" + ((_projectCatalogState == null) ? 0 : _projectCatalogState.FamilyTypeCount).ToString(CultureInfo.InvariantCulture)
                + ":systemTypes=" + ((_projectCatalogState == null) ? 0 : _projectCatalogState.SystemTypeCount).ToString(CultureInfo.InvariantCulture)
                + ":elapsedMs=" + ((_projectCatalogState == null) ? 0L : _projectCatalogState.ElapsedMilliseconds).ToString(CultureInfo.InvariantCulture),
            -1,
            -1L);
        RefreshDocumentShellOnly();
        if (showResult)
        {
            ShowProjectCatalogResult(_projectCatalogState, false);
            return;
        }
        ShowProjectCatalogWarningIfNeeded(_projectCatalogState);
    }

    private void AcceptCurrentProjectCatalog()
    {
        if (!EnsurePermission("ManagePolicy", T("Accept Project Catalog Baseline", "프로젝트 카탈로그 기준선 승인")))
        {
            return;
        }
        string message = T(
            "This accepts only the current family and system type name inventory as the project tracking baseline. It does not prove that family fingerprints, parameters, routing rules, layers, or dependent components match the Standard RVT. Run Current Model Check when content equality must be verified.\n\nAccept the current name catalog?",
            "현재 패밀리명과 시스템 타입명 목록만 프로젝트 추적 기준선으로 승인합니다. 패밀리 Fingerprint, 파라미터, 라우팅 규칙, 레이어, 의존 구성요소가 표준 RVT와 같다는 뜻은 아닙니다. 내용 일치까지 확인하려면 현재 모델 검사를 실행하세요.\n\n현재 이름 카탈로그를 기준선으로 승인할까요?");
        if (ShowDashboardMessage(this, message, T("Accept Project Catalog Baseline", "프로젝트 카탈로그 기준선 승인"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
        {
            return;
        }
        _projectCatalogState = FamilyBrowserProjectCatalogService.AcceptCurrent(_workspaceRoot, GetActiveDocument(), "AdminAccept", Environment.UserName);
        _lastProjectCatalogWarningToken = string.Empty;
        RefreshDocumentShellOnly();
        bool accepted = _projectCatalogState != null &&
            string.Equals(_projectCatalogState.StateCode, "Current", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(_projectCatalogState.AcceptedCatalogHash) &&
            string.IsNullOrWhiteSpace(_projectCatalogState.ErrorMessage);
        ShowProjectCatalogResult(_projectCatalogState, accepted);
    }

    private void AcceptProjectCatalogAfterCurrentModelCheck(Document doc, ProjectContentSnapshot projectSnapshot)
    {
        _projectCatalogState = FamilyBrowserProjectCatalogService.AcceptFromProjectSnapshot(_workspaceRoot, doc, projectSnapshot, "CurrentModelCheck", Environment.UserName);
        _lastProjectCatalogWarningToken = string.Empty;
        WriteDashboardRuntimeDiagnostic(
            "project-catalog-baseline-result:current-model-check:" +
            (_projectCatalogState == null ? "null" : (_projectCatalogState.StateCode ?? string.Empty)) + ":" +
            (_projectCatalogState == null ? string.Empty : (_projectCatalogState.CurrentCatalogHash ?? string.Empty)),
            -1,
            -1L);
    }

    private void ReloadProjectCatalogStateAfterCommit(Document doc, string commitKind)
    {
        _projectCatalogState = FamilyBrowserProjectCatalogService.LoadLatestState(_workspaceRoot, doc);
        WriteDashboardRuntimeDiagnostic("project-catalog-state-reloaded-after:" + (commitKind ?? string.Empty) + ":" + (_projectCatalogState == null ? "null" : _projectCatalogState.StateCode), -1, -1L);
    }

    private void ShowProjectCatalogWarningIfNeeded(FamilyBrowserProjectCatalogState state)
    {
        if (state == null || !state.Changed || state.ExternalUntrackedChangeCount <= 0)
        {
            return;
        }
        string token = (state.ProjectComparableIdentity ?? string.Empty) + "|" + (state.CurrentCatalogHash ?? string.Empty);
        if (string.Equals(token, _lastProjectCatalogWarningToken, StringComparison.Ordinal))
        {
            return;
        }
        _lastProjectCatalogWarningToken = token;
        ShowProjectCatalogResult(state, false);
    }

    private void ShowProjectCatalogResult(FamilyBrowserProjectCatalogState state, bool accepted)
    {
        if (state == null)
        {
            return;
        }
        bool error = !string.IsNullOrWhiteSpace(state.ErrorMessage) || string.Equals(state.StateCode, "StorageUnavailable", StringComparison.OrdinalIgnoreCase);
        bool warning = state.Changed || state.BaselineMissing || string.Equals(state.StateCode, "PublicationDeferred", StringComparison.OrdinalIgnoreCase);
        string caption = accepted
            ? T("Project Catalog Baseline Accepted", "프로젝트 카탈로그 기준선 승인 완료")
            : T("Project Catalog Tracking", "프로젝트 카탈로그 추적");
        StringBuilder message = new StringBuilder();
        message.AppendLine(accepted
            ? T("The current project name catalog is now the accepted tracking baseline.", "현재 프로젝트 이름 카탈로그를 추적 기준선으로 승인했습니다.")
            : ProjectCatalogStatusTitle(state));
        message.AppendLine();
        message.AppendLine(T("Catalog summary", "카탈로그 요약"));
        message.AppendLine(T("Families: ", "패밀리: ") + state.FamilyCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Family types: ", "패밀리 타입: ") + state.FamilyTypeCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("System types: ", "시스템 타입: ") + state.SystemTypeCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Added / removed: ", "추가 / 삭제: ") + state.AddedCount.ToString(CultureInfo.InvariantCulture) + " / " + state.RemovedCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Browser-operation match / external or untracked: ", "브라우저 작업 기록 일치 / 외부 또는 미추적: ") + state.BrowserTrackedChangeCount.ToString(CultureInfo.InvariantCulture) + " / " + state.ExternalUntrackedChangeCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("A Browser-operation match means that a same-project committed Browser record matched the item name after the accepted baseline. It helps triage the change, but it does not by itself prove the actor or the exact synchronization where the change occurred.", "브라우저 작업 기록 일치는 승인 기준선 이후 같은 프로젝트에서 같은 항목명으로 확정된 Browser 기록이 있다는 뜻입니다. 변경 분류에는 도움이 되지만 작업자나 정확한 동기화 시점을 단독으로 증명하지는 않습니다."));
        if (state.Changed)
        {
            message.AppendLine();
            message.AppendLine(T("Detected differences", "감지된 차이"));
            foreach (FamilyBrowserProjectCatalogChange change in (state.Changes ?? new List<FamilyBrowserProjectCatalogChange>()).Take(12))
            {
                message.AppendLine("- " + ProjectCatalogChangeLabel(change));
            }
            if ((state.Changes ?? new List<FamilyBrowserProjectCatalogChange>()).Count > 12)
            {
                message.AppendLine(T("Additional differences are available through Export Excel.", "나머지 차이는 Excel 내보내기에서 확인할 수 있습니다."));
            }
            message.AppendLine();
            message.AppendLine(T("What to do now", "지금 할 일"));
            message.AppendLine(T("Run Current Model Check to verify detailed content. If the changes are intentional, an administrator can accept a new name-catalog baseline afterward.", "현재 모델 검사를 실행해 상세 내용을 확인하세요. 의도한 변경이면 확인 후 관리자가 새 이름 카탈로그 기준선을 승인할 수 있습니다."));
        }
        else if (state.BaselineMissing)
        {
            message.AppendLine();
            message.AppendLine(T("What to do now", "지금 할 일"));
            message.AppendLine(T("Run Current Model Check to create the safest baseline. Administrators may accept only the current name inventory when a detailed check is not yet available.", "가장 안전한 기준선을 만들려면 현재 모델 검사를 실행하세요. 상세 검사를 아직 할 수 없다면 관리자가 현재 이름 목록만 임시 기준선으로 승인할 수 있습니다."));
        }
        if (!string.IsNullOrWhiteSpace(state.ErrorMessage))
        {
            message.AppendLine();
            message.AppendLine(T("Error", "오류"));
            message.AppendLine(state.ErrorMessage);
        }

        List<string> headers = new List<string>
        {
            T("Change", "변경"),
            T("Item kind", "항목 종류"),
            T("Category", "카테고리"),
            T("Family", "패밀리"),
            T("Type", "타입"),
            T("System class", "시스템 클래스"),
            T("Attribution", "출처 판정"),
            T("User", "사용자"),
            T("Committed", "확정 시각")
        };
        List<List<string>> rows = new List<List<string>>();
        foreach (FamilyBrowserProjectCatalogChange change in state.Changes ?? new List<FamilyBrowserProjectCatalogChange>())
        {
            rows.Add(new List<string>
            {
                ProjectCatalogChangeKindLabel(change),
                ProjectCatalogEntryKindLabel(change == null ? string.Empty : change.EntryKind),
                change == null ? string.Empty : change.CategoryName ?? string.Empty,
                change == null ? string.Empty : change.FamilyName ?? string.Empty,
                change == null ? string.Empty : change.TypeName ?? string.Empty,
                change == null ? string.Empty : change.TypeClassName ?? string.Empty,
                change != null && string.Equals(change.Attribution, "KnownBrowser", StringComparison.OrdinalIgnoreCase) ? T("Browser operation matched / actor unproven", "Browser 작업 기록 일치 / 작업자 미확정") : T("External / untracked", "외부 / 미추적"),
                change == null ? string.Empty : change.OperationUser ?? string.Empty,
                change == null ? string.Empty : change.OperationAtUtc ?? string.Empty
            });
        }
        if (rows.Count == 0)
        {
            rows.Add(new List<string> { T("Summary", "요약"), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, ProjectCatalogStatusTitle(state), string.Empty, state.CheckedAtUtc ?? string.Empty });
        }
        ShowDashboardResultWithExcelExport(
            message.ToString().Trim(),
            caption,
            error ? MessageBoxIcon.Error : (warning ? MessageBoxIcon.Warning : MessageBoxIcon.Information),
            "KKY-FamilyBrowser-Project-Catalog",
            "ProjectCatalog",
            headers,
            rows);
    }

    private sealed class TrackedProjectHistorySelectionHtmlForm : System.Windows.Forms.Form
    {
        private readonly bool _isKorean;
        private readonly List<FamilyBrowserTrackedProjectHistorySummary> _projects;
        private readonly string _activeProjectIdentity;
        private readonly WebBrowser _browser;

        public string SelectedProjectIdentity { get; private set; }

        public TrackedProjectHistorySelectionHtmlForm(
            bool isKorean,
            IEnumerable<FamilyBrowserTrackedProjectHistorySummary> projects,
            string activeProjectIdentity)
        {
            _isKorean = isKorean;
            _projects = (projects ?? Enumerable.Empty<FamilyBrowserTrackedProjectHistorySummary>())
                .Where(delegate(FamilyBrowserTrackedProjectHistorySummary project) { return project != null && !string.IsNullOrWhiteSpace(project.ProjectIdentityPath); })
                .ToList();
            _activeProjectIdentity = activeProjectIdentity ?? string.Empty;
            SelectedProjectIdentity = string.Empty;
            Text = Tx("All Tracked Project History", "전체 추적 프로젝트 이력");
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = true;
            ShowInTaskbar = false;
            MinimumSize = new Size(920, 620);
            System.Drawing.Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
            ClientSize = new Size(Math.Min(1180, Math.Max(920, workingArea.Width - 220)), Math.Min(760, Math.Max(620, workingArea.Height - 180)));
            _browser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                ScriptErrorsSuppressed = true,
                AllowNavigation = true,
                IsWebBrowserContextMenuEnabled = false,
                WebBrowserShortcutsEnabled = true
            };
            _browser.Navigating += BrowserNavigating;
            Controls.Add(_browser);
            RenderHtml();
        }

        private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (e == null || e.Url == null)
            {
                return;
            }
            string action = e.Url.AbsoluteUri ?? string.Empty;
            if (!action.StartsWith("kkyfb-history:", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            e.Cancel = true;
            string payload = action.Substring("kkyfb-history:".Length);
            if (string.Equals(payload, "close", StringComparison.OrdinalIgnoreCase))
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }
            if (!payload.StartsWith("select/", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            int index;
            if (!int.TryParse(payload.Substring("select/".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out index) || index < 0 || index >= _projects.Count)
            {
                return;
            }
            SelectedProjectIdentity = _projects[index].ProjectIdentityPath ?? string.Empty;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void RenderHtml()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<!doctype html><html><head><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/><meta charset=\"utf-8\"/>");
            sb.AppendLine("<style>html,body{margin:0;width:100%;height:100%;overflow:hidden;background:#f5f7fb;color:#17233b;font-family:'Malgun Gothic','Segoe UI',Arial,sans-serif;font-size:14px}*{box-sizing:border-box}.shell{height:100%;display:flex;flex-direction:column}.head{background:#0b1730;color:#fff;border-bottom:4px solid #2f6bff;padding:18px 22px}.head strong{display:block;font-size:22px}.head span{display:block;margin-top:5px;color:#bdc9df}.scope-note{margin:12px 18px 0;padding:10px 12px;border:1px solid #c9d7f0;border-left:4px solid #2f6bff;border-radius:6px;background:#eef4ff;color:#334968;line-height:1.5}.scope-note strong{display:block;margin-bottom:2px;color:#174ebd}.tools{display:flex;gap:10px;align-items:center;padding:12px 18px;background:#fff;border-bottom:1px solid #d8e0ef}.search{flex:1;padding:10px 12px;border:1px solid #b8c5db;border-radius:6px}.count{color:#60708e}.list{flex:1;overflow:auto;padding:14px 18px}table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d8e0ef}th{position:sticky;top:0;background:#eaf0fb;color:#142442;text-align:left;padding:10px;border-bottom:1px solid #c8d4e8}td{padding:10px;border-bottom:1px solid #e7ecf5;vertical-align:middle}tr:hover td{background:#f3f7ff}.project{font-weight:800}.path{display:block;max-width:520px;margin-top:3px;color:#667692;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.metric{white-space:nowrap;text-align:right}.badge{display:inline-block;margin-left:7px;padding:2px 7px;border-radius:4px;background:#e8f0ff;color:#225ee5;font-size:11px;font-weight:800}.open{display:inline-block;padding:8px 12px;border-radius:6px;background:#2f6bff;color:#fff;text-decoration:none;font-weight:800;white-space:nowrap}.foot{padding:12px 18px;text-align:right;background:#fff;border-top:1px solid #d8e0ef}.close{display:inline-block;padding:9px 14px;border:1px solid #b8c5db;border-radius:6px;color:#17233b;text-decoration:none;font-weight:800}.empty{padding:60px;text-align:center;color:#60708e}.hidden{display:none}</style>");
            sb.AppendLine("<script>function filterRows(){var q=(document.getElementById('search').value||'').toLowerCase(),r=document.getElementsByName('projectRow'),n=0;for(var i=0;i<r.length;i++){var ok=(r[i].getAttribute('data-search')||'').toLowerCase().indexOf(q)>=0;r[i].className=ok?'':'hidden';if(ok)n++;}document.getElementById('shown').innerText=n;}</script></head>");
            sb.AppendLine("<body onload=\"filterRows()\"><div class=\"shell\"><div class=\"head\"><strong>" + Html(Tx("All Tracked Project History", "전체 추적 프로젝트 이력")) + "</strong><span>" + Html(Tx("Choose a project to review confirmed history and local Save records waiting for upload or synchronization.", "프로젝트를 선택하면 확정 이력과 업로드 또는 동기화를 기다리는 로컬 저장 기록을 함께 확인할 수 있습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"scope-note\"><strong>" + Html(Tx("Tracking scope", "이력 관리 대상")) + "</strong>" + Html(Tx("Only RVT files registered in Permissions / Guard with Element Change Tracking enabled are recorded. Existing history remains available after a file is removed from the policy.", "권한 / 차단에 등록되고 요소 변경 추적이 체크된 RVT 파일만 새 이력이 기록됩니다. 정책에서 파일을 해제해도 기존 이력은 계속 확인할 수 있습니다.")) + "</div>");
            sb.AppendLine("<div class=\"tools\"><input id=\"search\" class=\"search\" onkeyup=\"filterRows()\" placeholder=\"" + Attr(Tx("Search project name or path", "프로젝트명 또는 경로 검색")) + "\"/><span class=\"count\"><b id=\"shown\">0</b> / " + _projects.Count.ToString(CultureInfo.InvariantCulture) + "</span></div><div class=\"list\">");
            if (_projects.Count == 0)
            {
                sb.AppendLine("<div class=\"empty\">" + Html(Tx("No tracked project history is available yet.", "아직 확인할 수 있는 프로젝트 추적 이력이 없습니다.")) + "</div>");
            }
            else
            {
                sb.AppendLine("<table><thead><tr><th>" + Html(Tx("Project", "프로젝트")) + "</th><th>" + Html(Tx("Last activity", "최근 기록")) + "</th><th class=\"metric\">" + Html(Tx("Confirmed", "확정")) + "</th><th class=\"metric\">" + Html(Tx("Pending", "대기")) + "</th><th class=\"metric\">" + Html(Tx("Created / Modified / Deleted", "생성 / 수정 / 삭제")) + "</th><th></th></tr></thead><tbody>");
                for (int i = 0; i < _projects.Count; i++)
                {
                    FamilyBrowserTrackedProjectHistorySummary project = _projects[i];
                    bool active = string.Equals(project.ProjectComparableIdentity ?? string.Empty, _activeProjectIdentity, StringComparison.OrdinalIgnoreCase);
                    string name = string.IsNullOrWhiteSpace(project.ProjectTitle) ? System.IO.Path.GetFileNameWithoutExtension(project.ProjectIdentityPath ?? string.Empty) : project.ProjectTitle;
                    string search = name + " " + (project.ProjectIdentityPath ?? string.Empty);
                    int pending = project.UploadPendingCommitCount + project.LocalSavePendingCommitCount;
                    sb.AppendLine("<tr name=\"projectRow\" data-search=\"" + Attr(search) + "\"><td><span class=\"project\">" + Html(name) + "</span>" + (active ? "<span class=\"badge\">" + Html(Tx("Current", "현재")) + "</span>" : string.Empty) + "<span class=\"path\" title=\"" + Attr(project.ProjectIdentityPath ?? string.Empty) + "\">" + Html(project.ProjectIdentityPath ?? string.Empty) + "</span></td><td>" + Html(FormatProjectCatalogDate(project.LastActivityAtUtc)) + "</td><td class=\"metric\">" + project.ConfirmedCommitCount.ToString(CultureInfo.InvariantCulture) + "</td><td class=\"metric\">" + pending.ToString(CultureInfo.InvariantCulture) + "</td><td class=\"metric\">" + project.CreatedCount.ToString(CultureInfo.InvariantCulture) + " / " + project.ModifiedCount.ToString(CultureInfo.InvariantCulture) + " / " + project.DeletedCount.ToString(CultureInfo.InvariantCulture) + "</td><td><a class=\"open\" href=\"kkyfb-history:select/" + i.ToString(CultureInfo.InvariantCulture) + "\">" + Html(Tx("Open history", "이력 열기")) + "</a></td></tr>");
                }
                sb.AppendLine("</tbody></table>");
            }
            sb.AppendLine("</div><div class=\"foot\"><a class=\"close\" href=\"kkyfb-history:close\">" + Html(Tx("Close", "닫기")) + "</a></div></div></body></html>");
            _browser.DocumentText = sb.ToString();
        }

        private string Tx(string englishText, string koreanText)
        {
            return _isKorean ? koreanText : englishText;
        }
    }

    private sealed class TrackedProjectElementHistoryHtmlForm : System.Windows.Forms.Form
    {
        private readonly bool _isKorean;
        private readonly string _caption;
        private readonly string _summary;
        private readonly IList<string> _headers;
        private readonly IList<List<string>> _rows;
        private readonly int _createdCount;
        private readonly int _modifiedCount;
        private readonly int _deletedCount;
        private readonly int _confirmedCount;
        private readonly int _pendingCount;
        private readonly WebBrowser _browser;

        public event EventHandler ExportRequested;

        public TrackedProjectElementHistoryHtmlForm(
            bool isKorean,
            string caption,
            string summary,
            IList<string> headers,
            IList<List<string>> rows,
            int createdCount,
            int modifiedCount,
            int deletedCount,
            int confirmedCount,
            int pendingCount)
        {
            _isKorean = isKorean;
            _caption = caption ?? string.Empty;
            _summary = summary ?? string.Empty;
            _headers = headers ?? new List<string>();
            _rows = rows ?? new List<List<string>>();
            _createdCount = createdCount;
            _modifiedCount = modifiedCount;
            _deletedCount = deletedCount;
            _confirmedCount = confirmedCount;
            _pendingCount = pendingCount;
            Text = _caption;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = true;
            ShowInTaskbar = false;
            MinimumSize = new Size(980, 640);
            System.Drawing.Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
            ClientSize = new Size(Math.Min(1440, Math.Max(1080, workingArea.Width - 120)), Math.Min(900, Math.Max(700, workingArea.Height - 100)));
            _browser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                ScriptErrorsSuppressed = true,
                AllowNavigation = true,
                IsWebBrowserContextMenuEnabled = false,
                WebBrowserShortcutsEnabled = true
            };
            _browser.Navigating += BrowserNavigating;
            Controls.Add(_browser);
            RenderHtml();
        }

        private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (e == null || e.Url == null)
            {
                return;
            }
            string action = e.Url.AbsoluteUri ?? string.Empty;
            if (!action.StartsWith("kkyfb-element-history:", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            e.Cancel = true;
            string payload = action.Substring("kkyfb-element-history:".Length);
            if (string.Equals(payload, "close", StringComparison.OrdinalIgnoreCase))
            {
                Close();
            }
            else if (string.Equals(payload, "export", StringComparison.OrdinalIgnoreCase))
            {
                EventHandler handler = ExportRequested;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            }
        }

        private void RenderHtml()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<!doctype html><html><head><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/><meta charset=\"utf-8\"/>");
            sb.AppendLine("<style>html,body{margin:0;width:100%;height:100%;overflow:hidden;background:#f5f7fb;color:#17233b;font-family:'Malgun Gothic','Segoe UI',Arial,sans-serif;font-size:13px}*{box-sizing:border-box}.shell{height:100%;display:flex;flex-direction:column}.head{background:#0b1730;color:#fff;border-bottom:4px solid #2f6bff;padding:14px 20px}.head strong{font-size:20px}.scope-note{margin:10px 16px 0;padding:9px 11px;border:1px solid #c9d7f0;border-left:4px solid #2f6bff;border-radius:6px;background:#eef4ff;color:#334968;line-height:1.45}.scope-note strong{margin-right:8px;color:#174ebd}.metrics{display:flex;gap:8px;flex-wrap:wrap;padding:12px 16px;background:#fff;border-bottom:1px solid #d8e0ef}.metric{min-width:135px;border:1px solid #d8e0ef;border-left:4px solid #2f6bff;border-radius:6px;padding:8px 10px}.metric span{display:block;color:#667692;font-size:11px}.metric b{display:block;margin-top:2px;font-size:19px}.metric.created{border-left-color:#16835f}.metric.deleted{border-left-color:#cf3f35}.metric.pending{border-left-color:#d89b12}.tools{display:flex;gap:7px;align-items:center;padding:10px 16px;background:#fff;border-bottom:1px solid #d8e0ef}.search{flex:1;min-width:260px;padding:9px 11px;border:1px solid #b8c5db;border-radius:6px}.filter,.action{display:inline-block;padding:8px 11px;border:1px solid #b8c5db;border-radius:6px;background:#fff;color:#17233b;text-decoration:none;font-weight:800;white-space:nowrap}.filter.active{background:#2f6bff;border-color:#2f6bff;color:#fff}.tablewrap{flex:1;overflow:auto;padding:0 16px 14px}table{border-collapse:collapse;min-width:1900px;width:100%;background:#fff;border:1px solid #d8e0ef}th{position:sticky;top:0;z-index:2;background:#eaf0fb;color:#142442;text-align:left;padding:9px;border-bottom:1px solid #c8d4e8;white-space:nowrap}td{max-width:330px;padding:8px 9px;border-bottom:1px solid #e7ecf5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}tr:hover td{background:#f3f7ff}tr.kind-deleted td:nth-child(3){color:#b82d25;font-weight:800}tr.kind-created td:nth-child(3){color:#0d7553;font-weight:800}tr.kind-modified td:nth-child(3){color:#225ee5;font-weight:800}.hidden{display:none}.summary{display:none;max-height:170px;overflow:auto;margin:0 16px 10px;padding:11px 13px;border:1px solid #d8e0ef;background:#fff;line-height:1.55;color:#4c5d7a}.summary.open{display:block}.foot{display:flex;justify-content:space-between;align-items:center;padding:11px 16px;background:#fff;border-top:1px solid #d8e0ef}.foot span{color:#667692}.action.primary{background:#2f6bff;border-color:#2f6bff;color:#fff}</style>");
            sb.AppendLine("<script>var activeKind='all';function setKind(k){activeKind=k;var b=document.getElementsByName('kindButton');for(var i=0;i<b.length;i++){b[i].className='filter'+(b[i].getAttribute('data-kind')==k?' active':'');}filterRows();}function filterRows(){var q=(document.getElementById('historySearch').value||'').toLowerCase(),r=document.getElementsByName('historyRow'),n=0;for(var i=0;i<r.length;i++){var kind=r[i].getAttribute('data-kind')||'',ok=(activeKind=='all'||kind==activeKind)&&((r[i].getAttribute('data-search')||'').toLowerCase().indexOf(q)>=0);r[i].className='kind-'+kind+(ok?'':' hidden');if(ok)n++;}document.getElementById('visibleRows').innerText=n;}function toggleSummary(){var s=document.getElementById('summary');s.className=s.className=='summary'?'summary open':'summary';}</script></head>");
            sb.AppendLine("<body onload=\"setKind('all')\"><div class=\"shell\"><div class=\"head\"><strong>" + Html(_caption) + "</strong></div>");
            sb.AppendLine("<div class=\"scope-note\"><strong>" + Html(Tx("Tracking scope", "이력 관리 대상")) + "</strong>" + Html(Tx("Only RVT files registered in Permissions / Guard with Element Change Tracking enabled are recorded. Existing history remains available after a file is removed from the policy.", "권한 / 차단에 등록되고 요소 변경 추적이 체크된 RVT 파일만 새 이력이 기록됩니다. 정책에서 파일을 해제해도 기존 이력은 계속 확인할 수 있습니다.")) + "</div><div class=\"metrics\">");
            AppendMetric(sb, Tx("Created", "생성"), _createdCount, "created");
            AppendMetric(sb, Tx("Modified", "수정"), _modifiedCount, string.Empty);
            AppendMetric(sb, Tx("Deleted", "삭제"), _deletedCount, "deleted");
            AppendMetric(sb, Tx("Confirmed records", "확정 기록"), _confirmedCount, string.Empty);
            AppendMetric(sb, Tx("Pending records", "대기 기록"), _pendingCount, "pending");
            sb.AppendLine("</div><div class=\"tools\"><input id=\"historySearch\" class=\"search\" onkeyup=\"filterRows()\" placeholder=\"" + Attr(Tx("Search user, ID, category, family, type, or transaction", "사용자, ID, 카테고리, 패밀리, 타입, 트랜잭션 검색")) + "\"/><a name=\"kindButton\" data-kind=\"all\" class=\"filter active\" href=\"javascript:setKind('all')\">" + Html(Tx("All", "전체")) + "</a><a name=\"kindButton\" data-kind=\"created\" class=\"filter\" href=\"javascript:setKind('created')\">" + Html(Tx("Created", "생성")) + "</a><a name=\"kindButton\" data-kind=\"modified\" class=\"filter\" href=\"javascript:setKind('modified')\">" + Html(Tx("Modified", "수정")) + "</a><a name=\"kindButton\" data-kind=\"deleted\" class=\"filter\" href=\"javascript:setKind('deleted')\">" + Html(Tx("Deleted", "삭제")) + "</a><a class=\"filter\" href=\"javascript:toggleSummary()\">" + Html(Tx("Summary", "요약")) + "</a></div>");
            sb.AppendLine("<div id=\"summary\" class=\"summary\">" + Html(_summary).Replace("\r\n", "<br/>").Replace("\n", "<br/>") + "</div><div class=\"tablewrap\"><table><thead><tr>");
            foreach (string header in _headers)
            {
                sb.AppendLine("<th>" + Html(header ?? string.Empty) + "</th>");
            }
            sb.AppendLine("</tr></thead><tbody>");
            foreach (List<string> row in _rows)
            {
                string kind = ResolveRowKind(row);
                string search = string.Join(" ", row ?? new List<string>());
                sb.AppendLine("<tr name=\"historyRow\" class=\"kind-" + Attr(kind) + "\" data-kind=\"" + Attr(kind) + "\" data-search=\"" + Attr(search) + "\">");
                foreach (string cell in row ?? new List<string>())
                {
                    sb.AppendLine("<td title=\"" + Attr(cell ?? string.Empty) + "\">" + Html(cell ?? string.Empty) + "</td>");
                }
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody></table></div><div class=\"foot\"><span><b id=\"visibleRows\">0</b> / " + _rows.Count.ToString(CultureInfo.InvariantCulture) + " " + Html(Tx("rows", "개")) + "</span><div><a class=\"action primary\" href=\"kkyfb-element-history:export\">" + Html(Tx("Export Excel", "Excel 내보내기")) + "</a> <a class=\"action\" href=\"kkyfb-element-history:close\">" + Html(Tx("Close", "닫기")) + "</a></div></div></div></body></html>");
            _browser.DocumentText = sb.ToString();
        }

        private void AppendMetric(StringBuilder sb, string label, int value, string cssClass)
        {
            sb.AppendLine("<div class=\"metric " + Attr(cssClass ?? string.Empty) + "\"><span>" + Html(label) + "</span><b>" + value.ToString(CultureInfo.InvariantCulture) + "</b></div>");
        }

        private string ResolveRowKind(IList<string> row)
        {
            string value = row != null && row.Count > 2 ? row[2] ?? string.Empty : string.Empty;
            if (value.IndexOf(Tx("Deleted", "삭제"), StringComparison.OrdinalIgnoreCase) >= 0) return "deleted";
            if (value.IndexOf(Tx("Created", "생성"), StringComparison.OrdinalIgnoreCase) >= 0) return "created";
            if (value.IndexOf(Tx("Modified", "수정"), StringComparison.OrdinalIgnoreCase) >= 0) return "modified";
            return "other";
        }

        private string Tx(string englishText, string koreanText)
        {
            return _isKorean ? koreanText : englishText;
        }
    }

    private void ShowProjectElementChangeHistory()
    {
        if (!EnsurePermission("ManagePolicy", T("View Project Element Change History", "프로젝트 요소 변경 이력 보기")))
        {
            return;
        }

        Document doc = GetActiveDocument();
        string projectIdentity = doc == null ? string.Empty : ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
        if (string.IsNullOrWhiteSpace(projectIdentity))
        {
            ShowDashboardMessage(
                this,
                T(
                    "Save the current project first. Element change history is separated by the saved project or central-model identity.",
                    "현재 프로젝트를 먼저 저장하세요. 요소 변경 이력은 저장된 프로젝트 또는 센트럴 모델 경로를 기준으로 구분됩니다."),
                T("Project Element Change History", "프로젝트 요소 변경 이력"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        ShowSelectedProjectElementChangeHistory(projectIdentity, doc == null ? string.Empty : doc.Title ?? string.Empty);
    }

    private void ShowAllProjectElementChangeHistory()
    {
        if (!EnsurePermission("ManagePolicy", T("View All Project Element Change History", "전체 프로젝트 요소 변경 이력 보기")))
        {
            return;
        }
        if (_projectElementHistoryLoadPending)
        {
            _statusMessage = T("Project history is already loading.", "프로젝트 이력을 이미 불러오는 중입니다.");
            RefreshDocumentShellOnly();
            return;
        }
        Document activeDocument = GetActiveDocument();
        string activeIdentity = activeDocument == null ? string.Empty : ProjectSnapshotStore.ResolveProjectIdentityPath(activeDocument);
        string activeStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(activeIdentity);
        string activeProjectTitle = activeDocument == null ? string.Empty : activeDocument.Title ?? string.Empty;
        FamilyBrowserStandardPolicy activePolicy = _standardPolicy ?? LoadStandardPolicy();
        bool activeProjectTrackingEnabled = activeDocument != null &&
            FamilyBrowserStandardPolicyStore.IsProjectElementChangeTrackingEnabled(activePolicy) &&
            FamilyBrowserSecurityPolicyService.IsProjectElementTrackingScopeEnabled(activePolicy, CurrentProjectPolicyContext());
        string loadingWorkspaceRoot = _workspaceRoot ?? string.Empty;
        _projectElementHistoryLoadPending = true;
        _statusMessage = T("Loading all project history in the background...", "전체 프로젝트 이력을 백그라운드에서 불러오는 중입니다...");
        RefreshDocumentShellOnly();

        Task<List<FamilyBrowserTrackedProjectHistorySummary>> loadTask = Task.Run(delegate
        {
            return FamilyBrowserTrackingPersistenceService.LoadTrackedProjectHistorySummaries(loadingWorkspaceRoot);
        });
        loadTask.ContinueWith(delegate(Task<List<FamilyBrowserTrackedProjectHistorySummary>> completed)
        {
            if (IsDisposed)
            {
                return;
            }
            try
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    CompleteAllProjectElementChangeHistoryLoad(
                        completed,
                        loadingWorkspaceRoot,
                        activeIdentity,
                        activeStableIdentity,
                        activeProjectTitle,
                        activeProjectTrackingEnabled);
                });
            }
            catch (InvalidOperationException)
            {
            }
        });
    }

    private void CompleteAllProjectElementChangeHistoryLoad(
        Task<List<FamilyBrowserTrackedProjectHistorySummary>> loadTask,
        string loadingWorkspaceRoot,
        string activeIdentity,
        string activeStableIdentity,
        string activeProjectTitle,
        bool activeProjectTrackingEnabled)
    {
        _projectElementHistoryLoadPending = false;
        if (IsDisposed)
        {
            return;
        }
        if (!string.Equals(_workspaceRoot ?? string.Empty, loadingWorkspaceRoot ?? string.Empty, StringComparison.OrdinalIgnoreCase))
        {
            _statusMessage = T("The management folder changed while history was loading. Open the history again.", "이력을 불러오는 동안 관리폴더가 변경되었습니다. 이력을 다시 열어주세요.");
            RefreshDocumentShellOnly();
            return;
        }
        if (loadTask == null || loadTask.IsCanceled || loadTask.IsFaulted)
        {
            string detail = loadTask != null && loadTask.Exception != null
                ? loadTask.Exception.GetBaseException().Message
                : T("The background history read was canceled.", "백그라운드 이력 읽기가 취소되었습니다.");
            _statusMessage = T("Project history could not be loaded.", "프로젝트 이력을 불러오지 못했습니다.");
            RefreshDocumentShellOnly();
            ShowDashboardMessage(
                this,
                _statusMessage + Environment.NewLine + Environment.NewLine + detail,
                T("All Project History", "전체 프로젝트 이력"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        List<FamilyBrowserTrackedProjectHistorySummary> projects = loadTask.Result ?? new List<FamilyBrowserTrackedProjectHistorySummary>();
        if (activeProjectTrackingEnabled && !string.IsNullOrWhiteSpace(activeStableIdentity) && !projects.Any(delegate(FamilyBrowserTrackedProjectHistorySummary project)
        {
            return project != null && string.Equals(project.ProjectComparableIdentity, activeStableIdentity, StringComparison.OrdinalIgnoreCase);
        }))
        {
            projects.Insert(0, new FamilyBrowserTrackedProjectHistorySummary
            {
                ProjectIdentityPath = activeIdentity,
                ProjectComparableIdentity = activeStableIdentity,
                ProjectTitle = activeProjectTitle
            });
        }
        _statusMessage = T("All project history loaded: ", "전체 프로젝트 이력 불러오기 완료: ") + projects.Count.ToString(CultureInfo.InvariantCulture);
        RefreshDocumentShellOnly();
        using (TrackedProjectHistorySelectionHtmlForm dialog = new TrackedProjectHistorySelectionHtmlForm(IsKorean(), projects, activeStableIdentity))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedProjectIdentity))
            {
                return;
            }
            FamilyBrowserTrackedProjectHistorySummary selected = projects.FirstOrDefault(delegate(FamilyBrowserTrackedProjectHistorySummary project)
            {
                return project != null && string.Equals(project.ProjectIdentityPath ?? string.Empty, dialog.SelectedProjectIdentity, StringComparison.OrdinalIgnoreCase);
            });
            ShowSelectedProjectElementChangeHistory(
                dialog.SelectedProjectIdentity,
                selected == null ? string.Empty : selected.ProjectTitle ?? string.Empty);
        }
    }

    private void ShowSelectedProjectElementChangeHistory(string projectIdentity, string projectTitle)
    {
        const int commitLimit = 200;
        if (_projectElementHistoryLoadPending)
        {
            _statusMessage = T("Project history is already loading.", "프로젝트 이력을 이미 불러오는 중입니다.");
            RefreshDocumentShellOnly();
            return;
        }
        string loadingWorkspaceRoot = _workspaceRoot ?? string.Empty;
        _projectElementHistoryLoadPending = true;
        _statusMessage = T("Loading project history in the background...", "프로젝트 이력을 백그라운드에서 불러오는 중입니다...");
        RefreshDocumentShellOnly();

        Task<ProjectElementHistoryLoadBundle> loadTask = Task.Run(delegate
        {
            return LoadProjectElementHistoryBundle(loadingWorkspaceRoot, projectIdentity, commitLimit);
        });
        loadTask.ContinueWith(delegate(Task<ProjectElementHistoryLoadBundle> completed)
        {
            if (IsDisposed)
            {
                return;
            }
            try
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    CompleteSelectedProjectElementChangeHistoryLoad(
                        completed,
                        loadingWorkspaceRoot,
                        projectIdentity,
                        projectTitle);
                });
            }
            catch (InvalidOperationException)
            {
            }
        });
    }

    private static ProjectElementHistoryLoadBundle LoadProjectElementHistoryBundle(
        string workspaceRoot,
        string projectIdentity,
        int commitLimit)
    {
        ProjectElementHistoryLoadBundle bundle = new ProjectElementHistoryLoadBundle();
        bundle.History = FamilyBrowserTrackingPersistenceService
            .LoadImmutableElementChangeCommitResult(workspaceRoot, projectIdentity, commitLimit);
        bundle.UploadPendingHistory = FamilyBrowserTrackingPersistenceService
            .LoadPendingElementChangeCommitResult(workspaceRoot, projectIdentity, commitLimit);
        bundle.CheckpointHistory = FamilyBrowserTrackingPersistenceService
            .LoadElementSessionCheckpointHistory(workspaceRoot, projectIdentity, commitLimit);
        bundle.LocalSyncPendingStatus = FamilyBrowserTrackingPersistenceService
            .GetPendingElementSessionCheckpointStatus(workspaceRoot, projectIdentity);
        if (!bundle.LocalSyncPendingStatus.LockUnavailable)
        {
            bundle.InvalidLocalSyncPending = FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount();
            bundle.MismatchedLocalSyncPending = FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount(workspaceRoot);
        }
        bundle.LatestPolicyChange = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries(workspaceRoot, 2000)
            .Where(delegate(FamilyBrowserOperationLogEntry entry)
            {
                return entry != null && string.Equals(entry.OperationKind, "ProjectElementChangeTrackingPolicy", StringComparison.OrdinalIgnoreCase);
            })
            .OrderByDescending(delegate(FamilyBrowserOperationLogEntry entry)
            {
                return string.IsNullOrWhiteSpace(entry.CommittedAtUtc) ? entry.RecordedAtUtc ?? string.Empty : entry.CommittedAtUtc;
            }, StringComparer.Ordinal)
            .FirstOrDefault();
        return bundle;
    }

    private void CompleteSelectedProjectElementChangeHistoryLoad(
        Task<ProjectElementHistoryLoadBundle> loadTask,
        string loadingWorkspaceRoot,
        string projectIdentity,
        string projectTitle)
    {
        _projectElementHistoryLoadPending = false;
        if (IsDisposed)
        {
            return;
        }
        if (!string.Equals(_workspaceRoot ?? string.Empty, loadingWorkspaceRoot ?? string.Empty, StringComparison.OrdinalIgnoreCase))
        {
            _statusMessage = T("The management folder changed while history was loading. Open the history again.", "이력을 불러오는 동안 관리폴더가 변경되었습니다. 이력을 다시 열어주세요.");
            RefreshDocumentShellOnly();
            return;
        }
        if (loadTask == null || loadTask.IsCanceled || loadTask.IsFaulted)
        {
            string detail = loadTask != null && loadTask.Exception != null
                ? loadTask.Exception.GetBaseException().Message
                : T("The background history read was canceled.", "백그라운드 이력 읽기가 취소되었습니다.");
            _statusMessage = T("Project history could not be loaded.", "프로젝트 이력을 불러오지 못했습니다.");
            RefreshDocumentShellOnly();
            ShowDashboardMessage(
                this,
                _statusMessage + Environment.NewLine + Environment.NewLine + detail,
                T("Project Element Change History", "프로젝트 요소 변경 이력"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }
        _statusMessage = T("Project history loaded.", "프로젝트 이력을 불러왔습니다.");
        RefreshDocumentShellOnly();
        RenderSelectedProjectElementChangeHistory(projectIdentity, projectTitle, loadTask.Result);
    }

    private void RenderSelectedProjectElementChangeHistory(
        string projectIdentity,
        string projectTitle,
        ProjectElementHistoryLoadBundle loadBundle)
    {

        const int commitLimit = 200;
        const int rowLimit = 5000;
        loadBundle = loadBundle ?? new ProjectElementHistoryLoadBundle();
        FamilyBrowserElementChangeHistoryLoadResult history = loadBundle.History ?? new FamilyBrowserElementChangeHistoryLoadResult();
        FamilyBrowserElementChangeHistoryLoadResult uploadPendingHistory = loadBundle.UploadPendingHistory ?? new FamilyBrowserElementChangeHistoryLoadResult();
        FamilyBrowserElementSessionCheckpointHistoryLoadResult checkpointHistory = loadBundle.CheckpointHistory ?? new FamilyBrowserElementSessionCheckpointHistoryLoadResult();
        HashSet<string> immutableEntryIds = new HashSet<string>((history.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        HashSet<string> uploadPendingEntryIds = new HashSet<string>((uploadPendingHistory.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        HashSet<string> checkpointEntryIds = new HashSet<string>((checkpointHistory.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        List<FamilyBrowserElementChangeCommit> commits = (history.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Concat(uploadPendingHistory.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Concat(checkpointHistory.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .GroupBy(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
            .Select(delegate(IGrouping<string, FamilyBrowserElementChangeCommit> group) { return group.First(); })
            .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit)
            {
                string timestamp = string.IsNullOrWhiteSpace(commit.PublishedAtUtc)
                    ? (string.IsNullOrWhiteSpace(commit.LocalSaveProtectedAtUtc) ? commit.CommittedAtUtc : commit.LocalSaveProtectedAtUtc)
                    : commit.PublishedAtUtc;
                DateTime parsed;
                return DateTime.TryParse(timestamp ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed) ? parsed : DateTime.MinValue;
            })
            .Take(commitLimit)
            .ToList();

        Dictionary<FamilyBrowserElementChangeCommit, FamilyBrowserElementHistoryProjectionCounts> projectionCountsByCommit = commits.ToDictionary(
            delegate(FamilyBrowserElementChangeCommit commit) { return commit; },
            delegate(FamilyBrowserElementChangeCommit commit)
            {
                return FamilyBrowserElementHistoryProjectionPolicy.CountUserFacingChanges(
                    commit.Changes ?? new List<FamilyBrowserElementChangeItem>());
            });
        Dictionary<FamilyBrowserElementChangeCommit, List<FamilyBrowserElementChangeItem>> visibleChangesByCommit = commits.ToDictionary(
            delegate(FamilyBrowserElementChangeCommit commit) { return commit; },
            delegate(FamilyBrowserElementChangeCommit commit)
            {
                return (commit.Changes ?? new List<FamilyBrowserElementChangeItem>())
                    .Where(IsUserFacingProjectElementChange)
                    .ToList();
            });
        int hiddenAuxiliaryRows = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit)
        {
            return projectionCountsByCommit[commit].HiddenAuxiliaryCount;
        });
        int hiddenUnresolvedTransientRows = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit)
        {
            return projectionCountsByCommit[commit].HiddenUnresolvedTransientCount;
        });
        int created = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit) { return projectionCountsByCommit[commit].CreatedCount; });
        int modified = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit) { return projectionCountsByCommit[commit].ModifiedCount; });
        int deleted = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit) { return projectionCountsByCommit[commit].DeletedCount; });
        int transientCreatedDeleted = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit) { return projectionCountsByCommit[commit].TransientCreatedDeletedCount; });
        int overlap = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit) { return visibleChangesByCommit[commit].Count(delegate(FamilyBrowserElementChangeItem item) { return item.ExternalUpdateOverlap; }); });
        int coverageGapOnlyCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return commit.CoverageGapOnly; });
        int externalRebaseGapCommits = commits.Count(ProjectElementHasExternalRebaseGap);
        int eventReadGapCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EventReadFailureCount > 0 || (commit.CoverageNote ?? string.Empty).IndexOf("DocumentChanged ID or operation-metadata", StringComparison.OrdinalIgnoreCase) >= 0; });
        int commitBoundaryGapCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return commit.CommitBoundaryReadFailureCount > 0 || string.Equals(commit.AttributionConfidence, "ClientObservedWithCommitBoundaryGap", StringComparison.OrdinalIgnoreCase) || (commit.CoverageNote ?? string.Empty).IndexOf("Save or Save As completion boundaries", StringComparison.OrdinalIgnoreCase) >= 0; });
        int ambiguousEventCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return commit.UnmatchedUndoCount > 0 || commit.UnmatchedRedoCount > 0; });
        int identityGapCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return string.IsNullOrWhiteSpace(commit.RevitUserName); });
        int policyFallbackCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return string.Equals(commit.PolicyValidationState, "LastKnownEnabled", StringComparison.OrdinalIgnoreCase); });
        int policyDisableDeferredCommits = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return string.Equals(commit.PolicyValidationState, "DisablePendingCommit", StringComparison.OrdinalIgnoreCase); });
        int totalValidRecordCount = history.TotalValidRecordCount +
            uploadPendingHistory.TotalValidRecordCount +
            checkpointHistory.TotalValidRecordCount;
        bool commitHistoryTruncated = totalValidRecordCount > commits.Count;
        int availableRowsInLoadedCommits = commits.Sum(delegate(FamilyBrowserElementChangeCommit commit)
        {
            int changeRows = visibleChangesByCommit[commit].Count;
            return changeRows == 0 && commit.CoverageGapOnly ? 1 : changeRows;
        });
        bool rowHistoryTruncated = availableRowsInLoadedCommits > rowLimit;
        int pending = uploadPendingHistory.TotalValidRecordCount;
        FamilyBrowserElementSessionCheckpointCountResult localSyncPendingStatus =
            loadBundle.LocalSyncPendingStatus ?? new FamilyBrowserElementSessionCheckpointCountResult();
        int localSyncPending = localSyncPendingStatus.Count;
        int synchronizedHistoryPromotionPending = Math.Min(localSyncPending, Math.Max(0, localSyncPendingStatus.SynchronizationSucceededCount));
        int synchronizationPending = Math.Max(0, localSyncPending - synchronizedHistoryPromotionPending);
        int invalidLocalSyncPending = loadBundle.InvalidLocalSyncPending;
        int mismatchedLocalSyncPending = loadBundle.MismatchedLocalSyncPending;
        FamilyBrowserOperationLogEntry latestPolicyChange = loadBundle.LatestPolicyChange;

        StringBuilder message = new StringBuilder();
        message.AppendLine(T("Recent project element change history", "최근 프로젝트 요소 변경 이력"));
        message.AppendLine(T("Project: ", "프로젝트: ") + (string.IsNullOrWhiteSpace(projectTitle) ? System.IO.Path.GetFileNameWithoutExtension(projectIdentity ?? string.Empty) : projectTitle));
        message.AppendLine(T("Path: ", "경로: ") + (projectIdentity ?? string.Empty));
        message.AppendLine();
        message.AppendLine(T("History summary", "이력 요약"));
        message.AppendLine(T("Loaded commits: ", "불러온 확정 기록: ") + commits.Count.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Total valid records found: ", "확인된 전체 유효 기록: ") + totalValidRecordCount.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Created / modified / deleted: ", "생성 / 수정 / 삭제: ")
            + created.ToString(CultureInfo.InvariantCulture) + " / "
            + modified.ToString(CultureInfo.InvariantCulture) + " / "
            + deleted.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Local upload queue: ", "로컬 전송 대기: ") + pending.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Saved locally / waiting for synchronization: ", "로컬 저장 완료 / 동기화 대기: ") +
            (localSyncPendingStatus.LockUnavailable
                ? T("checkpoint busy / refresh required", "체크포인트 잠금 중 / 재확인 필요")
                : synchronizationPending.ToString(CultureInfo.InvariantCulture)));
        message.AppendLine(T("Synchronized / immutable history promotion pending: ", "동기화 완료 / 관리 이력 승격 대기: ") +
            (localSyncPendingStatus.LockUnavailable
                ? T("checkpoint busy / refresh required", "체크포인트 잠금 중 / 재확인 필요")
                : synchronizedHistoryPromotionPending.ToString(CultureInfo.InvariantCulture)));
        message.AppendLine(T("Local checkpoints requiring recovery: ", "복구가 필요한 로컬 체크포인트: ") + invalidLocalSyncPending.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Local checkpoints bound to another management folder: ", "다른 관리폴더에 묶인 로컬 체크포인트: ") + mismatchedLocalSyncPending.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Created then deleted before commit: ", "확정 전 생성 후 삭제: ") + transientCreatedDeleted.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Incoming-update overlap: ", "외부 업데이트 겹침: ") + overlap.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Internal support rows hidden from this view: ", "화면에서 제외한 내부 보조 요소 행: ") + hiddenAuxiliaryRows.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Unresolved same-boundary transient rows hidden from this view: ", "화면에서 제외한 동일 작업 경계 식별 불가 일시 요소 행: ") + hiddenUnresolvedTransientRows.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Coverage-gap records without a trustworthy element ID: ", "신뢰할 수 있는 요소 ID가 없는 관찰 공백 기록: ") + coverageGapOnlyCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits with incomplete external-update rebase: ", "외부 업데이트 재기준화 불완전 기록: ") + externalRebaseGapCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits with incomplete DocumentChanged event reads: ", "DocumentChanged 이벤트 관찰 공백 기록: ") + eventReadGapCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits spanning an unreadable Save boundary: ", "저장 경계 관찰 공백 기록: ") + commitBoundaryGapCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits with unmatched Undo/Redo callbacks: ", "Undo/Redo 연결 불확실 기록: ") + ambiguousEventCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits without a Revit username: ", "Revit 사용자명 확인 불가 기록: ") + identityGapCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits using last confirmed policy: ", "마지막 확인 정책을 사용한 기록: ") + policyFallbackCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Commits protected after a remote tracking disable: ", "원격 추적 해제 후 보호 확정한 기록: ") + policyDisableDeferredCommits.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Invalid or checksum-failed records skipped: ", "손상 또는 체크섬 불일치로 제외: ") + (history.InvalidRecordCount + uploadPendingHistory.InvalidRecordCount + checkpointHistory.InvalidRecordCount).ToString(CultureInfo.InvariantCulture));
        message.AppendLine(T("Legacy records without checksum: ", "체크섬 없는 이전 기록: ") + history.LegacyUnverifiedCount.ToString(CultureInfo.InvariantCulture));
        if (commitHistoryTruncated || rowHistoryTruncated)
        {
            message.AppendLine(T("Display limit reached: older commits or rows are not shown in this window. Exporting this view is subject to the same limit.", "표시 한도 도달: 이 창에는 더 오래된 확정 기록 또는 행이 표시되지 않습니다. 이 화면의 Excel 내보내기에도 같은 한도가 적용됩니다."));
        }
        if (history.PendingDestinationMismatchCount > 0)
        {
            message.AppendLine(T("Pending records bound to another management folder: ", "다른 관리폴더에 묶인 전송 대기 기록: ") + history.PendingDestinationMismatchCount.ToString(CultureInfo.InvariantCulture));
        }
        if (uploadPendingHistory.PendingDestinationMismatchCount > 0 || checkpointHistory.DestinationMismatchCount > 0)
        {
            message.AppendLine(T("Local records bound to another management folder: ", "다른 관리폴더에 묶인 로컬 기록: ") + (uploadPendingHistory.PendingDestinationMismatchCount + checkpointHistory.DestinationMismatchCount).ToString(CultureInfo.InvariantCulture));
        }
        if (history.PendingCorruptRecordCount > 0 || history.PendingFailedCount > 0)
        {
            message.AppendLine(T("Pending records requiring recovery: ", "복구가 필요한 전송 대기 기록: ") + Math.Max(history.PendingCorruptRecordCount, history.PendingFailedCount).ToString(CultureInfo.InvariantCulture));
        }
        if (latestPolicyChange != null)
        {
            string policyChangeAt = string.IsNullOrWhiteSpace(latestPolicyChange.CommittedAtUtc) ? latestPolicyChange.RecordedAtUtc : latestPolicyChange.CommittedAtUtc;
            message.AppendLine(T("Last tracking setting change: ", "마지막 추적 설정 변경: ")
                + FormatProjectCatalogDate(policyChangeAt) + " · "
                + (latestPolicyChange.PlannedAction ?? string.Empty) + " · "
                + (latestPolicyChange.Outcome ?? string.Empty) + " · "
                + (latestPolicyChange.UserName ?? string.Empty));
        }
        message.AppendLine();
        message.AppendLine(T(
            "This view loads up to the latest 200 committed records and 5,000 element rows. For recovered workshared activity, Committed means the successful central synchronization time and Local Save shows the successful local Save that protected the checkpoint. First/last observed times come from the editing PC, so cross-PC ordering assumes Windows clocks are synchronized. Tracking covers user-facing model elements, element types, families, materials, grids, and project/shared parameter definitions with binding metadata. Views, DataStorage, ProjectInfo, temporary negative-ID elements, and other categoryless Revit internals are intentionally excluded. Logical MEP systems, centerline graphics, cable-tray/conduit run containers, and newly captured dependent nested Family instances are also excluded. Legacy records created before dependency classification can retain nested-Family rows because their immutable payload has no parent relationship. Excel is created only when Export Excel is selected. Exact user attribution requires this add-in on every editing workstation. A checksum detects accidental corruption or unsophisticated editing; it is not a cryptographic user signature.",
            "최근 확정 기록 200개와 요소 행 5,000개까지 표시합니다. 복구된 워크셰어링 작업에서 확정 시각은 센트럴 동기화 성공 시각이고, 로컬 저장은 체크포인트를 보호한 로컬 저장 성공 시각입니다. 최초/최종 관찰 시각은 작업 PC 기준이므로 여러 PC 사이의 정확한 순서는 Windows 시간이 동기화되어 있어야 합니다. 사용자가 직접 다루는 모델 요소, 요소 타입, 패밀리, 재질, 그리드와 바인딩 정보가 포함된 프로젝트/공유 파라미터 정의를 추적합니다. View, DataStorage, ProjectInfo, 임시 음수 ID 요소, 카테고리 없는 Revit 내부 요소, 논리 MEP 시스템, 중심선 그래픽, 케이블 트레이/전선관 Run 컨테이너와 새로 수집되는 종속 하위 패밀리 인스턴스는 의도적으로 제외합니다. 종속 관계 분류 기능 이전의 불변 이력에는 부모 관계가 저장되지 않았으므로 과거 하위 패밀리 행이 남을 수 있습니다. Excel은 Excel 내보내기를 선택할 때만 생성됩니다. 정확한 사용자 귀속에는 모든 작업 PC에 이 애드인이 필요합니다. 체크섬은 우발적 손상이나 단순 수정을 감지하지만 사용자를 증명하는 전자서명은 아닙니다."));

        int previewIndex = 0;
        foreach (FamilyBrowserElementChangeCommit commit in commits)
        {
            if (commit.CoverageGapOnly && (commit.Changes == null || commit.Changes.Count == 0) && previewIndex < 20)
            {
                if (previewIndex == 0)
                {
                    message.AppendLine();
                    message.AppendLine(T("Recent changes and coverage gaps", "최근 변경 및 관찰 공백"));
                }
                previewIndex++;
                message.AppendLine(previewIndex.ToString(CultureInfo.InvariantCulture) + ": "
                    + FormatProjectCatalogDate(commit.CommittedAtUtc) + " · "
                    + T("Coverage gap / element ID unavailable", "관찰 공백 / 요소 ID 확인 불가") + " · "
                    + (string.IsNullOrWhiteSpace(commit.RevitUserName) ? commit.WindowsUserName ?? string.Empty : commit.RevitUserName));
            }
            foreach (FamilyBrowserElementChangeItem change in visibleChangesByCommit[commit])
            {
                if (previewIndex >= 20)
                {
                    break;
                }
                if (change == null)
                {
                    continue;
                }
                if (previewIndex == 0)
                {
                    message.AppendLine();
                    message.AppendLine(T("Recent changes", "최근 변경"));
                }
                previewIndex++;
                string itemName = string.Join(" / ", new[] { ProjectElementCategoryLabel(change), ProjectElementName(change), change.FamilyName, change.TypeName }
                    .Where(delegate(string value) { return !string.IsNullOrWhiteSpace(value); })
                    .Select(delegate(string value) { return value.Trim(); }));
                message.AppendLine(previewIndex.ToString(CultureInfo.InvariantCulture) + ": "
                    + FormatProjectCatalogDate(commit.CommittedAtUtc) + " · "
                    + ProjectElementChangeKindLabel(change.ChangeKind) + " · "
                    + (string.IsNullOrWhiteSpace(itemName) ? T("Element ", "요소 ") + (change.ElementId ?? string.Empty) : itemName) + " · "
                    + (string.IsNullOrWhiteSpace(commit.RevitUserName) ? commit.WindowsUserName ?? string.Empty : commit.RevitUserName));
            }
            if (previewIndex >= 20)
            {
                break;
            }
        }
        if (previewIndex == 0)
        {
            message.AppendLine();
            message.AppendLine(T("History status", "이력 상태"));
            message.AppendLine(commits.Count > 0 && (hiddenAuxiliaryRows > 0 || hiddenUnresolvedTransientRows > 0)
                ? T("No user-facing model-element changes remain after internal or unresolved transient rows are hidden.", "내부 보조 요소와 식별 불가 일시 요소 행을 제외하면 표시할 사용자 모델 요소 변경이 없습니다.")
                : T("Committed history: None", "확정 이력: 없음"));
        }

        List<string> headers = new List<string>
        {
            T("Time", "시각"),
            T("User", "사용자"),
            T("Change", "변경"),
            T("Element ID", "요소 ID"),
            T("Category", "카테고리"),
            T("Name", "이름"),
            T("Family", "패밀리"),
            T("Type", "타입"),
            T("Transaction", "트랜잭션"),
            T("PC", "PC"),
            T("Commit", "확정 종류"),
            T("Storage status", "보관 상태"),
            T("Local Save", "로컬 저장"),
            T("First observed", "최초 관찰"),
            T("Last observed", "최종 관찰"),
            T("Windows user", "Windows 사용자"),
            T("Summary", "요약"),
            T("Attribution", "귀속 신뢰도"),
            T("Policy validation", "정책 확인"),
            T("Integrity", "무결성")
        };
        List<List<string>> rows = new List<List<string>>();
        foreach (FamilyBrowserElementChangeCommit commit in commits)
        {
            if (commit.CoverageGapOnly && (commit.Changes == null || commit.Changes.Count == 0) && rows.Count < rowLimit)
            {
                rows.Add(new List<string>
                {
                    FormatProjectCatalogDate(commit.CommittedAtUtc),
                    string.IsNullOrWhiteSpace(commit.RevitUserName) ? commit.WindowsUserName ?? string.Empty : commit.RevitUserName,
                    T("Coverage gap", "관찰 공백"),
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    string.Join(" / ", commit.TransactionNames ?? new List<string>()),
                    commit.MachineName ?? string.Empty,
                    ProjectElementCommitKindLabel(commit.CommitKind),
                    ProjectElementStorageStatusLabel(commit, immutableEntryIds, uploadPendingEntryIds, checkpointEntryIds, checkpointHistory.SynchronizationSucceededEntryIds),
                    FormatProjectCatalogDate(commit.LocalSaveProtectedAtUtc),
                    FormatProjectCatalogDate(commit.TrackingStartedAtUtc),
                    FormatProjectCatalogDate(commit.CommittedAtUtc),
                    commit.WindowsUserName ?? string.Empty,
                    commit.CoverageNote ?? string.Empty,
                    ProjectElementAttributionLabel(commit, null),
                    string.Equals(commit.PolicyValidationState, "DisablePendingCommit", StringComparison.OrdinalIgnoreCase)
                        ? T("Disabled remotely / protected through commit", "원격 OFF / 저장 경계까지 보호")
                        : (string.Equals(commit.PolicyValidationState, "LastKnownEnabled", StringComparison.OrdinalIgnoreCase)
                            ? T("Last confirmed enabled", "마지막 확인 ON")
                            : T("Live enabled", "실시간 ON 확인")),
                    string.IsNullOrWhiteSpace(commit.IntegritySha256) ? T("Legacy / unverified", "이전 형식 / 미검증") : T("Checksum verified", "체크섬 확인")
                });
            }
            foreach (FamilyBrowserElementChangeItem change in visibleChangesByCommit[commit])
            {
                if (change == null || rows.Count >= rowLimit)
                {
                    continue;
                }
                rows.Add(new List<string>
                {
                    FormatProjectCatalogDate(commit.CommittedAtUtc),
                    string.IsNullOrWhiteSpace(commit.RevitUserName) ? commit.WindowsUserName ?? string.Empty : commit.RevitUserName,
                    ProjectElementChangeKindLabel(change.ChangeKind),
                    change.ElementId ?? string.Empty,
                    ProjectElementCategoryLabel(change),
                    ProjectElementName(change),
                    change.FamilyName ?? string.Empty,
                    change.TypeName ?? string.Empty,
                    string.Join(" / ", change.TransactionNames ?? new List<string>()),
                    commit.MachineName ?? string.Empty,
                    ProjectElementCommitKindLabel(commit.CommitKind),
                    ProjectElementStorageStatusLabel(commit, immutableEntryIds, uploadPendingEntryIds, checkpointEntryIds, checkpointHistory.SynchronizationSucceededEntryIds),
                    FormatProjectCatalogDate(commit.LocalSaveProtectedAtUtc),
                    FormatProjectCatalogDate(change.FirstObservedAtUtc),
                    FormatProjectCatalogDate(change.LastObservedAtUtc),
                    commit.WindowsUserName ?? string.Empty,
                    ProjectElementChangeSummaryLabel(change),
                    ProjectElementAttributionLabel(commit, change),
                    string.Equals(commit.PolicyValidationState, "DisablePendingCommit", StringComparison.OrdinalIgnoreCase)
                        ? T("Disabled remotely / protected through commit", "원격 OFF / 저장 경계까지 보호")
                        : (string.Equals(commit.PolicyValidationState, "LastKnownEnabled", StringComparison.OrdinalIgnoreCase)
                            ? T("Last confirmed enabled", "마지막 확인 ON")
                            : T("Live enabled", "실시간 ON 확인")),
                    string.IsNullOrWhiteSpace(commit.IntegritySha256) ? T("Legacy / unverified", "이전 형식 / 미검증") : T("Checksum verified", "체크섬 확인")
                });
            }
            if (rows.Count >= rowLimit)
            {
                break;
            }
        }
        if (rows.Count == 0)
        {
            bool onlyNonUserFacingRowsWereHidden = commits.Count > 0 &&
                (hiddenAuxiliaryRows > 0 || hiddenUnresolvedTransientRows > 0);
            rows.Add(new List<string>
            {
                string.Empty,
                string.Empty,
                onlyNonUserFacingRowsWereHidden ? T("No user-facing change", "사용자 모델 변경 없음") : T("No history", "이력 없음"),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                onlyNonUserFacingRowsWereHidden ? T("Internal or unresolved transient rows hidden", "내부 보조 또는 식별 불가 일시 요소 제외") : T("No committed history", "확정 이력 없음"),
                onlyNonUserFacingRowsWereHidden ? T("Stored record retained", "원본 기록 유지") : T("No stored record", "보관 기록 없음"),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                onlyNonUserFacingRowsWereHidden
                    ? T("The immutable source records remain intact; internal support rows and unresolved same-boundary transient rows are omitted from this view and Excel export.", "무결성 원본 기록은 그대로 유지되며 내부 보조 요소와 동일 작업 경계에서 식별할 수 없는 일시 요소 행은 이 화면과 Excel 내보내기에서 제외됩니다.")
                    : T("Tracking begins after this RVT is registered in Permissions / Guard with Element Change Tracking enabled.", "이 RVT를 권한 / 차단에 등록하고 요소 변경 추적을 체크한 이후부터 기록됩니다."),
                string.Empty,
                string.Empty,
                string.Empty
            });
        }

        string historyCaption = T("Project Element Change History", "프로젝트 요소 변경 이력") + (string.IsNullOrWhiteSpace(projectTitle) ? string.Empty : " - " + projectTitle);
        int confirmedRecordCount = commits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return immutableEntryIds.Contains(commit.EntryId ?? string.Empty); });
        int pendingRecordCount = Math.Max(0, commits.Count - confirmedRecordCount);
        using (TrackedProjectElementHistoryHtmlForm dialog = new TrackedProjectElementHistoryHtmlForm(
            IsKorean(),
            historyCaption,
            message.ToString().Trim(),
            headers,
            rows,
            created,
            modified,
            deleted,
            confirmedRecordCount,
            pendingRecordCount))
        {
            dialog.ExportRequested += delegate
            {
                FamilyBrowserResultExcelExportUi.SaveRows(
                    dialog,
                    null,
                    IsKorean(),
                    FamilyBrowserResultExcelExportUi.TimestampedFileName("KKY-FamilyBrowser-Element-Change-History"),
                    "ElementChanges",
                    headers,
                    rows);
            };
            dialog.ShowDialog(this);
        }
    }

    private bool RecordProjectElementChangeTrackingPolicyChange(
        bool previousEnabled,
        bool enabled,
        string correlationId,
        string phase,
        string outcome,
        string failureDetail,
        string changeSource = "")
    {
        Document doc = GetActiveDocument();
        string now = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        string normalizedCorrelationId = string.IsNullOrWhiteSpace(correlationId) ? Guid.NewGuid().ToString("N") : correlationId.Trim();
        string normalizedPhase = string.IsNullOrWhiteSpace(phase) ? "Completed" : phase.Trim();
        bool prepared = string.Equals(normalizedPhase, "Prepared", StringComparison.OrdinalIgnoreCase);
        FamilyBrowserOperationLogEntry entry = new FamilyBrowserOperationLogEntry
        {
            EntryId = normalizedCorrelationId + "-" + normalizedPhase.ToLowerInvariant(),
            RecordedAtUtc = now,
            UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity(),
            OperationKind = "ProjectElementChangeTrackingPolicy",
            DocumentTitle = doc == null ? string.Empty : doc.Title ?? string.Empty,
            DocumentPath = doc == null ? string.Empty : ProjectSnapshotStore.ResolveProjectIdentityPath(doc),
            CandidateKind = "Policy",
            PlannedAction = enabled ? "Enable" : "Disable",
            Outcome = string.IsNullOrWhiteSpace(outcome) ? normalizedPhase : outcome.Trim(),
            Details = "correlation=" + normalizedCorrelationId
                + ";phase=" + normalizedPhase
                + ";previous=" + (previousEnabled ? "on" : "off")
                + ";new=" + (enabled ? "on" : "off")
                + ";windowsUser=" + Environment.UserName
                + ";machine=" + Environment.MachineName
                + (string.IsNullOrWhiteSpace(changeSource) ? string.Empty : ";source=" + changeSource.Trim())
                + (string.IsNullOrWhiteSpace(failureDetail) ? string.Empty : ";failure=" + failureDetail.Trim()),
            CommitState = prepared ? "Prepared" : (string.Equals(normalizedPhase, "Failed", StringComparison.OrdinalIgnoreCase) ? "Failed" : "Committed"),
            CommitKind = "PolicyChange" + normalizedPhase,
            CommittedAtUtc = prepared ? string.Empty : now
        };
        return FamilyBrowserTrackingPersistenceService.PersistOperationEntries(_workspaceRoot, new[] { entry });
    }

    private static bool IsUserFacingProjectElementChange(FamilyBrowserElementChangeItem change)
    {
        return FamilyBrowserElementHistoryProjectionPolicy.IsUserFacingChange(change);
    }

    private string ProjectElementCategoryLabel(FamilyBrowserElementChangeItem change)
    {
        string trackingKind = ProjectElementTrackingKind(change);
        if (string.Equals(trackingKind, "SharedParameter", StringComparison.OrdinalIgnoreCase)) return T("Shared Parameter", "공유 파라미터");
        if (string.Equals(trackingKind, "ProjectParameter", StringComparison.OrdinalIgnoreCase)) return T("Project Parameter", "프로젝트 파라미터");
        if (string.Equals(trackingKind, "Grid", StringComparison.OrdinalIgnoreCase)) return T("Grid", "그리드");
        return change == null ? string.Empty : change.CategoryName ?? string.Empty;
    }

    private static string ProjectElementName(FamilyBrowserElementChangeItem change)
    {
        if (change == null)
        {
            return string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(change.ElementName))
        {
            return change.ElementName;
        }
        FamilyBrowserTrackedElementState state = change.After ?? change.Before;
        return state == null ? string.Empty : state.ElementName ?? string.Empty;
    }

    private static string ProjectElementTrackingKind(FamilyBrowserElementChangeItem change)
    {
        if (change == null)
        {
            return string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(change.TrackingKind))
        {
            return change.TrackingKind;
        }
        FamilyBrowserTrackedElementState state = change.After ?? change.Before;
        return state == null ? string.Empty : state.TrackingKind ?? string.Empty;
    }

    private string ProjectElementChangeSummaryLabel(FamilyBrowserElementChangeItem change)
    {
        if (change == null)
        {
            return string.Empty;
        }
        FamilyBrowserTrackedElementState before = change.Before;
        FamilyBrowserTrackedElementState after = change.After;
        string trackingKind = ProjectElementTrackingKind(change);
        bool parameter = string.Equals(trackingKind, "SharedParameter", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(trackingKind, "ProjectParameter", StringComparison.OrdinalIgnoreCase);
        if (parameter)
        {
            string subject = string.Equals(trackingKind, "SharedParameter", StringComparison.OrdinalIgnoreCase)
                ? T("Shared parameter", "공유 파라미터")
                : T("Project parameter", "프로젝트 파라미터");
            if (string.Equals(change.ChangeKind, "Created", StringComparison.OrdinalIgnoreCase))
            {
                return subject + T(" added: ", " 추가: ") + DescribeProjectParameterState(after);
            }
            if (string.Equals(change.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase))
            {
                return subject + T(" removed: ", " 삭제: ") + DescribeProjectParameterState(before);
            }
            if (string.Equals(change.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase))
            {
                return subject + T(" was added and removed before the Save or Synchronize boundary.", "가 저장 또는 동기화 전에 추가된 뒤 삭제되었습니다.");
            }
            List<string> parameterChanges = new List<string>();
            AppendProjectHistoryDifference(parameterChanges, T("Name", "이름"), before == null ? string.Empty : before.ElementName, after == null ? string.Empty : after.ElementName);
            AppendProjectHistoryDifference(parameterChanges, "GUID", before == null ? string.Empty : before.SharedParameterGuid, after == null ? string.Empty : after.SharedParameterGuid);
            AppendProjectHistoryDifference(parameterChanges, T("Binding", "바인딩"), ProjectParameterBindingLabel(before == null ? string.Empty : before.ParameterBindingKind), ProjectParameterBindingLabel(after == null ? string.Empty : after.ParameterBindingKind));
            AppendProjectHistoryDifference(parameterChanges, T("Categories", "적용 카테고리"), before == null ? string.Empty : before.ParameterBoundCategories, after == null ? string.Empty : after.ParameterBoundCategories);
            AppendProjectHistoryDifference(parameterChanges, T("Parameter group", "파라미터 그룹"), before == null ? string.Empty : before.ParameterGroup, after == null ? string.Empty : after.ParameterGroup);
            AppendProjectHistoryDifference(parameterChanges, T("Data type", "데이터 형식"), before == null ? string.Empty : before.ParameterDataType, after == null ? string.Empty : after.ParameterDataType);
            AppendProjectHistoryDifference(parameterChanges, T("Varies across groups", "그룹별 값 변경"), before == null ? string.Empty : before.ParameterVariesAcrossGroups, after == null ? string.Empty : after.ParameterVariesAcrossGroups);
            return parameterChanges.Count == 0
                ? subject + T(" definition or binding changed.", " 정의 또는 바인딩이 변경되었습니다.")
                : subject + T(" changed: ", " 변경: ") + string.Join("; ", parameterChanges);
        }
        if (string.Equals(trackingKind, "Grid", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(change.ChangeKind, "Created", StringComparison.OrdinalIgnoreCase)) return T("Grid added: ", "그리드 추가: ") + ProjectElementName(change);
            if (string.Equals(change.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase)) return T("Grid removed: ", "그리드 삭제: ") + ProjectElementName(change);
            if (string.Equals(change.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase)) return T("Grid was added and removed before the Save or Synchronize boundary.", "그리드가 저장 또는 동기화 전에 추가된 뒤 삭제되었습니다.");
            List<string> gridChanges = new List<string>();
            AppendProjectHistoryDifference(gridChanges, T("Name", "이름"), before == null ? string.Empty : before.ElementName, after == null ? string.Empty : after.ElementName);
            AppendProjectHistoryDifference(gridChanges, T("Type", "타입"), before == null ? string.Empty : before.TypeName, after == null ? string.Empty : after.TypeName);
            AppendProjectHistoryDifference(gridChanges, T("Curve", "기준선"), before == null ? string.Empty : before.GridCurveSignature, after == null ? string.Empty : after.GridCurveSignature);
            AppendProjectHistoryDifference(gridChanges, T("Extents", "범위"), before == null ? string.Empty : before.GridExtentsSignature, after == null ? string.Empty : after.GridExtentsSignature);
            AppendProjectHistoryDifference(gridChanges, T("Pinned", "고정"), before == null ? string.Empty : before.GridPinnedState, after == null ? string.Empty : after.GridPinnedState);
            AppendProjectHistoryDifference(gridChanges, T("Workset", "작업세트"), before == null ? string.Empty : before.WorksetId, after == null ? string.Empty : after.WorksetId);
            return gridChanges.Count == 0
                ? T("Grid geometry, parameter, or datum state changed.", "그리드의 형상, 파라미터 또는 기준 상태가 변경되었습니다.")
                : T("Grid changed: ", "그리드 변경: ") + string.Join("; ", gridChanges);
        }
        return change.ChangeSummary ?? string.Empty;
    }

    private string DescribeProjectParameterState(FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            return "-";
        }
        List<string> parts = new List<string> { string.IsNullOrWhiteSpace(state.ElementName) ? "-" : state.ElementName };
        if (!string.IsNullOrWhiteSpace(state.SharedParameterGuid)) parts.Add("GUID " + state.SharedParameterGuid);
        if (!string.IsNullOrWhiteSpace(state.ParameterBindingKind)) parts.Add(T("Binding: ", "바인딩: ") + ProjectParameterBindingLabel(state.ParameterBindingKind));
        if (!string.IsNullOrWhiteSpace(state.ParameterBoundCategories)) parts.Add(T("Categories: ", "적용 카테고리: ") + state.ParameterBoundCategories);
        if (!string.IsNullOrWhiteSpace(state.ParameterGroup)) parts.Add(T("Group: ", "그룹: ") + state.ParameterGroup);
        if (!string.IsNullOrWhiteSpace(state.ParameterDataType)) parts.Add(T("Data type: ", "데이터 형식: ") + state.ParameterDataType);
        return string.Join("; ", parts);
    }

    private string ProjectParameterBindingLabel(string value)
    {
        if (string.Equals(value, "Instance", StringComparison.OrdinalIgnoreCase)) return T("Instance", "인스턴스");
        if (string.Equals(value, "Type", StringComparison.OrdinalIgnoreCase)) return T("Type", "타입");
        if (string.Equals(value, "Unbound", StringComparison.OrdinalIgnoreCase)) return T("Unbound", "미바인딩");
        return value ?? string.Empty;
    }

    private static void AppendProjectHistoryDifference(ICollection<string> target, string label, string before, string after)
    {
        if (!string.Equals(before ?? string.Empty, after ?? string.Empty, StringComparison.Ordinal))
        {
            target.Add(label + ": " + (string.IsNullOrWhiteSpace(before) ? "-" : before) + " -> " + (string.IsNullOrWhiteSpace(after) ? "-" : after));
        }
    }

    private string ProjectElementChangeKindLabel(string changeKind)
    {
        if (string.Equals(changeKind, "Created", StringComparison.OrdinalIgnoreCase)) return T("Created", "생성");
        if (string.Equals(changeKind, "Modified", StringComparison.OrdinalIgnoreCase)) return T("Modified", "수정");
        if (string.Equals(changeKind, "Deleted", StringComparison.OrdinalIgnoreCase)) return T("Deleted", "삭제");
        if (string.Equals(changeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase)) return T("Created then deleted", "생성 후 삭제");
        return changeKind ?? string.Empty;
    }

    private static bool ProjectElementHasExternalRebaseGap(FamilyBrowserElementChangeCommit commit)
    {
        return commit != null &&
            (string.Equals(commit.AttributionConfidence, "ClientObservedWithExternalRebaseGap", StringComparison.OrdinalIgnoreCase) ||
             (commit.CoverageNote ?? string.Empty).IndexOf("incoming central/reload update could not be fully rebased", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private string ProjectElementAttributionLabel(FamilyBrowserElementChangeCommit commit, FamilyBrowserElementChangeItem change)
    {
        if (commit == null)
        {
            return T("Review required", "검토 필요");
        }
        List<string> labels = new List<string>();
        if (change != null && change.ExternalUpdateOverlap)
        {
            labels.Add(T("Incoming update overlap", "외부 업데이트 겹침"));
        }
        if (ProjectElementHasExternalRebaseGap(commit))
        {
            labels.Add(T("External update coverage gap", "외부 업데이트 관찰 공백"));
        }
        if (commit.CommitBoundaryReadFailureCount > 0 ||
            string.Equals(commit.AttributionConfidence, "ClientObservedWithCommitBoundaryGap", StringComparison.OrdinalIgnoreCase))
        {
            labels.Add(T("Save boundary coverage gap", "저장 경계 관찰 공백"));
        }
        if (commit.EventReadFailureCount > 0 ||
            (commit.CoverageNote ?? string.Empty).IndexOf("DocumentChanged ID or operation-metadata", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            labels.Add(T("DocumentChanged coverage gap", "DocumentChanged 관찰 공백"));
        }
        if (commit.UnmatchedUndoCount > 0 || commit.UnmatchedRedoCount > 0)
        {
            labels.Add(T("Undo/Redo sequence uncertain", "Undo/Redo 순서 불확실"));
        }
        if (string.IsNullOrWhiteSpace(commit.RevitUserName))
        {
            labels.Add(T("Revit user unavailable", "Revit 사용자 확인 불가"));
        }
        if (change == null && commit.CoverageGapOnly)
        {
            labels.Add(T("Element ID unavailable", "요소 ID 확인 불가"));
        }
        if (labels.Count == 0)
        {
            return T("Client observed", "클라이언트 관찰");
        }
        return string.Join(" / ", labels.Distinct(StringComparer.OrdinalIgnoreCase)) + " / " + T("review required", "검토 필요");
    }

    private string ProjectElementCommitKindLabel(string commitKind)
    {
        if (string.Equals(commitKind, "SynchronizeWithCentral", StringComparison.OrdinalIgnoreCase)) return T("Synchronize with Central", "센트럴 동기화");
        if (string.Equals(commitKind, "SaveAs", StringComparison.OrdinalIgnoreCase)) return T("Save As", "다른 이름으로 저장");
        if (string.Equals(commitKind, "Save", StringComparison.OrdinalIgnoreCase)) return T("Save", "저장");
        if (string.Equals(commitKind, "WorksharedLocalSavePendingSync", StringComparison.OrdinalIgnoreCase)) return T("Local Save", "로컬 저장");
        return commitKind ?? string.Empty;
    }

    private string ProjectElementStorageStatusLabel(
        FamilyBrowserElementChangeCommit commit,
        ISet<string> immutableEntryIds,
        ISet<string> uploadPendingEntryIds,
        ISet<string> checkpointEntryIds,
        ISet<string> synchronizationSucceededEntryIds)
    {
        string entryId = commit == null ? string.Empty : commit.EntryId ?? string.Empty;
        if (immutableEntryIds != null && immutableEntryIds.Contains(entryId))
        {
            return T("Confirmed in management history", "관리 이력 확정");
        }
        if (synchronizationSucceededEntryIds != null && synchronizationSucceededEntryIds.Contains(entryId))
        {
            return T("Synchronized / history promotion pending", "동기화 완료 / 이력 승격 대기");
        }
        if (checkpointEntryIds != null && checkpointEntryIds.Contains(entryId))
        {
            return T("Saved locally / synchronization pending", "로컬 저장 / 동기화 대기");
        }
        if (uploadPendingEntryIds != null && uploadPendingEntryIds.Contains(entryId))
        {
            return T("Protected locally / upload pending", "로컬 보호 / 업로드 대기");
        }
        return T("Stored", "보관됨");
    }

    private void AppendProjectCatalogPill(StringBuilder sb)
    {
        FamilyBrowserProjectCatalogState state = _projectCatalogState;
        if (state == null)
        {
            state = FamilyBrowserProjectCatalogService.LoadLatestState(_workspaceRoot, GetActiveDocument());
            _projectCatalogState = state;
        }
        if (state != null && !string.Equals(state.StateCode, "NoProject", StringComparison.OrdinalIgnoreCase))
        {
            sb.AppendLine(Pill(T("Project Catalog: ", "프로젝트 카탈로그: ") + ProjectCatalogStatusTitle(state), ProjectCatalogTone(state)));
        }
        int pendingCount = ResolvePendingTrackingCount();
        if (pendingCount > 0)
        {
            string label = T("Tracking upload pending: ", "추적 전송 대기: ") + pendingCount.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "Saved tracking records are safe on this PC and will be uploaded when the management folder is reachable.",
                "저장된 추적 기록은 이 PC에 안전하게 보관 중이며 관리폴더 연결이 복구되면 전송됩니다.");
            sb.AppendLine("<span id=\"pendingTrackingPill\" class=\"pill warn\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        int unprotectedLocalSaveSessions = ResolveUnprotectedLocalSaveSessionCount();
        if (unprotectedLocalSaveSessions > 0)
        {
            string label = T("Local-save tracking not protected: ", "로컬 저장 추적 미보호: ") + unprotectedLocalSaveSessions.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "A successful workshared local Save could not write its restart-safe checkpoint. Keep this Revit session open and Synchronize with Central before closing.",
                "워크셰어링 로컬 저장은 성공했지만 재시작 안전 체크포인트를 기록하지 못했습니다. 현재 Revit 세션을 닫지 말고 센트럴과 동기화하세요.");
            sb.AppendLine("<span id=\"unprotectedLocalSavePill\" class=\"pill bad\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        int unprotectedCommitBoundarySessions = ResolveUnprotectedCommitBoundarySessionCount();
        if (unprotectedCommitBoundarySessions > 0)
        {
            string label = T("Save/Sync tracking protection failed: ", "저장/동기화 추적 보호 실패: ") + unprotectedCommitBoundarySessions.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "A Save or Synchronize boundary could not protect observed element activity. Keep Revit open and repeat a successful Save or Synchronize until this warning disappears.",
                "저장 또는 동기화 경계에서 관찰된 요소 작업을 보호하지 못했습니다. Revit을 닫지 말고 이 경고가 사라질 때까지 저장 또는 동기화를 다시 완료하세요.");
            sb.AppendLine("<span id=\"unprotectedCommitBoundaryPill\" class=\"pill bad\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        FamilyBrowserElementSessionCheckpointCountResult localSyncPendingStatus = ResolvePendingElementSessionCheckpointStatus();
        int localSyncPending = localSyncPendingStatus.Count;
        int synchronizedHistoryPromotionPending = Math.Min(localSyncPending, Math.Max(0, localSyncPendingStatus.SynchronizationSucceededCount));
        int synchronizationPending = Math.Max(0, localSyncPending - synchronizedHistoryPromotionPending);
        if (localSyncPendingStatus.LockUnavailable)
        {
            string label = T("Local checkpoint busy / review required", "로컬 체크포인트 잠금 중 / 확인 필요");
            string detail = T(
                "Another Revit process may be updating the protected local checkpoint. The Browser does not treat this as zero pending work; close the other process or refresh after the operation finishes.",
                "다른 Revit 프로세스가 보호된 로컬 체크포인트를 갱신 중일 수 있습니다. Browser는 이를 대기 기록 0개로 취급하지 않습니다. 다른 프로세스를 닫거나 작업이 끝난 뒤 새로고침하세요.");
            sb.AppendLine("<span id=\"pendingElementSessionLockPill\" class=\"pill bad\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        else if (synchronizationPending > 0)
        {
            string label = T("Local saves awaiting sync: ", "동기화 대기 로컬 저장: ") + synchronizationPending.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "Workshared element activity was protected after a local Save. It is added to managed immutable history only after a successful Synchronize with Central.",
                "워크셰어링 요소 작업을 로컬 저장 후 보호 중입니다. 센트럴과 동기화가 성공해야 관리 이력으로 확정됩니다.");
            sb.AppendLine("<span id=\"pendingElementSessionPill\" class=\"pill warn\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        if (!localSyncPendingStatus.LockUnavailable && synchronizedHistoryPromotionPending > 0)
        {
            string label = T("Synchronized history promotion pending: ", "동기화 완료 / 이력 승격 대기: ") + synchronizedHistoryPromotionPending.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "Synchronize with Central already succeeded. The finalized checkpoint remains restart-safe and will retry immutable history promotion; another synchronization is not required merely for this warning.",
                "센트럴 동기화는 이미 성공했습니다. 확정 체크포인트는 재시작 후에도 보호되며 관리 이력 승격을 다시 시도합니다. 이 경고만을 위해 다시 동기화할 필요는 없습니다.");
            sb.AppendLine("<span id=\"finalizedElementSessionPill\" class=\"pill warn\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        int invalidLocalSyncPending = localSyncPendingStatus.LockUnavailable ? 0 : ResolveInvalidElementSessionCheckpointCount();
        if (invalidLocalSyncPending > 0)
        {
            string label = T("Local tracking recovery required: ", "로컬 추적 복구 필요: ") + invalidLocalSyncPending.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "One or more workshared local-save checkpoints failed integrity validation. Open Debug Log before deleting or replacing local tracking data.",
                "워크셰어링 로컬 저장 체크포인트의 무결성 확인에 실패했습니다. 로컬 추적 데이터를 삭제하거나 교체하기 전에 디버그 로그를 확인하세요.");
            sb.AppendLine("<span id=\"invalidElementSessionPill\" class=\"pill bad\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        int mismatchedLocalSyncPending = localSyncPendingStatus.LockUnavailable ? 0 : ResolveMismatchedElementSessionCheckpointCount();
        if (mismatchedLocalSyncPending > 0)
        {
            string label = T("Local tracking folder mismatch: ", "로컬 추적 폴더 불일치: ") + mismatchedLocalSyncPending.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "Protected checkpoints are still bound to a previous management folder. Use the managed-folder migration workflow; do not delete the local tracking folder.",
                "보호된 체크포인트가 이전 관리폴더에 묶여 있습니다. 관리폴더 이관 기능을 사용하고 로컬 추적 폴더를 삭제하지 마세요.");
            sb.AppendLine("<span id=\"mismatchedElementSessionPill\" class=\"pill warn\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
        int deferredPolicyDisableSessions = ResolveDeferredPolicyDisableSessionCount();
        if (deferredPolicyDisableSessions > 0)
        {
            string label = T("Tracking disable pending Save/Sync: ", "추적 해제 Save/Sync 대기: ") + deferredPolicyDisableSessions.ToString(CultureInfo.InvariantCulture);
            string detail = T(
                "Another administrator disabled tracking after this client had already observed uncommitted activity. Existing evidence remains protected only until the next successful Save or Synchronize boundary; no new session starts afterward.",
                "다른 관리자가 이 PC에 미확정 관찰 작업이 있는 상태에서 추적을 해제했습니다. 기존 증거는 다음 저장 또는 동기화 성공 경계까지만 보호되며 이후 새 추적 세션은 시작하지 않습니다.");
            sb.AppendLine("<span id=\"deferredTrackingDisablePill\" class=\"pill warn\" title=\"" + Attr(detail) + "\">" + Html(label) + "</span>");
        }
    }

    private void AppendHomeProjectCatalogBoard(StringBuilder sb)
    {
        AppendHomePendingTrackingQueueBoard(sb);
        FamilyBrowserProjectCatalogState state = _projectCatalogState;
        if (state == null)
        {
            state = FamilyBrowserProjectCatalogService.LoadLatestState(_workspaceRoot, GetActiveDocument());
            _projectCatalogState = state;
        }
        if (state == null || string.Equals(state.StateCode, "NoProject", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        string tone = ProjectCatalogTone(state);
        sb.AppendLine("<div class=\"home-board project-catalog-board " + Attr(tone) + "\">");
        sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Project Name Catalog Tracking", "프로젝트 이름 카탈로그 추적")) + "</strong><span>" + Html(T("A lightweight scan compares family and system type names with the last accepted project baseline.", "패밀리명과 시스템 타입명만 빠르게 훑어 마지막 승인 프로젝트 기준선과 비교합니다.")) + "</span></div>");
        sb.AppendLine("<div class=\"project-catalog-summary\">");
        AppendProjectCatalogMetric(sb, T("Status", "상태"), ProjectCatalogStatusTitle(state), tone);
        AppendProjectCatalogMetric(sb, T("Families", "패밀리"), state.FamilyCount.ToString(CultureInfo.InvariantCulture), "info");
        AppendProjectCatalogMetric(sb, T("Family Types", "패밀리 타입"), state.FamilyTypeCount.ToString(CultureInfo.InvariantCulture), "info");
        AppendProjectCatalogMetric(sb, T("System Types", "시스템 타입"), state.SystemTypeCount.ToString(CultureInfo.InvariantCulture), "info");
        AppendProjectCatalogMetric(sb, T("Added / Removed", "추가 / 삭제"), state.AddedCount.ToString(CultureInfo.InvariantCulture) + " / " + state.RemovedCount.ToString(CultureInfo.InvariantCulture), state.Changed ? "bad" : "good");
        AppendProjectCatalogMetric(sb, T("External / Untracked", "외부 / 미추적"), state.ExternalUntrackedChangeCount.ToString(CultureInfo.InvariantCulture), state.ExternalUntrackedChangeCount > 0 ? "bad" : "good");
        sb.AppendLine("</div>");
        sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(ProjectCatalogReason(state)) + "</div>");
        if (state.Changed && state.Changes != null && state.Changes.Count > 0)
        {
            sb.AppendLine("<div class=\"table-scroll project-catalog-table-wrap\"><table class=\"project-catalog-table\"><tr><th>" + Html(T("Change", "변경")) + "</th><th>" + Html(T("Kind", "종류")) + "</th><th>" + Html(T("Item", "항목")) + "</th><th>" + Html(T("Source", "출처")) + "</th></tr>");
            foreach (FamilyBrowserProjectCatalogChange change in state.Changes.Take(8))
            {
                string item = ProjectCatalogItemLabel(change);
                string source = string.Equals(change.Attribution, "KnownBrowser", StringComparison.OrdinalIgnoreCase) ? T("Family Browser tracked", "Family Browser 추적") : T("External / untracked", "외부 / 미추적");
                sb.AppendLine("<tr><td>" + Html(ProjectCatalogChangeKindLabel(change)) + "</td><td>" + Html(ProjectCatalogEntryKindLabel(change.EntryKind)) + "</td><td title=\"" + Attr(item) + "\">" + Html(SafeShortName(item, 90)) + "</td><td>" + Html(source) + "</td></tr>");
            }
            sb.AppendLine("</table></div>");
        }
        sb.AppendLine("<div class=\"project-catalog-actions\"><a class=\"tool primary\" href=\"kkyfb:project-catalog-check\">" + Html(T("Check Now", "지금 확인")) + "</a>");
        if (_adminModeEnabled)
        {
            sb.AppendLine("<a class=\"tool\" href=\"kkyfb:project-catalog-accept\">" + Html(T("Accept Current Names as Baseline", "현재 이름 목록을 기준선으로 승인")) + "</a>");
        }
        sb.AppendLine("<span>" + Html(T("Last checked: ", "마지막 확인: ") + FormatProjectCatalogDate(state.CheckedAtUtc)) + "</span></div>");
        sb.AppendLine("</div>");
    }

    private int ResolvePendingTrackingCount()
    {
        if (_auditTrackingPendingCount >= 0)
        {
            return _auditTrackingPendingCount;
        }
        try
        {
            return FamilyBrowserTrackingPersistenceService.GetPendingCount();
        }
        catch
        {
            return 0;
        }
    }

    private FamilyBrowserElementSessionCheckpointCountResult ResolvePendingElementSessionCheckpointStatus()
    {
        try
        {
            return FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus(_workspaceRoot);
        }
        catch
        {
            return new FamilyBrowserElementSessionCheckpointCountResult { LockUnavailable = true };
        }
    }

    private int ResolveInvalidElementSessionCheckpointCount()
    {
        try
        {
            return FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount();
        }
        catch
        {
            return 0;
        }
    }

    private int ResolveMismatchedElementSessionCheckpointCount()
    {
        try
        {
            return FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount(_workspaceRoot);
        }
        catch
        {
            return 0;
        }
    }

    private static int ResolveDeferredPolicyDisableSessionCount()
    {
        try
        {
            return FamilyBrowserElementChangeTrackingService.GetDeferredPolicyDisableSessionCount();
        }
        catch
        {
            return 0;
        }
    }

    private static int ResolveUnprotectedLocalSaveSessionCount()
    {
        try
        {
            return FamilyBrowserElementChangeTrackingService.GetUnprotectedLocalSaveSessionCount();
        }
        catch
        {
            return 0;
        }
    }

    private static int ResolveUnprotectedCommitBoundarySessionCount()
    {
        try
        {
            return FamilyBrowserElementChangeTrackingService.GetUnprotectedCommitBoundarySessionCount();
        }
        catch
        {
            return 0;
        }
    }

    private void AppendHomePendingTrackingQueueBoard(StringBuilder sb)
    {
        FamilyBrowserElementSessionCheckpointCountResult localSyncPendingStatus = ResolvePendingElementSessionCheckpointStatus();
        int localSyncPending = localSyncPendingStatus.Count;
        int synchronizedHistoryPromotionPending = Math.Min(localSyncPending, Math.Max(0, localSyncPendingStatus.SynchronizationSucceededCount));
        int synchronizationPending = Math.Max(0, localSyncPending - synchronizedHistoryPromotionPending);
        int unprotectedLocalSaveSessions = ResolveUnprotectedLocalSaveSessionCount();
        if (unprotectedLocalSaveSessions > 0)
        {
            sb.AppendLine("<div id=\"unprotectedLocalSaveQueue\" class=\"home-board pending-tracking-board bad\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Local-save tracking is not restart-safe", "로컬 저장 추적이 재시작 안전 상태가 아님")) + "</strong><span>" + Html(T("The model was saved locally, but the pending element history could not be written to its protected checkpoint.", "모델은 로컬 저장되었지만 대기 중인 요소 이력을 보호 체크포인트에 기록하지 못했습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Affected open sessions", "영향받은 열린 세션"), unprotectedLocalSaveSessions.ToString(CultureInfo.InvariantCulture), "bad");
            AppendProjectCatalogMetric(sb, T("Safe to close Revit", "Revit 종료 안전 여부"), T("No", "아니오"), "bad");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Do not close this Revit session. Restore access to the local tracking folder if needed, then Save again or complete Synchronize with Central until this warning disappears.", "현재 Revit 세션을 닫지 마세요. 필요하면 로컬 추적 폴더 접근을 복구한 뒤 이 경고가 사라질 때까지 다시 저장하거나 센트럴과 동기화를 완료하세요.")) + "</div>");
            sb.AppendLine("</div>");
        }
        int unprotectedCommitBoundarySessions = ResolveUnprotectedCommitBoundarySessionCount();
        if (unprotectedCommitBoundarySessions > 0)
        {
            sb.AppendLine("<div id=\"unprotectedCommitBoundaryQueue\" class=\"home-board pending-tracking-board bad\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Save or synchronization did not protect tracking evidence", "저장 또는 동기화가 추적 근거를 보호하지 못함")) + "</strong><span>" + Html(T("The Revit boundary was unreadable or local protection failed before observed activity became restart-safe.", "Revit 저장 경계를 읽지 못했거나 관찰 작업이 재시작 안전 상태가 되기 전에 로컬 보호가 실패했습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Affected open sessions", "영향받은 열린 세션"), unprotectedCommitBoundarySessions.ToString(CultureInfo.InvariantCulture), "bad");
            AppendProjectCatalogMetric(sb, T("Safe to close Revit", "Revit 종료 안전 여부"), T("No", "아니오"), "bad");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Do not close the affected Revit session. Complete another Save for a standalone model or Synchronize with Central for a workshared model, then confirm this warning is gone.", "해당 Revit 세션을 닫지 마세요. 비워크셰어링 모델은 다시 저장하고 워크셰어링 모델은 센트럴과 동기화한 뒤 이 경고가 사라졌는지 확인하세요.")) + "</div>");
            sb.AppendLine("</div>");
        }
        int deferredPolicyDisableSessions = ResolveDeferredPolicyDisableSessionCount();
        if (deferredPolicyDisableSessions > 0)
        {
            sb.AppendLine("<div id=\"deferredTrackingDisableQueue\" class=\"home-board pending-tracking-board warn\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Remote tracking disable is waiting for a commit boundary", "원격 추적 해제가 저장 경계를 기다리는 중")) + "</strong><span>" + Html(T("This client already held observed activity when another administrator disabled the shared option.", "다른 관리자가 공유 옵션을 끌 때 이 PC에 이미 관찰된 미확정 작업이 있었습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Affected open sessions", "영향받은 열린 세션"), deferredPolicyDisableSessions.ToString(CultureInfo.InvariantCulture), "warn");
            AppendProjectCatalogMetric(sb, T("New sessions after boundary", "저장 경계 이후 새 세션"), T("Disabled", "사용 안 함"), "good");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Save a standalone project or Synchronize a workshared project to protect the already-observed evidence and finish disabling tracking. Work performed before and after the remote toggle can share this final commit when Revit has not yet reached a save boundary.", "비워크셰어링 프로젝트는 저장하고 워크셰어링 프로젝트는 센트럴과 동기화하면 이미 관찰된 증거를 보호한 뒤 추적 해제가 완료됩니다. Revit이 아직 저장 경계에 도달하지 않았다면 원격 해제 전후 작업이 마지막 확정 기록 한 묶음에 포함될 수 있습니다.")) + "</div>");
            sb.AppendLine("</div>");
        }
        int invalidLocalSyncPending = localSyncPendingStatus.LockUnavailable ? 0 : ResolveInvalidElementSessionCheckpointCount();
        if (invalidLocalSyncPending > 0)
        {
            sb.AppendLine("<div id=\"invalidElementSessionQueue\" class=\"home-board pending-tracking-board bad\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Local tracking checkpoint requires recovery", "로컬 추적 체크포인트 복구 필요")) + "</strong><span>" + Html(T("A saved workshared checkpoint could not pass identity or checksum validation.", "저장된 워크셰어링 체크포인트가 파일 정체성 또는 체크섬 확인을 통과하지 못했습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Affected local checkpoints", "영향받은 로컬 체크포인트"), invalidLocalSyncPending.ToString(CultureInfo.InvariantCulture), "bad");
            AppendProjectCatalogMetric(sb, T("Automatic publish", "자동 확정"), T("Blocked", "차단됨"), "bad");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Do not delete the local tracking folder. Open Debug Log and preserve the affected local RVT until an administrator reviews the checkpoint.", "로컬 추적 폴더를 삭제하지 마세요. 디버그 로그를 열고 관리자가 체크포인트를 확인할 때까지 해당 로컬 RVT를 보존하세요.")) + "</div>");
            sb.AppendLine("</div>");
        }
        int mismatchedLocalSyncPending = localSyncPendingStatus.LockUnavailable ? 0 : ResolveMismatchedElementSessionCheckpointCount();
        if (mismatchedLocalSyncPending > 0)
        {
            sb.AppendLine("<div id=\"mismatchedElementSessionQueue\" class=\"home-board pending-tracking-board warn\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Local tracking management folder mismatch", "로컬 추적 관리폴더 불일치")) + "</strong><span>" + Html(T("Valid workshared checkpoints are bound to a different management destination and cannot be published here automatically.", "유효한 워크셰어링 체크포인트가 다른 관리 대상에 묶여 있어 현재 관리폴더로 자동 확정할 수 없습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Affected local checkpoints", "영향받은 로컬 체크포인트"), mismatchedLocalSyncPending.ToString(CultureInfo.InvariantCulture), "warn");
            AppendProjectCatalogMetric(sb, T("Evidence state", "증거 상태"), T("Protected / migration required", "보호됨 / 이관 필요"), "warn");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Do not delete the local tracking folder. Return to the previous management folder or use the explicit existing-data migration action before synchronizing.", "로컬 추적 폴더를 삭제하지 마세요. 이전 관리폴더로 돌아가거나 기존 데이터 이관 기능을 명시적으로 실행한 뒤 동기화하세요.")) + "</div>");
            sb.AppendLine("</div>");
        }
        if (localSyncPendingStatus.LockUnavailable)
        {
            sb.AppendLine("<div id=\"pendingElementSessionLockQueue\" class=\"home-board pending-tracking-board bad\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Local tracking checkpoint is busy", "로컬 추적 체크포인트 잠금 중")) + "</strong><span>" + Html(T("Pending work could not be counted safely because another process holds the checkpoint lock.", "다른 프로세스가 체크포인트 잠금을 보유해 대기 기록을 안전하게 계산할 수 없습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Pending count", "대기 개수"), T("Unknown / retry required", "확인 불가 / 재시도 필요"), "bad");
            AppendProjectCatalogMetric(sb, T("Automatic recovery", "자동 복구"), T("Fail-closed", "보수적으로 차단"), "bad");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Do not delete the local tracking folder. Close the other Revit process if it is using the same local file, then refresh.", "로컬 추적 폴더를 삭제하지 마세요. 같은 로컬 파일을 사용하는 다른 Revit 프로세스가 있으면 닫은 뒤 새로고침하세요.")) + "</div>");
            sb.AppendLine("</div>");
        }
        else if (synchronizationPending > 0)
        {
            sb.AppendLine("<div id=\"pendingElementSessionQueue\" class=\"home-board pending-tracking-board warn\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Workshared local saves awaiting synchronization", "동기화 대기 중인 워크셰어링 로컬 저장")) + "</strong><span>" + Html(T("Tracked element activity survived a Revit close or restart and remains local until Central synchronization succeeds.", "추적된 요소 작업이 Revit 종료·재시작 후에도 로컬에 보호되어 있으며 센트럴 동기화 성공 전까지 확정되지 않습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Pending local projects", "대기 로컬 프로젝트"), synchronizationPending.ToString(CultureInfo.InvariantCulture), "warn");
            AppendProjectCatalogMetric(sb, T("Local checkpoint", "로컬 체크포인트"), T("Protected", "보호됨"), "good");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Open the same local file and complete Synchronize with Central. Merely pressing Refresh cannot finalize this evidence.", "같은 로컬 파일을 열어 센트럴과 동기화를 완료하세요. 새로고침만으로는 이 기록이 확정되지 않습니다.")) + "</div>");
            sb.AppendLine("</div>");
        }
        if (!localSyncPendingStatus.LockUnavailable && synchronizedHistoryPromotionPending > 0)
        {
            sb.AppendLine("<div id=\"finalizedElementSessionQueue\" class=\"home-board pending-tracking-board warn\">");
            sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Synchronization succeeded; history promotion is pending", "동기화 완료 / 관리 이력 승격 대기")) + "</strong><span>" + Html(T("The central commit succeeded and its finalized local checkpoint is restart-safe, but immutable managed history was not yet confirmed.", "센트럴 확정은 성공했고 로컬 체크포인트도 재시작 안전 상태지만, 관리 이력 승격이 아직 확인되지 않았습니다.")) + "</span></div>");
            sb.AppendLine("<div class=\"project-catalog-summary\">");
            AppendProjectCatalogMetric(sb, T("Finalized checkpoints", "동기화 완료 체크포인트"), synchronizedHistoryPromotionPending.ToString(CultureInfo.InvariantCulture), "warn");
            AppendProjectCatalogMetric(sb, T("Central synchronization", "센트럴 동기화"), T("Completed", "완료"), "good");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("Keep the local tracking folder intact. Refresh after management storage is reachable; the finalized checkpoint will replay idempotently, so another synchronization is not required solely for this state.", "로컬 추적 폴더를 그대로 보존하세요. 관리 저장소 연결 후 새로고침하면 확정 체크포인트가 중복 없이 다시 승격됩니다. 이 상태만을 이유로 다시 동기화할 필요는 없습니다.")) + "</div>");
            sb.AppendLine("</div>");
        }
        int pendingCount = ResolvePendingTrackingCount();
        if (pendingCount <= 0)
        {
            return;
        }
        string countText = pendingCount.ToString(CultureInfo.InvariantCulture);
        sb.AppendLine("<div id=\"pendingTrackingQueue\" class=\"home-board pending-tracking-board warn\">");
        sb.AppendLine("<div class=\"home-board-head\"><strong>" + Html(T("Tracking upload pending", "추적 기록 전송 대기")) + "</strong><span>" + Html(T("The management folder was unavailable when a save or synchronization was confirmed.", "저장 또는 동기화가 확정될 때 관리폴더에 연결할 수 없었습니다.")) + "</span></div>");
        sb.AppendLine("<div class=\"project-catalog-summary\">");
        AppendProjectCatalogMetric(sb, T("Pending records", "대기 기록"), countText, "warn");
        AppendProjectCatalogMetric(sb, T("Local safety copy", "로컬 안전 보관"), T("Stored", "보관됨"), "good");
        sb.AppendLine("</div>");
        sb.AppendLine("<div class=\"project-catalog-reason\">" + Html(T("No record is discarded. Restore the network connection and press Refresh; successful upload removes this warning automatically.", "기록은 삭제되지 않습니다. 네트워크 연결을 복구하고 새로고침을 누르면 전송 완료 후 이 경고가 자동으로 사라집니다.")) + "</div>");
        sb.AppendLine("<div class=\"project-catalog-actions\"><a class=\"tool primary\" href=\"kkyfb:homepage-security-refresh\">" + Html(T("Retry now", "지금 다시 시도")) + "</a></div>");
        sb.AppendLine("</div>");
    }

    private void AppendProjectCatalogMetric(StringBuilder sb, string label, string value, string tone)
    {
        sb.AppendLine("<div class=\"project-catalog-metric " + Attr(tone) + "\"><span>" + Html(label) + "</span><strong title=\"" + Attr(value) + "\">" + Html(value) + "</strong></div>");
    }

    private void InitializeProjectCatalogAuditState(FamilyBrowserDashboardAuditScenario scenario)
    {
        _projectCatalogState = new FamilyBrowserProjectCatalogState
        {
            StateCode = scenario.ProjectCatalogBaselineMissing ? "BaselineMissing" : (scenario.ProjectCatalogChanged ? "Changed" : "Current"),
            ProjectTitle = scenario.ProjectTitle,
            ProjectIdentityPath = scenario.CentralPath,
            ProjectComparableIdentity = "PATH:" + (scenario.CentralPath ?? string.Empty).ToUpperInvariant(),
            CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            AcceptedAtUtc = scenario.ProjectCatalogBaselineMissing ? string.Empty : DateTime.UtcNow.AddDays(-1.0).ToString("O", CultureInfo.InvariantCulture),
            AcceptedBy = scenario.ProjectCatalogBaselineMissing ? string.Empty : "KKY_UI_AUDIT_ADMIN",
            AcceptedCatalogHash = scenario.ProjectCatalogBaselineMissing ? string.Empty : "audit-baseline",
            CurrentCatalogHash = scenario.ProjectCatalogChanged ? "audit-current-changed" : "audit-baseline",
            Trigger = "UiAudit",
            ElapsedMilliseconds = 28,
            FamilyCount = 24,
            FamilyTypeCount = 96,
            SystemTypeCount = 18,
            AddedCount = scenario.ProjectCatalogChanged ? 1 : 0,
            RemovedCount = scenario.ProjectCatalogChanged ? 1 : 0,
            BrowserTrackedChangeCount = scenario.ProjectCatalogChanged && !scenario.ProjectCatalogUntracked ? 2 : 0,
            ExternalUntrackedChangeCount = scenario.ProjectCatalogChanged && scenario.ProjectCatalogUntracked ? 2 : 0,
            Reason = scenario.ProjectCatalogBaselineMissing ? "Audit baseline missing." : (scenario.ProjectCatalogChanged ? "Audit catalog changed." : "Audit catalog current.")
        };
        if (scenario.ProjectCatalogChanged)
        {
            _projectCatalogState.Changes.Add(new FamilyBrowserProjectCatalogChange
            {
                ChangeKind = "Added",
                EntryKind = "FamilyType",
                CategoryName = "Mechanical Equipment",
                FamilyName = "AUDIT_EXTERNAL_FAMILY",
                TypeName = "1200x600",
                Attribution = scenario.ProjectCatalogUntracked ? "ExternalUntracked" : "KnownBrowser",
                OperationUser = scenario.ProjectCatalogUntracked ? string.Empty : "KKY_UI_AUDIT_ADMIN"
            });
            _projectCatalogState.Changes.Add(new FamilyBrowserProjectCatalogChange
            {
                ChangeKind = "Removed",
                EntryKind = "SystemType",
                CategoryName = "Pipes",
                TypeName = "AUDIT_OLD_PIPE_TYPE",
                TypeClassName = "PipeType",
                Attribution = "ExternalUntracked"
            });
        }
    }

    private string ProjectCatalogStatusTitle(FamilyBrowserProjectCatalogState state)
    {
        if (state == null) return T("not checked", "미확인");
        if (!string.IsNullOrWhiteSpace(state.ErrorMessage)) return T("check failed", "확인 실패");
        if (state.Changed) return T("differences detected", "차이 감지");
        if (state.BaselineMissing) return T("baseline required", "기준선 필요");
        if (string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase)) return T("current", "일치");
        if (string.Equals(state.StateCode, "NoIdentity", StringComparison.OrdinalIgnoreCase)) return T("save required", "저장 필요");
        if (string.Equals(state.StateCode, "PublicationDeferred", StringComparison.OrdinalIgnoreCase)) return T("save / sync / reload required", "저장 / 동기화 / 최신 상태 반영 필요");
        if (string.Equals(state.StateCode, "StorageUnavailable", StringComparison.OrdinalIgnoreCase)) return T("management folder unavailable", "관리폴더 연결 안 됨");
        return T("not checked", "미확인");
    }

    private string ProjectCatalogReason(FamilyBrowserProjectCatalogState state)
    {
        if (state == null) return T("The project catalog has not been checked yet.", "프로젝트 카탈로그를 아직 확인하지 않았습니다.");
        if (!string.IsNullOrWhiteSpace(state.ErrorMessage)) return T("Project catalog check failed: ", "프로젝트 카탈로그 확인 실패: ") + state.ErrorMessage;
        if (state.Changed)
        {
            return T(
                "The accepted baseline and current project names differ. External/untracked means the exact author and time could not be proven from Family Browser operation logs.",
                "승인 기준선과 현재 프로젝트 이름 목록이 다릅니다. 외부/미추적은 Family Browser 작업 로그만으로 정확한 작업자와 시각을 확인할 수 없다는 뜻입니다.");
        }
        if (state.BaselineMissing) return T("Run Current Model Check or let an administrator accept the current name inventory.", "현재 모델 검사를 실행하거나 관리자가 현재 이름 목록을 기준선으로 승인하세요.");
        if (string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase)) return T("No family or system type name differences were found.", "패밀리명 또는 시스템 타입명 차이가 없습니다.");
        if (!string.IsNullOrWhiteSpace(state.Reason)) return state.Reason;
        return T("The project catalog is waiting for a check.", "프로젝트 카탈로그 확인을 기다리고 있습니다.");
    }

    private static string ProjectCatalogTone(FamilyBrowserProjectCatalogState state)
    {
        if (state == null) return "info";
        if (!string.IsNullOrWhiteSpace(state.ErrorMessage) || state.Changed || string.Equals(state.StateCode, "StorageUnavailable", StringComparison.OrdinalIgnoreCase)) return "bad";
        if (state.BaselineMissing || string.Equals(state.StateCode, "NoIdentity", StringComparison.OrdinalIgnoreCase) || string.Equals(state.StateCode, "PublicationDeferred", StringComparison.OrdinalIgnoreCase)) return "warn";
        return string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase) ? "good" : "info";
    }

    private string ProjectCatalogChangeLabel(FamilyBrowserProjectCatalogChange change)
    {
        if (change == null) return string.Empty;
        string source = string.Equals(change.Attribution, "KnownBrowser", StringComparison.OrdinalIgnoreCase) ? T("Browser record matched", "Browser 기록 일치") : T("external/untracked", "외부/미추적");
        return ProjectCatalogChangeKindLabel(change) + " · " + ProjectCatalogEntryKindLabel(change.EntryKind) + " · " + ProjectCatalogItemLabel(change) + " · " + source;
    }

    private string ProjectCatalogChangeKindLabel(FamilyBrowserProjectCatalogChange change)
    {
        return change != null && string.Equals(change.ChangeKind, "Removed", StringComparison.OrdinalIgnoreCase) ? T("Removed", "삭제") : T("Added", "추가");
    }

    private string ProjectCatalogEntryKindLabel(string entryKind)
    {
        if (string.Equals(entryKind, "Family", StringComparison.OrdinalIgnoreCase)) return T("Family", "패밀리");
        if (string.Equals(entryKind, "FamilyType", StringComparison.OrdinalIgnoreCase)) return T("Family Type", "패밀리 타입");
        if (string.Equals(entryKind, "SystemType", StringComparison.OrdinalIgnoreCase)) return T("System Type", "시스템 타입");
        return entryKind ?? string.Empty;
    }

    private static string ProjectCatalogItemLabel(FamilyBrowserProjectCatalogChange change)
    {
        if (change == null) return string.Empty;
        List<string> parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(change.CategoryName)) parts.Add(change.CategoryName.Trim());
        if (!string.IsNullOrWhiteSpace(change.FamilyName)) parts.Add(change.FamilyName.Trim());
        if (!string.IsNullOrWhiteSpace(change.TypeName)) parts.Add(change.TypeName.Trim());
        if (!string.IsNullOrWhiteSpace(change.TypeClassName) && string.IsNullOrWhiteSpace(change.FamilyName)) parts.Add(change.TypeClassName.Trim());
        return string.Join(" / ", parts);
    }

    private static string FormatProjectCatalogDate(string value)
    {
        DateTime parsed;
        if (!DateTime.TryParse(value ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed))
        {
            return "-";
        }
        return parsed.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
    }
}
