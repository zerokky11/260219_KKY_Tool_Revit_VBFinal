using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

public sealed class FamilyBrowserAutomaticModelCheckStatus
{
	public int SchemaVersion { get; set; }

	public string ProjectTitle { get; set; }

	public string ProjectPath { get; set; }

	public string CentralPath { get; set; }

	public string Discipline { get; set; }

	public string DisciplineLabel { get; set; }

	public string StandardRvtPath { get; set; }

	public string Trigger { get; set; }

	public string Status { get; set; }

	public string Message { get; set; }

	public string ScheduledAtUtc { get; set; }

	public string StartedAtUtc { get; set; }

	public string CompletedAtUtc { get; set; }

	public bool UsedCachedResult { get; set; }

	public string ProjectSnapshotPath { get; set; }

	public string ComparisonReportPath { get; set; }

	public int LoadableFamilyCount { get; set; }

	public int SystemTypeCount { get; set; }

	public int ProgressCurrent { get; set; }

	public int ProgressTotal { get; set; }

	public FamilyBrowserAutomaticModelCheckStatus()
	{
		SchemaVersion = 2;
		ProjectTitle = string.Empty;
		ProjectPath = string.Empty;
		CentralPath = string.Empty;
		Discipline = string.Empty;
		DisciplineLabel = string.Empty;
		StandardRvtPath = string.Empty;
		Trigger = string.Empty;
		Status = string.Empty;
		Message = string.Empty;
		ScheduledAtUtc = string.Empty;
		StartedAtUtc = string.Empty;
		CompletedAtUtc = string.Empty;
		ProjectSnapshotPath = string.Empty;
		ComparisonReportPath = string.Empty;
		ProgressTotal = 100;
	}
}

public static class FamilyBrowserAutomaticModelCheckService
{
	private sealed class PendingRequest
	{
		public Document Document { get; set; }

		public int DocumentToken { get; set; }

		public int IdleTicks { get; set; }

		public int Attempts { get; set; }

		public DateTime NextAttemptUtc { get; set; }

		public DateTime RetryDeadlineUtc { get; set; }

		public DateTime LastStatusWriteUtc { get; set; }

		public bool Running { get; set; }

		public bool Completed { get; set; }

		public DateTime CompletedUtc { get; set; }

		public FamilyBrowserAutomaticModelCheckStatus Status { get; set; }
	}

	private sealed class ExecutionOutcome
	{
		public bool Retry { get; set; }

		public FamilyBrowserAutomaticModelCheckStatus Status { get; set; }
	}

	private static readonly object SyncRoot = new object();

	private static readonly Dictionary<int, PendingRequest> Requests = new Dictionary<int, PendingRequest>();

	private const int InitialIdleDelay = 3;

	private static readonly TimeSpan RetryWindow = TimeSpan.FromMinutes(30.0);

	private static readonly TimeSpan RetryInterval = TimeSpan.FromSeconds(5.0);

	private static readonly TimeSpan StatusWriteInterval = TimeSpan.FromSeconds(15.0);

	public static void ResetAll()
	{
		lock (SyncRoot)
		{
			Requests.Clear();
		}
	}

	public static void Remove(Document document)
	{
		if (document == null)
		{
			return;
		}
		lock (SyncRoot)
		{
			Requests.Remove(RuntimeHelpers.GetHashCode(document));
		}
	}

	public static void Schedule(Document document, string trigger, bool force = false)
	{
		if (!CanInspect(document))
		{
			return;
		}
		int token = RuntimeHelpers.GetHashCode(document);
		PendingRequest scheduled = null;
		lock (SyncRoot)
		{
			PendingRequest existing;
			if (Requests.TryGetValue(token, out existing) && existing != null && !force)
			{
				if (!existing.Completed || DateTime.UtcNow - existing.CompletedUtc < TimeSpan.FromMinutes(5.0))
				{
					return;
				}
			}
			DateTime now = DateTime.UtcNow;
			scheduled = new PendingRequest
			{
				Document = document,
				DocumentToken = token,
				NextAttemptUtc = now,
				RetryDeadlineUtc = now.Add(RetryWindow),
				LastStatusWriteUtc = now,
				Status = CreateStatus(document, trigger, "Scheduled", "Waiting for Revit to finish activating the project.")
			};
			Requests[token] = scheduled;
		}
		WriteStatus(HostWorkspacePathResolver.ResolveRoot(), document, CloneStatus(scheduled.Status));
	}

	public static FamilyBrowserAutomaticModelCheckStatus GetStatus(Document document)
	{
		if (document == null)
		{
			return null;
		}
		lock (SyncRoot)
		{
			PendingRequest request;
			return Requests.TryGetValue(RuntimeHelpers.GetHashCode(document), out request) && request != null
				? CloneStatus(request.Status)
				: null;
		}
	}

	public static bool ProcessPending(UIApplication uiApplication)
	{
		PruneInvalidRequests();
		Document document = null;
		try
		{
			document = uiApplication == null || uiApplication.ActiveUIDocument == null
				? null
				: uiApplication.ActiveUIDocument.Document;
		}
		catch
		{
			return HasPendingRequest();
		}
		if (!CanInspect(document))
		{
			return HasPendingRequest();
		}

		PendingRequest request;
		FamilyBrowserAutomaticModelCheckStatus runningStatus;
		int token = RuntimeHelpers.GetHashCode(document);
		lock (SyncRoot)
		{
			if (!Requests.TryGetValue(token, out request) || request == null || request.Completed || request.Running)
			{
				return HasPendingRequestLocked();
			}
			request.IdleTicks++;
			if (request.IdleTicks < InitialIdleDelay || DateTime.UtcNow < request.NextAttemptUtc || SafeIsModifiable(document))
			{
				return true;
			}
			request.Running = true;
			request.Status.Status = "Running";
			request.Status.StartedAtUtc = UtcNowText();
			request.Status.Message = "Running automatic Current Model Check for the assigned trade.";
			request.Status.ProgressCurrent = 0;
			request.Status.ProgressTotal = 100;
			request.LastStatusWriteUtc = DateTime.UtcNow;
			runningStatus = CloneStatus(request.Status);
		}
		string statusWorkspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		WriteStatus(statusWorkspaceRoot, document, runningStatus);
		NotifyDashboard(document);

		ExecutionOutcome outcome;
		try
		{
			outcome = Execute(uiApplication, document, request.Status);
		}
		catch (Exception ex)
		{
			string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
			FamilyBrowserErrorHelp.WriteLog(workspaceRoot, "Automatic Current Model Check failed", ex, "Project=" + SafeTitle(document));
			outcome = new ExecutionOutcome
			{
				Status = CompleteStatus(request.Status, "Failed", ex.GetType().Name + ": " + ex.Message)
			};
		}

		FamilyBrowserAutomaticModelCheckStatus statusToWrite = null;
		lock (SyncRoot)
		{
			PendingRequest current;
			if (!Requests.TryGetValue(token, out current) || current == null)
			{
				return HasPendingRequestLocked();
			}
			current.Running = false;
			current.Status = outcome.Status ?? current.Status;
			DateTime now = DateTime.UtcNow;
			if (outcome.Retry && now < current.RetryDeadlineUtc)
			{
				current.Attempts++;
				current.IdleTicks = 0;
				current.NextAttemptUtc = now.Add(RetryInterval);
				current.Status.Status = "Waiting";
				if (current.Attempts == 1 || now - current.LastStatusWriteUtc >= StatusWriteInterval)
				{
					current.LastStatusWriteUtc = now;
					statusToWrite = CloneStatus(current.Status);
				}
			}
			else
			{
			if (outcome.Retry)
			{
				string lastReason = current.Status == null ? string.Empty : (current.Status.Message ?? string.Empty).Trim();
				string deferredMessage = "Automatic Current Model Check could not start within 30 minutes.";
				if (!string.IsNullOrWhiteSpace(lastReason))
				{
					deferredMessage += " Last reason: " + lastReason;
				}
				current.Status = CompleteStatus(current.Status, "Deferred", deferredMessage);
			}
				current.Completed = true;
				current.CompletedUtc = now;
				current.LastStatusWriteUtc = now;
				statusToWrite = CloneStatus(current.Status);
			}
		}
		if (statusToWrite != null)
		{
			WriteStatus(statusWorkspaceRoot, document, statusToWrite);
			NotifyDashboard(document);
		}
		return HasPendingRequest();
	}

	private static ExecutionOutcome Execute(UIApplication uiApplication, Document document, FamilyBrowserAutomaticModelCheckStatus status)
	{
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			status.Message = "The managed data folder is not available yet.";
			return new ExecutionOutcome { Retry = true, Status = status };
		}
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(document);
		FamilyBrowserFileGuardTarget target = FamilyBrowserFileGuardPathMatcher.FindMatchingTarget(policy == null ? null : policy.FileGuard, context);
		if (target == null)
		{
			return new ExecutionOutcome { Status = CompleteStatus(status, "NotApplicable", "This project is not registered in File Guard, so automatic Current Model Check was skipped.") };
		}
		string publicationReason;
		if (!ProjectSnapshotStore.CanPublishSharedProjectState(document, out publicationReason))
		{
			status.Message = publicationReason;
			return new ExecutionOutcome { Retry = true, Status = status };
		}

		string discipline = FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline(policy, target, allowLegacyFallback: false);
		FamilyBrowserStandardLibrarySlot slot = FamilyBrowserFileGuardDisciplineService.ResolveSlot(policy, discipline);
		status.Discipline = discipline ?? string.Empty;
		status.DisciplineLabel = slot == null ? (discipline ?? string.Empty) : FamilyBrowserStandardPolicyStore.ResolveSlotDisplayName(slot, true);
		if (slot == null)
		{
			return new ExecutionOutcome { Status = CompleteStatus(status, "ConfigurationRequired", "The guarded RVT has no valid assigned trade. Open Permissions / Guard and assign a registered trade.") };
		}

		string registrationPath = FamilyBrowserStandardPolicyStore.ResolveSlotRegistrationPath(workspaceRoot, slot);
		if (string.IsNullOrWhiteSpace(registrationPath) || !File.Exists(registrationPath))
		{
			return new ExecutionOutcome { Status = CompleteStatus(status, "StandardRequired", "The assigned trade has no registered standard RVT.") };
		}
		StandardLibraryRegistrationRecord registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registrationPath);
		string snapshotPath = FamilyBrowserStandardPolicyStore.ResolveSlotSnapshotPath(workspaceRoot, slot, registration);
		if (registration == null || string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
		{
			return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", "The assigned trade standard RVT needs a completed scan before automatic model checking can run.") };
		}
		registration.LastSnapshotPath = snapshotPath;
		status.StandardRvtPath = registration.ResolvedPath ?? registration.Locator ?? string.Empty;
		StandardLibrarySnapshot standardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(snapshotPath);
		if (standardSnapshot == null)
		{
			return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", "The assigned trade standard snapshot could not be read.") };
		}
		FamilyBrowserStandardRevisionState standardRevisionBeforeScan = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, true);
		if (standardRevisionBeforeScan == null || standardRevisionBeforeScan.BlocksStandardUse)
		{
			string reason = standardRevisionBeforeScan == null ? "The assigned trade standard revision could not be verified." : (standardRevisionBeforeScan.Reason ?? "The assigned trade standard changed.");
			return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", reason) };
		}

		ProjectTrackingDirtyMarker dirtyMarkerAtCheckStart = ProjectTrackingStoreService.LoadCurrentModelCheckMarker(document);
		ProjectScanCacheLoadResult cached = ProjectSnapshotStore.TryLoadLatestProjectScan(workspaceRoot, document, registration, standardSnapshot);
		if (dirtyMarkerAtCheckStart == null && cached != null && cached.Success)
		{
			ApplyCachedStatus(status, cached);
			FamilyBrowserNestedOnlyPlacementRuntimeService.SeedFromProjectSnapshot(document, policy, discipline, cached.Snapshot);
			NotifyDashboard(document);
			CompleteStatus(status, "CacheReused", "The saved Current Model Check already matches this project and standard revision.");
			return new ExecutionOutcome { Status = status };
		}

		FileStream scanLock = ProjectSnapshotStore.TryAcquireProjectPublicationLock(workspaceRoot, document);
		if (scanLock == null)
		{
			status.Message = "Another Revit session is checking this project. Waiting to reuse its result.";
			return new ExecutionOutcome { Retry = true, Status = status };
		}
		using (scanLock)
		{
			FamilyBrowserStandardRevisionState lockedStandardRevision = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, true);
			if (lockedStandardRevision == null || lockedStandardRevision.BlocksStandardUse)
			{
				string reason = lockedStandardRevision == null ? "The assigned trade standard revision could not be verified." : (lockedStandardRevision.Reason ?? "The assigned trade standard changed.");
				return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", reason) };
			}
			standardRevisionBeforeScan = lockedStandardRevision;
			dirtyMarkerAtCheckStart = ProjectTrackingStoreService.LoadCurrentModelCheckMarker(document);
			cached = ProjectSnapshotStore.TryLoadLatestProjectScan(workspaceRoot, document, registration, standardSnapshot);
			if (dirtyMarkerAtCheckStart == null && cached != null && cached.Success)
			{
				ApplyCachedStatus(status, cached);
				FamilyBrowserNestedOnlyPlacementRuntimeService.SeedFromProjectSnapshot(document, policy, discipline, cached.Snapshot);
				NotifyDashboard(document);
				CompleteStatus(status, "CacheReused", "Another session completed the matching Current Model Check, so its result was reused.");
				return new ExecutionOutcome { Status = status };
			}

			bool includeDeepProjectContent = ShouldRunPreciseProjectContentCheck(standardSnapshot);
			ProjectContentSnapshot projectSnapshot;
			string projectSnapshotPath;
			ProjectStandardComparisonReport report;
			FamilyBrowserAutomaticModelCheckProgressWindow progressWindow = FamilyBrowserAutomaticModelCheckProgressWindow.Begin(uiApplication, SafeTitle(document), status.DisciplineLabel);
			try
			{
				if (SafeIsModified(document))
				{
					status.Message = "Save or synchronize the project before automatic Current Model Check publishes a shared result.";
					return new ExecutionOutcome { Retry = true, Status = status };
				}
				ReportProgress(document, progressWindow, 3, 100, FamilyBrowserLanguageService.Text("Preparing the assigned standard and current project...", "지정된 표준과 현재 프로젝트를 준비하는 중..."));
				using (FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication))
				{
					string debugFolder = FingerprintDebugSignatureStore.CreateProjectRunFolder(workspaceRoot, SafeTitle(document), UtcNowText(), ProjectSnapshotStore.ResolveProjectIdentityPath(document));
					projectSnapshot = ProjectSnapshotCaptureService.Capture(document, null, includeDeepProjectContent, delegate(int current, int total, string message)
					{
						int mapped = MapProgress(current, total, 5, 80);
						ReportProgress(document, progressWindow, mapped, 100, message);
					}, debugFolder, dialogGuard);
				}
				if (SafeIsModified(document))
				{
					status.Message = "The project changed while automatic Current Model Check was running. Save or synchronize it, then the check will retry.";
					return new ExecutionOutcome { Retry = true, Status = status };
				}
				FamilyBrowserStandardRevisionState standardRevisionAfterScan = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, true);
				if (!FamilyBrowserStandardRevisionService.IsSameCurrentRevision(standardRevisionBeforeScan, standardRevisionAfterScan))
				{
					string reason = standardRevisionAfterScan != null && standardRevisionAfterScan.BlocksStandardUse
						? (standardRevisionAfterScan.Reason ?? "The assigned trade standard changed.")
						: "The Standard RVT revision changed while automatic Current Model Check was running. No shared result was published.";
					return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", reason) };
				}
				ReportProgress(document, progressWindow, 84, 100, FamilyBrowserLanguageService.Text("Saving the current project snapshot...", "현재 프로젝트 스냅샷을 저장하는 중..."));
				projectSnapshotPath = ProjectSnapshotStore.Save(workspaceRoot, projectSnapshot, document);
				ProjectTrackingCatalog trackingCatalog = ProjectTrackingStoreService.Load(document);
				bool compareDetailedComponents = FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy);
				ReportProgress(document, progressWindow, 90, 100, FamilyBrowserLanguageService.Text("Comparing the project with the assigned standard...", "현재 프로젝트와 지정된 표준을 비교하는 중..."));
				report = ProjectStandardComparisonService.BuildReport(registration, snapshotPath, standardSnapshot, projectSnapshotPath, projectSnapshot, trackingCatalog, compareDetailedComponents);
				FamilyBrowserStandardRevisionState standardRevisionBeforePublication = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, true);
				if (!FamilyBrowserStandardRevisionService.IsSameCurrentRevision(standardRevisionBeforeScan, standardRevisionBeforePublication))
				{
					string reason = standardRevisionBeforePublication != null && standardRevisionBeforePublication.BlocksStandardUse
						? (standardRevisionBeforePublication.Reason ?? "The assigned trade standard changed.")
						: "The Standard RVT revision changed while the comparison report was being prepared. No shared result was published.";
					return new ExecutionOutcome { Status = CompleteStatus(status, "StandardScanRequired", reason) };
				}
				string finalPublicationReason;
				if (!ProjectSnapshotStore.CanPublishSharedProjectState(document, out finalPublicationReason))
				{
					status.Message = finalPublicationReason;
					return new ExecutionOutcome { Retry = true, Status = status };
				}
				ReportProgress(document, progressWindow, 96, 100, FamilyBrowserLanguageService.Text("Publishing the verified comparison result...", "검증된 비교 결과를 저장하는 중..."));
			string thumbnailSourceId = ProjectSnapshotStore.BuildProjectThumbnailSourceId(workspaceRoot, document);
			if (report != null && report.Project != null && !string.IsNullOrWhiteSpace(thumbnailSourceId))
			{
				report.Project.ThumbnailSourceId = thumbnailSourceId;
				report.Project.ThumbnailFolder = FamilyThumbnailPreviewService.GetCacheFolder(workspaceRoot, thumbnailSourceId);
			}
			string comparisonPath = ProjectStandardComparisonStore.Save(workspaceRoot, report);
			ProjectSnapshotStore.SaveLatestProjectScan(workspaceRoot, document, registration, standardSnapshot, projectSnapshotPath, comparisonPath, projectSnapshot, report, thumbnailSourceId);
			FamilyBrowserProjectCatalogService.AcceptFromProjectSnapshot(workspaceRoot, document, projectSnapshot, "AutomaticModelCheck", currentUser);
			ProjectTrackingStoreService.ClearCurrentModelCheckRequired(document, dirtyMarkerAtCheckStart);
			FamilyBrowserNestedOnlyPlacementRuntimeService.SeedFromProjectSnapshot(document, policy, discipline, projectSnapshot);

			status.ProjectSnapshotPath = projectSnapshotPath ?? string.Empty;
			status.ComparisonReportPath = comparisonPath ?? string.Empty;
			status.LoadableFamilyCount = projectSnapshot == null || projectSnapshot.Summary == null ? 0 : projectSnapshot.Summary.LoadableFamilyCount;
			status.SystemTypeCount = projectSnapshot == null || projectSnapshot.Summary == null ? 0 : projectSnapshot.Summary.SystemTypeCount;
			status.UsedCachedResult = false;
			status.ProgressCurrent = 100;
			status.ProgressTotal = 100;
			CompleteStatus(status, "Completed", "Automatic Current Model Check completed against the assigned trade standard.");
			ReportProgress(document, progressWindow, 100, 100, FamilyBrowserLanguageService.Text("Automatic Current Model Check completed.", "자동 현재 모델 검사를 완료했습니다."));
			NotifyDashboard(document);
			return new ExecutionOutcome { Status = status };
			}
			finally
			{
				if (progressWindow != null)
				{
					progressWindow.Dispose();
				}
			}
		}
	}

	private static bool ShouldRunPreciseProjectContentCheck(StandardLibrarySnapshot standardSnapshot)
	{
		if (standardSnapshot == null)
		{
			return false;
		}
		if (string.Equals(standardSnapshot.SnapshotMode, "Precise", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		return standardSnapshot.LoadableFamilies != null && standardSnapshot.LoadableFamilies.Any(delegate(StandardLoadableFamilySnapshotItem item)
		{
			return item != null && (string.Equals(item.MetadataMode ?? string.Empty, "Precise", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrWhiteSpace(item.ContentSignatureDebugPath));
		});
	}

	private static FamilyBrowserProjectPolicyContext BuildProjectContext(Document document)
	{
		FamilyBrowserProjectPolicyContext context = new FamilyBrowserProjectPolicyContext
		{
			ProjectTitle = SafeTitle(document),
			ModelPath = SafePath(document)
		};
		try
		{
			context.IsWorkshared = document != null && document.IsWorkshared;
			if (context.IsWorkshared)
			{
				context.CentralPath = ProjectSnapshotStore.ResolveProjectIdentityPath(document);
			}
		}
		catch
		{
			context.IsWorkshared = false;
			context.CentralPath = string.Empty;
		}
		return context;
	}

	private static void WriteStatus(string workspaceRoot, Document document, FamilyBrowserAutomaticModelCheckStatus status)
	{
		string temporaryPath = string.Empty;
		try
		{
			string folder = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, document), "AutomaticModelCheck");
			Directory.CreateDirectory(folder);
			string path = Path.Combine(folder, "automatic-model-check-latest.json");
			temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
			byte[] payload = new UTF8Encoding(false).GetBytes(PlainJsonReportWriter.Serialize(status));
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				stream.Write(payload, 0, payload.Length);
				stream.Flush(true);
			}
			FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
			temporaryPath = string.Empty;
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(workspaceRoot, "Automatic Current Model Check status write failed", ex, "Project=" + SafeTitle(document));
		}
		finally
		{
			try
			{
				if (!string.IsNullOrWhiteSpace(temporaryPath) && File.Exists(temporaryPath))
				{
					File.Delete(temporaryPath);
				}
			}
			catch
			{
			}
		}
	}

	private static void ApplyCachedStatus(FamilyBrowserAutomaticModelCheckStatus status, ProjectScanCacheLoadResult cached)
	{
		status.UsedCachedResult = true;
		status.ProjectSnapshotPath = cached == null || cached.Record == null ? string.Empty : (cached.Record.ProjectSnapshotPath ?? string.Empty);
		status.ComparisonReportPath = cached == null || cached.Record == null ? string.Empty : (cached.Record.ComparisonReportPath ?? string.Empty);
		status.LoadableFamilyCount = cached == null || cached.Snapshot == null || cached.Snapshot.Summary == null ? 0 : cached.Snapshot.Summary.LoadableFamilyCount;
		status.SystemTypeCount = cached == null || cached.Snapshot == null || cached.Snapshot.Summary == null ? 0 : cached.Snapshot.Summary.SystemTypeCount;
		status.ProgressCurrent = 100;
		status.ProgressTotal = 100;
	}

	private static void UpdateRunningMessage(Document document, string message, int current, int total)
	{
		if (document == null)
		{
			return;
		}
		FamilyBrowserAutomaticModelCheckStatus statusToWrite = null;
		lock (SyncRoot)
		{
			PendingRequest request;
			if (Requests.TryGetValue(RuntimeHelpers.GetHashCode(document), out request) && request != null && request.Status != null)
			{
				request.Status.Message = (message ?? string.Empty) + (total > 0 ? " " + current.ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) : string.Empty);
				request.Status.ProgressCurrent = Math.Max(0, current);
				request.Status.ProgressTotal = Math.Max(1, total);
				DateTime now = DateTime.UtcNow;
				if (now - request.LastStatusWriteUtc >= StatusWriteInterval)
				{
					request.LastStatusWriteUtc = now;
					statusToWrite = CloneStatus(request.Status);
				}
			}
		}
		if (statusToWrite != null)
		{
			WriteStatus(HostWorkspacePathResolver.ResolveRoot(), document, statusToWrite);
			NotifyDashboard(document);
		}
	}

	private static void ReportProgress(Document document, FamilyBrowserAutomaticModelCheckProgressWindow progressWindow, int current, int total, string message)
	{
		UpdateRunningMessage(document, message, current, total);
		if (progressWindow != null)
		{
			progressWindow.Report(current, total, message);
		}
	}

	private static int MapProgress(int current, int total, int start, int end)
	{
		int safeTotal = Math.Max(1, total);
		double ratio = Math.Max(0.0, Math.Min(1.0, (double)Math.Max(0, current) / safeTotal));
		return Math.Max(start, Math.Min(end, start + (int)Math.Round(ratio * Math.Max(0, end - start))));
	}

	private static void NotifyDashboard(Document document)
	{
		try
		{
			FamilyBrowserDashboardModelessRuntime.NotifyActiveDocumentChanged(document);
		}
		catch
		{
		}
	}

	private static FamilyBrowserAutomaticModelCheckStatus CreateStatus(Document document, string trigger, string status, string message)
	{
		return new FamilyBrowserAutomaticModelCheckStatus
		{
			ProjectTitle = SafeTitle(document),
			ProjectPath = SafePath(document),
			CentralPath = SafeIdentityPath(document),
			Trigger = trigger ?? string.Empty,
			Status = status ?? string.Empty,
			Message = message ?? string.Empty,
			ScheduledAtUtc = UtcNowText()
		};
	}

	private static FamilyBrowserAutomaticModelCheckStatus CompleteStatus(FamilyBrowserAutomaticModelCheckStatus status, string state, string message)
	{
		status = status ?? new FamilyBrowserAutomaticModelCheckStatus();
		status.Status = state ?? string.Empty;
		status.Message = message ?? string.Empty;
		status.CompletedAtUtc = UtcNowText();
		return status;
	}

	private static FamilyBrowserAutomaticModelCheckStatus CloneStatus(FamilyBrowserAutomaticModelCheckStatus source)
	{
		if (source == null)
		{
			return null;
		}
		return new FamilyBrowserAutomaticModelCheckStatus
		{
			SchemaVersion = source.SchemaVersion,
			ProjectTitle = source.ProjectTitle,
			ProjectPath = source.ProjectPath,
			CentralPath = source.CentralPath,
			Discipline = source.Discipline,
			DisciplineLabel = source.DisciplineLabel,
			StandardRvtPath = source.StandardRvtPath,
			Trigger = source.Trigger,
			Status = source.Status,
			Message = source.Message,
			ScheduledAtUtc = source.ScheduledAtUtc,
			StartedAtUtc = source.StartedAtUtc,
			CompletedAtUtc = source.CompletedAtUtc,
			UsedCachedResult = source.UsedCachedResult,
			ProjectSnapshotPath = source.ProjectSnapshotPath,
			ComparisonReportPath = source.ComparisonReportPath,
			LoadableFamilyCount = source.LoadableFamilyCount,
			SystemTypeCount = source.SystemTypeCount,
			ProgressCurrent = source.ProgressCurrent,
			ProgressTotal = source.ProgressTotal
		};
	}

	private static bool HasPendingRequest()
	{
		lock (SyncRoot)
		{
			return HasPendingRequestLocked();
		}
	}

	private static bool HasPendingRequestLocked()
	{
		return Requests.Values.Any(delegate(PendingRequest request) { return request != null && !request.Completed; });
	}

	private static void PruneInvalidRequests()
	{
		lock (SyncRoot)
		{
			foreach (int token in Requests
				.Where(delegate(KeyValuePair<int, PendingRequest> pair)
				{
					PendingRequest request = pair.Value;
					if (request == null || request.Document == null)
					{
						return true;
					}
					try
					{
						return !request.Document.IsValidObject;
					}
					catch
					{
						return true;
					}
				})
				.Select(delegate(KeyValuePair<int, PendingRequest> pair) { return pair.Key; })
				.ToList())
			{
				Requests.Remove(token);
			}
		}
	}

	private static bool CanInspect(Document document)
	{
		if (document == null)
		{
			return false;
		}
		try
		{
			return document.IsValidObject && !document.IsFamilyDocument && !string.IsNullOrWhiteSpace(document.PathName);
		}
		catch
		{
			return false;
		}
	}

	private static bool SafeIsModifiable(Document document)
	{
		try
		{
			return document != null && document.IsModifiable;
		}
		catch
		{
			return true;
		}
	}

	private static bool SafeIsModified(Document document)
	{
		try
		{
			return document == null || document.IsModified;
		}
		catch
		{
			return true;
		}
	}

	private static string SafeTitle(Document document)
	{
		try
		{
			return document == null ? string.Empty : (document.Title ?? string.Empty);
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string SafePath(Document document)
	{
		try
		{
			return document == null ? string.Empty : (document.PathName ?? string.Empty);
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string SafeIdentityPath(Document document)
	{
		try
		{
			return ProjectSnapshotStore.ResolveProjectIdentityPath(document) ?? string.Empty;
		}
		catch
		{
			return SafePath(document);
		}
	}

	private static string UtcNowText()
	{
		return DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
	}
}
