using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2027;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdApplyStandardLoadableFamilies : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Document standardDoc = null;
		bool shouldCloseStandardDoc = false;
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before applying standard families.", "표준 패밀리를 적용하기 전에 프로젝트 RVT를 먼저 열어주세요.");
				Execute = Result.Cancelled;
			}
			else
			{
				string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
				FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, Environment.UserName);
				string registryPath = FamilyBrowserStandardPolicyStore.ResolveEffectiveRegistrationPath(workspaceRoot, policy);
				if (!File.Exists(registryPath))
				{
					message = T("No registered standard RVT was found. Connect the approved standard RVT first.", "등록된 표준 RVT를 찾을 수 없습니다. 승인된 표준 RVT를 먼저 연결하세요.");
					Execute = Result.Failed;
				}
				else
				{
					StandardLibraryRegistrationRecord registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registryPath);
					if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
					{
						message = T("The standard snapshot could not be found. Refresh or reconnect the standard RVT.", "표준 스냅샷을 찾을 수 없습니다. 표준 RVT를 새로고침하거나 다시 연결하세요.");
						Execute = Result.Failed;
					}
					else
					{
						StandardLibrarySnapshot standardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
						ProjectTrackingCatalog trackingCatalog = ProjectTrackingStoreService.Load(uiDoc.Document);
						ProjectContentSnapshot projectSnapshot = ProjectSnapshotCaptureService.Capture(uiDoc.Document);
						string projectSnapshotPath = ProjectSnapshotStore.Save(workspaceRoot, projectSnapshot);
						ProjectStandardComparisonReport comparisonReport = ProjectStandardComparisonService.BuildReport(registration, registration.LastSnapshotPath, standardSnapshot, projectSnapshotPath, projectSnapshot, trackingCatalog, FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy));
						string comparisonPath = ProjectStandardComparisonStore.Save(workspaceRoot, comparisonReport);
						LoadableFamilySyncPlan plan = LoadableFamilySyncPlannerService.BuildPlan(comparisonReport, comparisonPath);
						if (checked(plan.Summary.LoadCount + plan.Summary.ReloadCount) > 0)
						{
							EnsureTargetDiffersFromStandard(uiDoc.Document, registration);
							standardDoc = StandardLibraryDocumentResolver.OpenRegisteredDocument(commandData.Application.Application, registration, ref shouldCloseStandardDoc);
						}
						LoadableFamilySyncExecutionReport executionReport = LoadableFamilySyncExecutionService.Execute(uiDoc.Document, standardDoc, plan);
						executionReport.ComparisonPath = comparisonPath;
						if (executionReport.Items.Any([SpecialName] (LoadableFamilySyncExecutionItem x) =>
						{
							string a = Normalize(x.Outcome);
							return string.Equals(a, "loaded", StringComparison.Ordinal) || string.Equals(a, "reloaded", StringComparison.Ordinal) || string.Equals(a, "stamprefreshpending", StringComparison.Ordinal);
						}))
						{
							ProjectContentSnapshot postSnapshot = ProjectSnapshotCaptureService.Capture(uiDoc.Document);
							string postSnapshotPath = ProjectSnapshotStore.Save(workspaceRoot, postSnapshot);
							ProjectTrackingCatalog refreshedTrackingCatalog = ProjectTrackingStampService.BuildCatalog(registration, standardSnapshot, postSnapshot, FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy));
							using (Transaction tx = new Transaction(uiDoc.Document, "KKY Family Browser Refresh Tracking"))
							{
								tx.Start();
								ProjectTrackingStoreService.Save(uiDoc.Document, refreshedTrackingCatalog);
								tx.Commit();
							}
							executionReport.TrackingPath = ProjectTrackingOutputStore.Save(workspaceRoot, refreshedTrackingCatalog);
							foreach (LoadableFamilySyncExecutionItem item in executionReport.Items)
							{
								if (string.Equals(Normalize(item.Outcome), "stamprefreshpending", StringComparison.Ordinal))
								{
									item.Outcome = "TrackingRefreshed";
									item.Details = T("Tracking metadata refreshed without reloading the family.", "패밀리를 재로드하지 않고 추적 정보만 새로고침했습니다.");
								}
							}
							executionReport.Summary.TrackingRefreshedCount = executionReport.Items.Where([SpecialName] (LoadableFamilySyncExecutionItem x) =>
							{
								string a = Normalize(x.Outcome);
								return string.Equals(a, "loaded", StringComparison.Ordinal) || string.Equals(a, "reloaded", StringComparison.Ordinal) || string.Equals(a, "trackingrefreshed", StringComparison.Ordinal);
							}).Count();
							ProjectStandardComparisonReport postComparisonReport = ProjectStandardComparisonService.BuildReport(registration, registration.LastSnapshotPath, standardSnapshot, postSnapshotPath, postSnapshot, refreshedTrackingCatalog, FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy));
							executionReport.PostComparisonPath = ProjectStandardComparisonStore.Save(workspaceRoot, postComparisonReport);
						}
						string executionPath = LoadableFamilySyncStore.Save(workspaceRoot, executionReport);
						string compareAfterText = (string.IsNullOrWhiteSpace(executionReport.PostComparisonPath) ? T("(not written)", "(기록 안 됨)") : executionReport.PostComparisonPath);
						string trackingText = (string.IsNullOrWhiteSpace(executionReport.TrackingPath) ? T("(not written)", "(기록 안 됨)") : executionReport.TrackingPath);
						FamilyBrowserResultDialog.Show(T("Apply Standard Families", "표준 패밀리 적용"), T("Standard family apply completed.", "표준 패밀리 적용이 완료되었습니다.") + "\r\n\r\n" + T("Project", "프로젝트") + ": " + uiDoc.Document.Title + "\r\n" + T("Standard", "표준") + ": " + registration.DisplayName + "\r\n" + T("Loaded", "로드") + ": " + executionReport.Summary.LoadedCount + "\r\n" + T("Reloaded", "재로드") + ": " + executionReport.Summary.ReloadedCount + "\r\n" + T("Tracking refreshed", "추적 갱신") + ": " + executionReport.Summary.TrackingRefreshedCount + "\r\n" + T("Blocked", "차단") + ": " + executionReport.Summary.BlockedCount + "\r\n" + T("Failed", "실패") + ": " + executionReport.Summary.FailedCount + "\r\n" + T("Skipped", "건너뜀") + ": " + executionReport.Summary.SkippedCount + "\r\n" + T("Compare before", "적용 전 비교") + ": " + comparisonPath + "\r\n" + T("Compare after", "적용 후 비교") + ": " + compareAfterText + "\r\n" + T("Tracking", "추적") + ": " + trackingText + "\r\n" + T("Execution", "실행 결과") + ": " + executionPath);
						Execute = Result.Succeeded;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			message = FamilyBrowserCommandError.ToExternalCommandMessage("Family Browser", ex2);
			Execute = Result.Failed;
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (shouldCloseStandardDoc && standardDoc != null)
			{
				try
				{
					standardDoc.Close(saveModified: false);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		return Execute;
	}

	Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Execute
		return this.Execute(commandData, ref message, elements);
	}

	private static void EnsureTargetDiffersFromStandard(Document targetDocument, StandardLibraryRegistrationRecord registration)
	{
		if (targetDocument != null && !string.IsNullOrWhiteSpace(targetDocument.PathName))
		{
			string fullPath = Path.GetFullPath(targetDocument.PathName);
			string standardPath = Path.GetFullPath(registration.ResolvedPath);
			if (string.Equals(fullPath, standardPath, StringComparison.OrdinalIgnoreCase))
			{
				throw new InvalidOperationException(T("The active project is the registered standard RVT itself. Open a different target project before applying standard families.", "현재 열린 프로젝트가 등록된 표준 RVT입니다. 표준 패밀리를 적용할 대상 프로젝트를 따로 열어주세요."));
			}
		}
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
