using System.Collections.Generic;

public class SystemTypePreflightReport
{
	public string GeneratedAtUtc { get; set; }

	public string StandardDisplayName { get; set; }

	public string ProjectDocumentTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public SystemTypePreflightSummary Summary { get; set; }

	public SystemTypeCatalogSnapshot StandardCatalog { get; set; }

	public SystemTypeCatalogSnapshot ProjectCatalog { get; set; }

	public RoutingFamilyCatalogSnapshot ProjectRoutingFamilies { get; set; }

	public SystemTypeSyncPlan SyncPlan { get; set; }

	public RoutingDependencyPreflightPlan DependencyPlan { get; set; }

	public SystemSyncExecutionPlan ExecutionPlan { get; set; }

	public List<SystemTypePreflightDiagnostic> Diagnostics { get; set; }

	public SystemTypePreflightReport()
	{
		GeneratedAtUtc = string.Empty;
		StandardDisplayName = string.Empty;
		ProjectDocumentTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		Summary = new SystemTypePreflightSummary();
		StandardCatalog = new SystemTypeCatalogSnapshot();
		ProjectCatalog = new SystemTypeCatalogSnapshot();
		ProjectRoutingFamilies = new RoutingFamilyCatalogSnapshot();
		SyncPlan = new SystemTypeSyncPlan();
		DependencyPlan = new RoutingDependencyPreflightPlan();
		ExecutionPlan = new SystemSyncExecutionPlan();
		Diagnostics = new List<SystemTypePreflightDiagnostic>();
	}
}
