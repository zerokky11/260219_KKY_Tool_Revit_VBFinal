using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

public sealed class SystemTypeSemanticFingerprintCatalogService
{
	private SystemTypeSemanticFingerprintCatalogService()
	{
	}

	public static Dictionary<string, string> BuildMap(Document doc, string sourceId, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = true, Action<int, int, string> progress = null)
	{
		Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<string, SystemTypeSemanticSnapshot> semanticSnapshots = BuildSnapshotMap(doc, sourceId, loadableContentFingerprintCache, includeDeepLoadableContent, progress);
		foreach (SystemTypeSemanticSnapshot item in semanticSnapshots.Values.OrderBy([SpecialName] (SystemTypeSemanticSnapshot x) => SystemTypeIdentityService.BuildKey(x.SystemFamilyKind, x.CategoryName, x.TypeName), StringComparer.Ordinal))
		{
			result[SystemTypeIdentityService.BuildKey(item.SystemFamilyKind, item.CategoryName, item.TypeName)] = SystemTypeFingerprintService.Compute(item);
		}
		return result;
	}

	public static Dictionary<string, SystemTypeSemanticSnapshot> BuildSnapshotMap(Document doc, string sourceId, IDictionary<string, string> loadableContentFingerprintCache = null, bool includeDeepLoadableContent = true, Action<int, int, string> progress = null)
	{
		Dictionary<string, SystemTypeSemanticSnapshot> result = new Dictionary<string, SystemTypeSemanticSnapshot>(StringComparer.Ordinal);
		if (doc == null)
		{
			return result;
		}
		SystemTypeCatalogSnapshot catalog = SystemTypeSemanticCaptureService.Capture(doc, sourceId, loadableContentFingerprintCache, includeDeepLoadableContent, progress);
		foreach (SystemTypeSemanticSnapshot item in catalog.Types.OrderBy([SpecialName] (SystemTypeSemanticSnapshot x) => SystemTypeIdentityService.BuildKey(x.SystemFamilyKind, x.CategoryName, x.TypeName), StringComparer.Ordinal))
		{
			result[SystemTypeIdentityService.BuildKey(item.SystemFamilyKind, item.CategoryName, item.TypeName)] = item;
		}
		return result;
	}
}
