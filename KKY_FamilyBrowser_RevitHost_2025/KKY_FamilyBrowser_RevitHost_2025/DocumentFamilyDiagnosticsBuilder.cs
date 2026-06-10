using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2025;

internal sealed class DocumentFamilyDiagnosticsBuilder
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__1_002D0
	{
		public Family _0024VB_0024Local_family;

		public _Closure_0024__1_002D0(_Closure_0024__1_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_family = arg0._0024VB_0024Local_family;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__2()
		{
			return _0024VB_0024Local_family.IsEditable;
		}

		[SpecialName]
		internal bool _Lambda_0024__3()
		{
			return _0024VB_0024Local_family.IsInPlace;
		}

		[SpecialName]
		internal int _Lambda_0024__4()
		{
			return _0024VB_0024Local_family.GetFamilySymbolIds().Count;
		}
	}

	private DocumentFamilyDiagnosticsBuilder()
	{
	}

	public static DocumentFamilyDiagnosticsReport Build(Document doc)
	{
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		DocumentFamilyDiagnosticsReport report = new DocumentFamilyDiagnosticsReport
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			DocumentTitle = (doc.Title ?? string.Empty),
			DocumentPath = (doc.PathName ?? string.Empty),
			RevitVersion = doc.Application.VersionNumber
		};
		IOrderedEnumerable<Family> families = (from Family x in (IEnumerable)new FilteredElementCollector(doc).OfClass(typeof(Family))
			where x != null
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(((Element)x).Name), StringComparer.Ordinal);
		using (IEnumerator<Family> enumerator = families.GetEnumerator())
		{
			_Closure_0024__1_002D0 closure_0024__1_002D = default(_Closure_0024__1_002D0);
			while (enumerator.MoveNext())
			{
				closure_0024__1_002D = new _Closure_0024__1_002D0(closure_0024__1_002D);
				closure_0024__1_002D._0024VB_0024Local_family = enumerator.Current;
				report.Families.Add(new DocumentFamilyDiagnosticsItem
				{
					FamilyName = (((Element)closure_0024__1_002D._0024VB_0024Local_family).Name ?? string.Empty),
					CategoryName = ResolveCategoryName(closure_0024__1_002D._0024VB_0024Local_family),
					IsEditable = SafeBool(closure_0024__1_002D._Lambda_0024__2),
					IsInPlace = SafeBool(closure_0024__1_002D._Lambda_0024__3),
					IsShared = ResolveIsShared(closure_0024__1_002D._0024VB_0024Local_family),
					TypeCount = SafeInt(closure_0024__1_002D._Lambda_0024__4),
					UniqueId = (((Element)closure_0024__1_002D._0024VB_0024Local_family).UniqueId ?? string.Empty)
				});
			}
		}
		report.Summary = new DocumentFamilyDiagnosticsSummary
		{
			FamilyCount = report.Families.Count,
			EditableFamilyCount = report.Families.Where([SpecialName] (DocumentFamilyDiagnosticsItem x) => x.IsEditable).Count(),
			InPlaceFamilyCount = report.Families.Where([SpecialName] (DocumentFamilyDiagnosticsItem x) => x.IsInPlace).Count(),
			SharedFamilyCount = report.Families.Where([SpecialName] (DocumentFamilyDiagnosticsItem x) => x.IsShared).Count()
		};
		return report;
	}

	private static string ResolveCategoryName(Family family)
	{
		string ResolveCategoryName;
		try
		{
			Category familyCategory = family.FamilyCategory;
			ResolveCategoryName = ((familyCategory != null) ? familyCategory.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static bool ResolveIsShared(Family family)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Invalid comparison between Unknown and I4
		try
		{
			Parameter parameterValue = ((Element)family)[(BuiltInParameter)(-1012834)];
			if (parameterValue != null && (int)parameterValue.StorageType == 1)
			{
				return parameterValue.AsInteger() != 0;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static bool SafeBool(Func<bool> getter)
	{
		bool SafeBool;
		try
		{
			SafeBool = getter();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeBool = false;
			ProjectData.ClearProjectError();
		}
		return SafeBool;
	}

	private static int SafeInt(Func<int> getter)
	{
		int SafeInt;
		try
		{
			SafeInt = getter();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeInt = 0;
			ProjectData.ClearProjectError();
		}
		return SafeInt;
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
