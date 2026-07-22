using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyDocumentParameterCaptureService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__1_002D0
	{
		public FamilyParameter _0024VB_0024Local_familyParameter;

		public _Closure_0024__1_002D0(_Closure_0024__1_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_familyParameter = arg0._0024VB_0024Local_familyParameter;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0()
		{
			return _0024VB_0024Local_familyParameter.IsInstance;
		}
	}

	private FamilyDocumentParameterCaptureService()
	{
	}

	public static List<StandardFamilyParameterSnapshotItem> Capture(Document familyDoc)
	{
		List<StandardFamilyParameterSnapshotItem> result = new List<StandardFamilyParameterSnapshotItem>();
		if (familyDoc == null || familyDoc.FamilyManager == null)
		{
			return result;
		}
		FamilyManager manager = familyDoc.FamilyManager;
		try
		{
			List<FamilyParameter> parameters = CaptureFamilyParameters(manager);
			List<FamilyType> familyTypes = CaptureFamilyTypes(manager);
			foreach (FamilyParameter familyParameter in parameters)
			{
				if (!ShouldCaptureFamilyManagerParameter(familyParameter))
				{
					continue;
				}
				bool isInstance = SafeBool(() => familyParameter.IsInstance);
				if (isInstance || familyTypes.Count == 0)
				{
					result.Add(BuildParameterSnapshot(familyDoc, manager, manager.CurrentType, familyParameter, isInstance, isInstance ? string.Empty : ResolveCurrentFamilyTypeName(manager)));
					continue;
				}
				foreach (FamilyType familyType in familyTypes)
				{
					string typeName = ResolveFamilyTypeName(familyType);
					if (string.IsNullOrWhiteSpace(typeName))
					{
						continue;
					}
					result.Add(BuildParameterSnapshot(familyDoc, manager, familyType, familyParameter, isInstance, typeName));
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return FamilyParameterSnapshotNormalizationService.DeduplicateDefinitionsAndTypeValues(result);
	}

	private static List<FamilyParameter> CaptureFamilyParameters(FamilyManager manager)
	{
		List<FamilyParameter> result = new List<FamilyParameter>();
		try
		{
			if (manager == null || manager.Parameters == null)
			{
				return result;
			}
			IEnumerator enumerator = manager.Parameters.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					if (enumerator.Current is FamilyParameter familyParameter)
					{
						result.Add(familyParameter);
					}
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static List<FamilyType> CaptureFamilyTypes(FamilyManager manager)
	{
		List<FamilyType> result = new List<FamilyType>();
		try
		{
			if (manager == null || manager.Types == null)
			{
				return result;
			}
			IEnumerator enumerator = manager.Types.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					if (enumerator.Current is FamilyType familyType)
					{
						result.Add(familyType);
					}
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static StandardFamilyParameterSnapshotItem BuildParameterSnapshot(Document familyDoc, FamilyManager manager, FamilyType familyType, FamilyParameter familyParameter, bool isInstance, string typeName)
	{
		return new StandardFamilyParameterSnapshotItem
		{
			Scope = (isInstance ? "Instance" : "Type"),
			TypeName = (isInstance ? string.Empty : (typeName ?? string.Empty)),
			Name = ResolveFamilyParameterName(familyParameter),
			StorageType = ResolveFamilyParameterStorageTypeName(familyParameter),
			ValuePreview = ResolveFamilyParameterValue(familyDoc, manager, familyType, familyParameter),
			Formula = ResolveFamilyParameterFormula(familyParameter),
			IsInstance = isInstance,
			IsReadOnly = false,
			IsShared = IsSharedFamilyParameter(familyParameter),
			ParameterId = ResolveFamilyParameterId(familyParameter),
			ExternalGuid = ResolveFamilyParameterExternalGuid(familyParameter)
		};
	}

	private static bool ShouldCaptureFamilyManagerParameter(FamilyParameter familyParameter)
	{
		if (familyParameter == null || familyParameter.Definition == null)
		{
			return false;
		}
		if (string.IsNullOrWhiteSpace(ResolveFamilyParameterName(familyParameter)))
		{
			return false;
		}
		int idValue = ResolveFamilyParameterIdInteger(familyParameter);
		if (idValue == -1002001)
		{
			return false;
		}
		if (idValue < 0 && !IsSharedFamilyParameter(familyParameter))
		{
			return false;
		}
		return true;
	}

	private static string ResolveCurrentFamilyTypeName(FamilyManager manager)
	{
		try
		{
			if (manager != null && manager.CurrentType != null)
			{
				return manager.CurrentType.Name ?? string.Empty;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveFamilyTypeName(FamilyType familyType)
	{
		try
		{
			if (familyType != null)
			{
				return familyType.Name ?? string.Empty;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveFamilyParameterName(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterName;
		try
		{
			ResolveFamilyParameterName = familyParameter?.Definition?.Name ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterName;
	}

	private static string ResolveFamilyParameterStorageTypeName(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterStorageTypeName;
		try
		{
			ResolveFamilyParameterStorageTypeName = familyParameter.StorageType.ToString();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterStorageTypeName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterStorageTypeName;
	}

	private static string ResolveFamilyParameterId(FamilyParameter familyParameter)
	{
		int idValue = ResolveFamilyParameterIdInteger(familyParameter);
		if (idValue == int.MinValue)
		{
			return string.Empty;
		}
		return idValue.ToString(CultureInfo.InvariantCulture);
	}

	private static int ResolveFamilyParameterIdInteger(FamilyParameter familyParameter)
	{
		try
		{
			if (familyParameter != null && (object)familyParameter.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(familyParameter.Id);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return int.MinValue;
	}

	private static bool IsSharedFamilyParameter(FamilyParameter familyParameter)
	{
		try
		{
			if (familyParameter != null && familyParameter.IsShared)
			{
				return true;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return !string.IsNullOrWhiteSpace(ResolveFamilyParameterExternalGuid(familyParameter));
	}

	private static string ResolveFamilyParameterExternalGuid(FamilyParameter familyParameter)
	{
		try
		{
			if (familyParameter?.Definition is ExternalDefinition { GUID: var gUID })
			{
				return gUID.ToString("D");
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string ResolveFamilyParameterFormula(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterFormula;
		if (familyParameter == null)
		{
			ResolveFamilyParameterFormula = string.Empty;
		}
		else
		{
			try
			{
				PropertyInfo propertyInfo = familyParameter.GetType().GetProperty("Formula");
				ResolveFamilyParameterFormula = (((object)propertyInfo != null) ? NormalizeMultiline(Convert.ToString(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(familyParameter, null))), CultureInfo.InvariantCulture)) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveFamilyParameterFormula = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveFamilyParameterFormula;
	}

	private static string ResolveFamilyParameterValue(Document familyDoc, FamilyManager manager, FamilyType familyType, FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterValue;
		try
		{
			FamilyType effectiveType = familyType ?? manager?.CurrentType;
			if (effectiveType == null || familyParameter == null)
			{
				ResolveFamilyParameterValue = string.Empty;
			}
			else
			{
				switch (familyParameter.StorageType)
				{
				case StorageType.String:
					ResolveFamilyParameterValue = NormalizeMultiline(effectiveType.AsString(familyParameter));
					break;
				case StorageType.Double:
				{
					object valueObject = effectiveType.AsDouble(familyParameter);
					ResolveFamilyParameterValue = ((valueObject != null) ? Convert.ToDouble(RuntimeHelpers.GetObjectValue(valueObject), CultureInfo.InvariantCulture).ToString("G17", CultureInfo.InvariantCulture) : string.Empty);
					break;
				}
				case StorageType.Integer:
				{
					object valueObject2 = effectiveType.AsInteger(familyParameter);
					ResolveFamilyParameterValue = ((valueObject2 != null) ? FamilyBrowserYesNoParameterFormatter.FormatInteger(familyParameter, Convert.ToInt32(RuntimeHelpers.GetObjectValue(valueObject2), CultureInfo.InvariantCulture)) : string.Empty);
					break;
				}
				case StorageType.ElementId:
				{
					ElementId id = effectiveType.AsElementId(familyParameter);
					if ((object)id == null || id == ElementId.InvalidElementId)
					{
						ResolveFamilyParameterValue = string.Empty;
						break;
					}
					Element referenced = familyDoc?.GetElement(id);
					ResolveFamilyParameterValue = ((referenced == null) ? RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture) : (referenced.GetType().Name + ":" + ResolveElementName(referenced)));
					break;
				}
				default:
					ResolveFamilyParameterValue = string.Empty;
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyParameterValue = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyParameterValue;
	}

	private static string ResolveElementName(Element element)
	{
		string ResolveElementName;
		if (element == null)
		{
			ResolveElementName = string.Empty;
		}
		else
		{
			try
			{
				ResolveElementName = element.Name ?? string.Empty;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveElementName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveElementName;
	}

	private static string NormalizeMultiline(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")
			.Trim();
	}

	private static bool SafeBool(Func<bool> reader)
	{
		bool SafeBool;
		try
		{
			SafeBool = reader();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeBool = false;
			ProjectData.ClearProjectError();
		}
		return SafeBool;
	}
}
