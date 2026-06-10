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
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Expected O, but got Unknown
		List<StandardFamilyParameterSnapshotItem> result = new List<StandardFamilyParameterSnapshotItem>();
		if (familyDoc == null || familyDoc.FamilyManager == null)
		{
			return result;
		}
		FamilyManager manager = familyDoc.FamilyManager;
		string currentTypeName = ResolveCurrentFamilyTypeName(manager);
		try
		{
			IEnumerator enumerator = manager.Parameters.GetEnumerator();
			try
			{
				_Closure_0024__1_002D0 closure_0024__1_002D = default(_Closure_0024__1_002D0);
				while (enumerator.MoveNext())
				{
					closure_0024__1_002D = new _Closure_0024__1_002D0(closure_0024__1_002D);
					closure_0024__1_002D._0024VB_0024Local_familyParameter = (FamilyParameter)enumerator.Current;
					if (ShouldCaptureFamilyManagerParameter(closure_0024__1_002D._0024VB_0024Local_familyParameter))
					{
						bool isInstance = SafeBool(closure_0024__1_002D._Lambda_0024__0);
						result.Add(new StandardFamilyParameterSnapshotItem
						{
							Scope = (isInstance ? "Instance" : "Type"),
							TypeName = (isInstance ? string.Empty : currentTypeName),
							Name = ResolveFamilyParameterName(closure_0024__1_002D._0024VB_0024Local_familyParameter),
							StorageType = ResolveFamilyParameterStorageTypeName(closure_0024__1_002D._0024VB_0024Local_familyParameter),
							ValuePreview = ResolveFamilyParameterValue(familyDoc, manager, closure_0024__1_002D._0024VB_0024Local_familyParameter),
							Formula = ResolveFamilyParameterFormula(closure_0024__1_002D._0024VB_0024Local_familyParameter),
							IsInstance = isInstance,
							IsReadOnly = false,
							IsShared = IsSharedFamilyParameter(closure_0024__1_002D._0024VB_0024Local_familyParameter),
							ParameterId = ResolveFamilyParameterId(closure_0024__1_002D._0024VB_0024Local_familyParameter),
							ExternalGuid = ResolveFamilyParameterExternalGuid(closure_0024__1_002D._0024VB_0024Local_familyParameter)
						});
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
		return FamilyParameterSnapshotNormalizationService.DeduplicateDefinitions(result);
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

	private static string ResolveFamilyParameterName(FamilyParameter familyParameter)
	{
		string ResolveFamilyParameterName;
		try
		{
			object obj;
			if (familyParameter == null)
			{
				obj = null;
			}
			else
			{
				Definition definition = familyParameter.Definition;
				obj = ((definition != null) ? definition.Name : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			ResolveFamilyParameterName = (string)obj;
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
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		string ResolveFamilyParameterStorageTypeName;
		try
		{
			ResolveFamilyParameterStorageTypeName = ((Enum)familyParameter.StorageType/*cast due to .constrained prefix*/).ToString();
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
			if (familyParameter != null && familyParameter.Id != null)
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
			Definition obj = ((familyParameter != null) ? familyParameter.Definition : null);
			ExternalDefinition externalDefinition = (ExternalDefinition)(object)((obj is ExternalDefinition) ? obj : null);
			if (externalDefinition != null)
			{
				return externalDefinition.GUID.ToString("D");
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
				PropertyInfo propertyInfo = ((object)familyParameter).GetType().GetProperty("Formula");
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

	private static string ResolveFamilyParameterValue(Document familyDoc, FamilyManager manager, FamilyParameter familyParameter)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Expected I4, but got Unknown
		string ResolveFamilyParameterValue;
		try
		{
			if (manager == null || manager.CurrentType == null || familyParameter == null)
			{
				ResolveFamilyParameterValue = string.Empty;
			}
			else
			{
				FamilyType familyType = manager.CurrentType;
				StorageType storageType = familyParameter.StorageType;
				switch (storageType - 1)
				{
				case 2:
					ResolveFamilyParameterValue = NormalizeMultiline(familyType.AsString(familyParameter));
					break;
				case 1:
				{
					object valueObject = familyType.AsDouble(familyParameter);
					ResolveFamilyParameterValue = ((valueObject != null) ? Convert.ToDouble(RuntimeHelpers.GetObjectValue(valueObject), CultureInfo.InvariantCulture).ToString("G17", CultureInfo.InvariantCulture) : string.Empty);
					break;
				}
				case 0:
				{
					object valueObject2 = familyType.AsInteger(familyParameter);
					ResolveFamilyParameterValue = ((valueObject2 != null) ? Convert.ToInt32(RuntimeHelpers.GetObjectValue(valueObject2), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) : string.Empty);
					break;
				}
				case 3:
				{
					ElementId id = familyType.AsElementId(familyParameter);
					if (id == null || id == ElementId.InvalidElementId)
					{
						ResolveFamilyParameterValue = string.Empty;
						break;
					}
					Element referenced = ((familyDoc == null) ? null : familyDoc.GetElement(id));
					ResolveFamilyParameterValue = ((referenced == null) ? RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture) : (((object)referenced).GetType().Name + ":" + ResolveElementName(referenced)));
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
