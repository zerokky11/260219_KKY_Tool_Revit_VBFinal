using Autodesk.Revit.DB;

public class LoadableFamilyLoadOptions : IFamilyLoadOptions
{
	private readonly bool _overwriteParameterValues;

	public LoadableFamilyLoadOptions(bool overwriteParameterValues)
	{
		_overwriteParameterValues = overwriteParameterValues;
	}

	public bool OnFamilyFound(bool familyInUse, ref bool overwriteParameterValues)
	{
		overwriteParameterValues = _overwriteParameterValues;
		return true;
	}

	public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, ref FamilySource source, ref bool overwriteParameterValues)
	{
		source = (FamilySource)1;
		overwriteParameterValues = _overwriteParameterValues;
		return true;
	}
}
