using Autodesk.Revit.DB;

public class LoadableFamilyLoadOptions : IFamilyLoadOptions
{
	private readonly bool _overwriteParameterValues;

	public LoadableFamilyLoadOptions(bool overwriteParameterValues)
	{
		_overwriteParameterValues = overwriteParameterValues;
	}

	public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
	{
		overwriteParameterValues = _overwriteParameterValues;
		return true;
	}

	public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
	{
		source = FamilySource.Family;
		overwriteParameterValues = _overwriteParameterValues;
		return true;
	}

}
