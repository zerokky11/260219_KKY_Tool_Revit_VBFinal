using Autodesk.Revit.DB;

public class CopyPasteUseDestinationTypesHandler : IDuplicateTypeNamesHandler
{
	public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
	{
		return (DuplicateTypeAction)1;
	}
}
