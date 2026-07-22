using Autodesk.Revit.DB;

public class CopyPasteUseDestinationTypesHandler : IDuplicateTypeNamesHandler
{
	public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
	{
		return DuplicateTypeAction.UseDestinationTypes;
	}

	DuplicateTypeAction IDuplicateTypeNamesHandler.OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
	{
		//ILSpy generated this explicit interface implementation from .override directive in OnDuplicateTypeNamesFound
		return this.OnDuplicateTypeNamesFound(args);
	}
}
