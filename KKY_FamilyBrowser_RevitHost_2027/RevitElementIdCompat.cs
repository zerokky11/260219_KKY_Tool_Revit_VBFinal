using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

[StandardModule]
internal static class RevitElementIdCompat
{
	internal static int CompatIntegerValue(this ElementId id)
	{
		if ((object)id == null)
		{
			return 0;
		}
		return checked((int)id.Value);
	}
}
