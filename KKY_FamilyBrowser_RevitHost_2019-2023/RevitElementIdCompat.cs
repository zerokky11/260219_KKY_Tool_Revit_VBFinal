using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

[StandardModule]
internal sealed class RevitElementIdCompat
{
	internal static int CompatIntegerValue(this ElementId id)
	{
		if (id == null)
		{
			return 0;
		}
		return id.IntegerValue;
	}
}
