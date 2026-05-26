using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    internal static class ElementIdCompatExtensions
    {
        internal static int CompatIntegerValue(this ElementId id)
        {
            if (id == null)
            {
                return 0;
            }

#if NET10_0_OR_GREATER
            return unchecked((int)id.Value);
#else
            return id.IntegerValue;
#endif
        }

        internal static int CompatIntegerValue(this WorksetId id)
        {
            if (id == null)
            {
                return 0;
            }

            return id.IntegerValue;
        }
    }
}
