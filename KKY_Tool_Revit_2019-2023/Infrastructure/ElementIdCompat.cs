using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Infrastructure
{
    internal static class ElementIdCompat
    {
        public static int IntValue(this ElementId id)
        {
            if (id == null) return -1;
#if REVIT2025
            return (int)id.Value;
#else
            return id.IntegerValue;
#endif
        }

        public static ElementId FromInt(int id)
        {
#if REVIT2025
            return new ElementId((long)id);
#else
            return new ElementId(id);
#endif
        }

        public static ElementId FromLong(long id)
        {
#if REVIT2025
            return new ElementId(id);
#else
            return new ElementId((int)id);
#endif
        }
    }
}
