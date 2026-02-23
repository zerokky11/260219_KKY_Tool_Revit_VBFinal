using Autodesk.Revit.DB;

namespace Autodesk.Revit.DB;

/// <summary>
/// Compatibility enum for code paths that still use BuiltInParameterGroup.
/// </summary>
public enum BuiltInParameterGroup
{
    PG_DATA,
    PG_TEXT,
    PG_GEOMETRY,
    PG_CONSTRAINTS,
    PG_IDENTITY_DATA,
    PG_MATERIALS,
    PG_GRAPHICS,
    PG_ANALYSIS,
    PG_GENERAL
}

public static class BuiltInParameterGroupCompat
{
    public static ForgeTypeId ToGroupTypeId(this BuiltInParameterGroup group) => MapGroup(group);

    public static BuiltInParameterGroup FromGroupTypeId(ForgeTypeId groupId) => ToBuiltInGroup(groupId);

    public static bool IsInGroup(this Definition def, BuiltInParameterGroup group)
    {
        if (def == null)
        {
            return false;
        }

        return ToBuiltInGroup(def.GetGroupTypeId()) == group;
    }

    private static ForgeTypeId MapGroup(BuiltInParameterGroup group)
    {
        return group switch
        {
            BuiltInParameterGroup.PG_TEXT => GroupTypeId.Text,
            BuiltInParameterGroup.PG_GEOMETRY => GroupTypeId.Geometry,
            BuiltInParameterGroup.PG_CONSTRAINTS => GroupTypeId.Constraints,
            BuiltInParameterGroup.PG_MATERIALS => GroupTypeId.Materials,
            BuiltInParameterGroup.PG_GRAPHICS => GroupTypeId.Graphics,
            _ => GroupTypeId.Data
        };
    }

    public static BuiltInParameterGroup ToBuiltInGroup(ForgeTypeId groupId)
    {
        if (groupId == null)
        {
            return BuiltInParameterGroup.PG_DATA;
        }

        if (groupId.Equals(GroupTypeId.Text)) return BuiltInParameterGroup.PG_TEXT;
        if (groupId.Equals(GroupTypeId.Geometry)) return BuiltInParameterGroup.PG_GEOMETRY;
        if (groupId.Equals(GroupTypeId.Constraints)) return BuiltInParameterGroup.PG_CONSTRAINTS;
        if (groupId.Equals(GroupTypeId.Materials)) return BuiltInParameterGroup.PG_MATERIALS;
        if (groupId.Equals(GroupTypeId.Graphics)) return BuiltInParameterGroup.PG_GRAPHICS;

        return BuiltInParameterGroup.PG_DATA;
    }
}

public static class FamilyManagerCompatExtensions
{
    public static FamilyParameter AddParameter(this FamilyManager fm, ExternalDefinition def, BuiltInParameterGroup group, bool isInstance)
    {
        ForgeTypeId gId = BuiltInParameterGroupCompat.ToGroupTypeId(group);
        return fm.AddParameter(def, gId, isInstance);
    }

    public static FamilyParameter ReplaceParameter(this FamilyManager fm, FamilyParameter param, ExternalDefinition def, BuiltInParameterGroup group, bool isInstance)
    {
        ForgeTypeId gId = BuiltInParameterGroupCompat.ToGroupTypeId(group);
        return fm.ReplaceParameter(param, def, gId, isInstance);
    }
}
