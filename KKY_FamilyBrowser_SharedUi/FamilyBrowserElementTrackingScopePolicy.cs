using System;
using System.Collections.Generic;
using System.Text;

public static class FamilyBrowserElementTrackingScopePolicy
{
    private static readonly HashSet<string> AuxiliaryElementClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "GraphicsStyle",
        "MEPSystem",
        "ElectricalSystem",
        "PipingSystem",
        "MechanicalSystem",
        "CableTrayRun",
        "ConduitRun"
    };

    private static readonly HashSet<string> AuxiliaryCategoryIds = new HashSet<string>(StringComparer.Ordinal)
    {
        "-2000288", // Center Lines
        "-2008045", // Pipe center lines
        "-2008072", // Pipe fitting center lines
        "-2008051", // Flex pipe center lines
        "-2008001", // Duct center lines
        "-2008066", // Duct fitting center lines
        "-2008021", // Flex duct center lines
        "-2008139", // Conduit center lines
        "-2008141", // Conduit fitting center lines
        "-2008136", // Cable tray center lines
        "-2008140", // Cable tray fitting center lines
        "-2008214", // Fabrication containment center lines
        "-2008196", // Fabrication ductwork center lines
        "-2008210", // Fabrication pipework center lines
        "-2008150", // Cable tray run containers
        "-2008149", // Conduit run containers
        "-2008037", // Electrical circuits
        "-2008152", // Electrical internal circuits
        "-2008043", // Piping systems
        "-2008015", // Duct systems
        "-2008158", // Piping system references
        "-2008159", // Piping system reference visibility
        "-2008156", // Duct system references
        "-2008157"  // Duct system reference visibility
    };

    private static readonly HashSet<string> AuxiliaryCategoryNames = new HashSet<string>(StringComparer.Ordinal)
    {
        "centerline",
        "centerlines",
        "centreline",
        "centrelines",
        "중심선",
        "electricalcircuit",
        "electricalcircuits",
        "electricalinternalcircuit",
        "electricalinternalcircuits",
        "전기회로",
        "pipesystem",
        "pipesystems",
        "pipingsystem",
        "pipingsystems",
        "배관시스템",
        "ductsystem",
        "ductsystems",
        "덕트시스템",
        "cabletrayrun",
        "cabletrayruns",
        "케이블트레이런",
        "conduitrun",
        "conduitruns",
        "전선관런",
        "pipingsystemreference",
        "pipingsystemreferences",
        "ductsystemreference",
        "ductsystemreferences"
    };

    public static bool IsAuxiliarySupportRecord(string elementClass, string categoryId, string categoryName)
    {
        return IsAuxiliarySupportRecord(elementClass, categoryId, categoryName, false);
    }

    public static bool IsAuxiliarySupportRecord(string elementClass, string categoryId, string categoryName, bool isElementType)
    {
        if (isElementType)
        {
            return false;
        }
        string className = (elementClass ?? string.Empty).Trim();
        if (AuxiliaryElementClasses.Contains(className))
        {
            return true;
        }

        string stableCategoryId = (categoryId ?? string.Empty).Trim();
        if (AuxiliaryCategoryIds.Contains(stableCategoryId))
        {
            return true;
        }

        string normalizedCategoryName = NormalizeCategoryName(categoryName);
        return normalizedCategoryName.Length > 0 && AuxiliaryCategoryNames.Contains(normalizedCategoryName);
    }

    private static string NormalizeCategoryName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        StringBuilder normalized = new StringBuilder(value.Length);
        foreach (char character in value.Trim().ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(character))
            {
                normalized.Append(character);
            }
        }
        return normalized.ToString();
    }
}
