using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    public static class UnconnectedConnectorReviewService
    {
        private const string ReviewError = "\uC624\uB958";
        private const string ReviewPartialError = "\uC77C\uBD80\uC624\uB958";

        public sealed class Settings
        {
            public bool HasAllowedElementScope { get; set; }
            public List<int> AllowedElementIds { get; set; } = new List<int>();
            public string CommonTargetFilterText { get; set; } = string.Empty;
            public string CommonExcludeTargetFilterText { get; set; } = string.Empty;
            public List<string> ExtraParameterNames { get; set; } = new List<string>();
        }

        public sealed class ReviewRow
        {
            public string File { get; set; } = string.Empty;
            public string Item { get; set; } = string.Empty;
            public string ItemBase { get; set; } = string.Empty;
            public string Id { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Result { get; set; } = string.Empty;
            public string Content { get; set; } = string.Empty;
            public string Etc { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string IssueKind { get; set; } = "unconnected";
            public bool IsInformational { get; set; }
            public int ConnectorCount { get; set; }
            public int UnconnectedCount { get; set; }
            public Dictionary<string, string> ExtraParameterValues { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int TargetElementCount { get; set; }
            public int ConnectorCount { get; set; }
            public int ErrorCount { get; set; }
            public int FullErrorCount { get; set; }
            public int PartialErrorCount { get; set; }
            public int OkCount { get; set; }
            public bool CenterAxisEnabled { get; set; }
            public int CenterAxisTargetCount { get; set; }
            public int CenterAxisErrorCount { get; set; }
            public bool TapDepthEnabled { get; set; }
            public int TapDepthTargetCount { get; set; }
            public int TapDepthErrorCount { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public int TargetElementCount { get; set; }
            public int ConnectorCount { get; set; }
            public int ErrorCount { get; set; }
            public int FullErrorCount { get; set; }
            public int PartialErrorCount { get; set; }
            public int OkCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
        }

        private sealed class ConnectorOwner
        {
            public Element Element { get; set; }
            public List<Connector> Connectors { get; set; } = new List<Connector>();
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            settings = settings ?? new Settings();

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
            var allowedElementIds = new HashSet<int>((settings.AllowedElementIds ?? Enumerable.Empty<int>()).Where(id => id > 0));
            bool hasCommonScopeFilter = settings.HasAllowedElementScope;
            List<string> extraParameterNames = NormalizeExtraParameterNames(settings.ExtraParameterNames);

            progress?.Invoke(5d, "\uCEE4\uB125\uD130 \uC18C\uC720 \uAC1D\uCCB4 \uC218\uC9D1 \uC911");
            List<ConnectorOwner> owners = CollectConnectorOwners(doc)
                .Where(owner => owner != null && owner.Element != null)
                .Where(owner => owner.Connectors != null && owner.Connectors.Count > 0)
                .Where(owner => !ShouldExcludeElement(owner.Element))
                .Where(owner => !hasCommonScopeFilter || allowedElementIds.Contains(GetElementIdValue(owner.Element)))
                .OrderBy(owner => GetElementIdValue(owner.Element))
                .ToList();

            var result = new ReviewResult
            {
                File = safeFileLabel,
                TargetElementCount = owners.Count,
                ConnectorCount = owners.Sum(owner => owner.Connectors.Count)
            };

            int totalOwners = Math.Max(owners.Count, 1);
            for (int index = 0; index < owners.Count; index++)
            {
                ConnectorOwner owner = owners[index];
                Element element = owner.Element;
                int totalConnectors = owner.Connectors.Count;
                int unconnected = owner.Connectors.Count(IsUnconnected);

                if (unconnected <= 0)
                {
                    result.OkCount++;
                }
                else
                {
                    string category = ModelParameterExtractionService.GetElementCategoryName(element);
                    string itemBase = ResolveItemBase(element, category);
                    string resultText = unconnected >= totalConnectors ? ReviewError : ReviewPartialError;

                    if (string.Equals(resultText, ReviewError, StringComparison.Ordinal))
                    {
                        result.FullErrorCount++;
                    }
                    else
                    {
                        result.PartialErrorCount++;
                    }

                    result.ErrorCount++;
                    result.Rows.Add(new ReviewRow
                    {
                        File = safeFileLabel,
                        ItemBase = itemBase,
                        Id = GetElementIdValue(element).ToString(CultureInfo.InvariantCulture),
                        Name = ModelParameterExtractionService.GetElementTypeName(doc, element),
                        Result = resultText,
                        Content = BuildContent(category, totalConnectors, unconnected),
                        Etc = string.Empty,
                        Category = category,
                        Family = ResolveFamilyName(doc, element),
                        ConnectorCount = totalConnectors,
                        UnconnectedCount = unconnected,
                        ExtraParameterValues = ReadExtraParameterValues(doc, element, extraParameterNames)
                    });
                }

                if (index == 0 || index == owners.Count - 1 || index % 120 == 0)
                {
                    progress?.Invoke(5d + (((double)(index + 1) / totalOwners) * 90d), $"\uCEE4\uB125\uD130 \uBBF8\uC5F0\uACB0 \uAC80\uD1A0 \uC911 ({index + 1}/{owners.Count})");
                }
            }

            ApplyGroupItemTexts(result.Rows);

            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                TargetElementCount = result.TargetElementCount,
                ConnectorCount = result.ConnectorCount,
                ErrorCount = result.ErrorCount,
                FullErrorCount = result.FullErrorCount,
                PartialErrorCount = result.PartialErrorCount,
                OkCount = result.OkCount,
                Status = "success",
                Reason = BuildSummaryReason(result, hasCommonScopeFilter)
            });

            progress?.Invoke(100d, "\uCEE4\uB125\uD130 \uBBF8\uC5F0\uACB0 \uAC80\uD1A0 \uC644\uB8CC");
            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows, IEnumerable<string> extraParameterNames = null)
        {
            List<ReviewRow> exportRows = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(ShouldExportRow)
                .OrderBy(row => row?.Item ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(row => row?.Category ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(row => GetExportSortId(row))
                .ToList();
            List<string> extraHeaders = ResolveExportExtraHeaders(exportRows, extraParameterNames);

            var table = new DataTable("UnconnectedConnectorReview");
            table.Columns.Add("\uD56D\uBAA9");
            table.Columns.Add("ID");
            table.Columns.Add("Name");
            table.Columns.Add("\uACB0\uACFC");
            table.Columns.Add("\uB0B4\uC6A9");
            table.Columns.Add("\uBE44\uACE0");
            table.Columns.Add("Category");
            table.Columns.Add("Family");
            foreach (string name in extraHeaders)
            {
                if (!IsBaseExportColumn(name) && !table.Columns.Contains(name))
                {
                    table.Columns.Add(name);
                }
            }

            foreach (ReviewRow row in exportRows)
            {
                DataRow dataRow = table.NewRow();
                dataRow["\uD56D\uBAA9"] = row.Item ?? string.Empty;
                dataRow["ID"] = FormatExcelId(row.Id);
                dataRow["Name"] = row.Name ?? string.Empty;
                dataRow["\uACB0\uACFC"] = row.Result ?? string.Empty;
                dataRow["\uB0B4\uC6A9"] = row.Content ?? string.Empty;
                dataRow["\uBE44\uACE0"] = string.Empty;
                dataRow["Category"] = row.Category ?? string.Empty;
                dataRow["Family"] = row.Family ?? string.Empty;
                foreach (string name in extraHeaders)
                {
                    if (IsBaseExportColumn(name)) continue;
                    string value = string.Empty;
                    if (row.ExtraParameterValues != null)
                    {
                        row.ExtraParameterValues.TryGetValue(name, out value);
                    }
                    dataRow[name] = value ?? string.Empty;
                }
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static bool IsBaseExportColumn(string name)
        {
            string value = (name ?? string.Empty).Trim();
            return string.Equals(value, "\uD56D\uBAA9", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "ID", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "Name", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "\uACB0\uACFC", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "\uB0B4\uC6A9", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "\uBE44\uACE0", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "Category", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "Family", StringComparison.OrdinalIgnoreCase);
        }

        private static int GetExportSortId(ReviewRow row)
        {
            string value = (row?.Id ?? string.Empty).Trim().TrimEnd(',');
            int parsed;
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                ? parsed
                : int.MaxValue;
        }

        public static string BuildEmptyExportMessage(FileSummary summary)
        {
            if (summary != null && summary.CenterAxisEnabled)
            {
                if (summary.CenterAxisTargetCount <= 0)
                {
                    return "\uAC80\uD1A0 \uB300\uC0C1 \uC694\uC18C\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
                }

                if (summary.CenterAxisErrorCount <= 0)
                {
                    return "\uC911\uC2EC\uCD95 \uC624\uB958\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
                }
            }

            if (summary != null && summary.TapDepthEnabled)
            {
                if (summary.TapDepthTargetCount <= 0)
                {
                    return "\uAC80\uD1A0 \uB300\uC0C1 \uC694\uC18C\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
                }

                if (summary.TapDepthErrorCount <= 0)
                {
                    return "Tap/Saddle \uBB3B\uD798 \uC624\uB958\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
                }
            }

            if (summary != null && summary.TargetElementCount <= 0)
            {
                return "\uAC80\uD1A0 \uB300\uC0C1 \uC694\uC18C\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            }

            return "\uBBF8\uC5F0\uACB0 \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
        }

        public static int CountIssueRows(IEnumerable<ReviewRow> rows)
        {
            return (rows ?? Enumerable.Empty<ReviewRow>()).Count(row => IsIssueRow(row));
        }

        private static List<ConnectorOwner> CollectConnectorOwners(Document doc)
        {
            var owners = new Dictionary<int, ConnectorOwner>();

            foreach (FamilyInstance instance in SafeCollect(doc, typeof(FamilyInstance)).OfType<FamilyInstance>())
            {
                AddConnectorOwner(owners, instance);
            }

            foreach (Element curve in SafeCollect(doc, typeof(MEPCurve)))
            {
                AddConnectorOwner(owners, curve);
            }

            Type fabricationPartType = typeof(Element).Assembly.GetType("Autodesk.Revit.DB.FabricationPart", false);
            if (fabricationPartType != null)
            {
                foreach (Element part in SafeCollect(doc, fabricationPartType))
                {
                    AddConnectorOwner(owners, part);
                }
            }

            return owners.Values.ToList();
        }

        private static IEnumerable<Element> SafeCollect(Document doc, Type elementType)
        {
            if (doc == null || elementType == null) yield break;

            IList<Element> elements;
            try
            {
                elements = new FilteredElementCollector(doc)
                    .OfClass(elementType)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
            }
            catch
            {
                yield break;
            }

            foreach (Element element in elements)
            {
                if (element != null) yield return element;
            }
        }

        private static void AddConnectorOwner(IDictionary<int, ConnectorOwner> owners, Element element)
        {
            if (owners == null || element == null) return;

            int elementId = GetElementIdValue(element);
            if (elementId <= 0 || owners.ContainsKey(elementId)) return;

            List<Connector> connectors = GetPhysicalConnectors(element);
            if (connectors.Count == 0) return;

            owners[elementId] = new ConnectorOwner
            {
                Element = element,
                Connectors = connectors
            };
        }

        private static List<Connector> GetPhysicalConnectors(Element element)
        {
            var result = new List<Connector>();
            ConnectorManager manager = GetConnectorManager(element);
            if (manager == null) return result;

            ConnectorSet connectorSet;
            try
            {
                connectorSet = manager.Connectors;
            }
            catch
            {
                return result;
            }

            if (connectorSet == null) return result;

            foreach (Connector connector in connectorSet)
            {
                if (connector == null) continue;
                if (IsLogicalConnector(connector)) continue;
                result.Add(connector);
            }

            return result;
        }

        private static ConnectorManager GetConnectorManager(Element element)
        {
            if (element == null) return null;

            try
            {
                var mepCurve = element as MEPCurve;
                if (mepCurve?.ConnectorManager != null) return mepCurve.ConnectorManager;
            }
            catch
            {
            }

            try
            {
                var familyInstance = element as FamilyInstance;
                if (familyInstance?.MEPModel?.ConnectorManager != null) return familyInstance.MEPModel.ConnectorManager;
            }
            catch
            {
            }

            try
            {
                PropertyInfo property = element.GetType().GetProperty("ConnectorManager", BindingFlags.Instance | BindingFlags.Public);
                return property?.GetValue(element, null) as ConnectorManager;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsLogicalConnector(Connector connector)
        {
            try
            {
                return connector.ConnectorType == ConnectorType.Logical;
            }
            catch
            {
                return true;
            }
        }

        private static bool IsUnconnected(Connector connector)
        {
            try
            {
                return !connector.IsConnected;
            }
            catch
            {
                return false;
            }
        }

        private static bool ShouldExcludeElement(Element element)
        {
            if (element == null) return true;
            if (element.ViewSpecific) return true;
            if (element is View) return true;
            if (element is MEPSystem) return true;

            Category category = null;
            try { category = element.Category; } catch { }
            if (category == null) return true;
            if (category.CategoryType != CategoryType.Model) return true;

            string categoryName = string.Empty;
            try { categoryName = category.Name ?? string.Empty; } catch { }
            if (categoryName.IndexOf("Runs", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (categoryName.IndexOf("Analytical", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (categoryName.IndexOf("Placeholder", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            return MatchesCategory(category,
                "OST_PipeInsulations",
                "OST_DuctInsulations",
                "OST_DuctLinings",
                "OST_PlaceHolderPipes",
                "OST_PlaceHolderDucts",
                "OST_PlaceHolderCableTray",
                "OST_PlaceHolderConduits",
                "OST_MEPAnalyticalAirLoop",
                "OST_MEPAnalyticalWaterLoop",
                "OST_AnalyticalNodes",
                "OST_AnalyticalLinks",
                "OST_AnalyticalPipeNodes",
                "OST_AnalyticalPipeConnections",
                "OST_AnalyticalSpaces",
                "OST_Levels",
                "OST_Grids",
                "OST_GridChains",
                "OST_Rooms",
                "OST_Areas",
                "OST_MEPSpaces",
                "OST_Cameras",
                "OST_SectionBox",
                "OST_VolumeOfInterest",
                "OST_RvtLinks",
                "OST_PointClouds");
        }

        public static void ApplyGroupItemTexts(IList<ReviewRow> rows)
        {
            if (rows == null) return;

            Dictionary<string, int> counts = rows
                .Where(row => row != null && IsIssueRow(row) && !string.IsNullOrWhiteSpace(row.ItemBase))
                .GroupBy(row => BuildGroupKey(row), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

            foreach (ReviewRow row in rows.Where(row => row != null))
            {
                string itemBase = string.IsNullOrWhiteSpace(row.ItemBase) ? "Connector" : row.ItemBase;
                int count;
                if (!counts.TryGetValue(BuildGroupKey(row), out count)) count = 0;
                row.Item = BuildItemText(itemBase, count, row.IssueKind);
            }
        }

        private static string ResolveItemBase(Element element, string categoryName)
        {
            Category category = SafeCategory(element);
            if (MatchesCategory(category, "OST_PipeCurves", "OST_FlexPipeCurves")) return "PIPE";
            if (MatchesCategory(category, "OST_PipeFitting", "OST_PipeAccessory")) return "PIPE Fitting, Accessory";
            if (MatchesCategory(category, "OST_DuctCurves", "OST_FlexDuctCurves")) return "DUCT";
            if (MatchesCategory(category, "OST_DuctFitting", "OST_DuctAccessory")) return "Duct Fitting, Accessory";
            if (MatchesCategory(category, "OST_CableTray")) return "Cable Tray";
            if (MatchesCategory(category, "OST_CableTrayFitting")) return "Cable Tray Fittings";
            if (MatchesCategory(category, "OST_Conduit")) return "Conduit";
            if (MatchesCategory(category, "OST_ConduitFitting")) return "Conduit Fittings";

            string fallback = (categoryName ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(fallback) ? "Connector" : fallback;
        }

        private static string ResolveFamilyName(Document doc, Element element)
        {
            Category category = SafeCategory(element);
            if (MatchesCategory(category, "OST_PipeCurves", "OST_FlexPipeCurves")) return "Pipe Types";
            if (MatchesCategory(category, "OST_DuctCurves", "OST_FlexDuctCurves")) return "Duct Type";
            if (MatchesCategory(category, "OST_CableTray")) return "Cable Tray With Fittings";
            if (MatchesCategory(category, "OST_Conduit")) return "Conduit With Fittings";

            return ModelParameterExtractionService.GetElementFamilyName(doc, element);
        }

        private static List<string> NormalizeExtraParameterNames(IEnumerable<string> parameterNames)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (parameterNames == null) return result;

            foreach (string raw in parameterNames)
            {
                string name = (raw ?? string.Empty).Trim();
                if (name.Length == 0) continue;
                if (seen.Add(name)) result.Add(name);
            }

            return result;
        }

        private static List<string> ResolveExportExtraHeaders(IEnumerable<ReviewRow> rows, IEnumerable<string> extraParameterNames)
        {
            var result = NormalizeExtraParameterNames(extraParameterNames);
            var seen = new HashSet<string>(result, StringComparer.OrdinalIgnoreCase);

            foreach (ReviewRow row in rows ?? Enumerable.Empty<ReviewRow>())
            {
                if (row == null || row.ExtraParameterValues == null) continue;
                foreach (string key in row.ExtraParameterValues.Keys)
                {
                    string name = (key ?? string.Empty).Trim();
                    if (name.Length == 0) continue;
                    if (seen.Add(name)) result.Add(name);
                }
            }

            return result;
        }

        private static Dictionary<string, string> ReadExtraParameterValues(Document doc, Element element, IEnumerable<string> parameterNames)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in NormalizeExtraParameterNames(parameterNames))
            {
                result[name] = ResolveExtraParameterValue(doc, element, name);
            }
            return result;
        }

        private static string ResolveExtraParameterValue(Document doc, Element element, string parameterName)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName)) return string.Empty;
            string token = NormalizeSyntheticParameterName(parameterName);
            if (token == "category") return ModelParameterExtractionService.GetElementCategoryName(element);
            if (token == "family" || token == "familyname") return ResolveFamilyName(doc, element);
            if (token == "type" || token == "typename" || token == "name") return ModelParameterExtractionService.GetElementTypeName(doc, element);
            if (token == "id" || token == "elementid") return GetElementIdValue(element).ToString(CultureInfo.InvariantCulture);

            return ModelParameterExtractionService.GetElementParameterValue(doc, element, parameterName);
        }

        private static string NormalizeSyntheticParameterName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var chars = value
                .Trim()
                .Where(ch => char.IsLetterOrDigit(ch))
                .Select(ch => char.ToLowerInvariant(ch))
                .ToArray();
            return new string(chars);
        }

        private static string BuildItemText(string itemBase, int errorCount, string issueKind)
        {
            if (IsCenterAxisKind(issueKind))
            {
                return $"{(string.IsNullOrWhiteSpace(itemBase) ? "\uC911\uC2EC\uCD95 \uC5F0\uACB0" : itemBase)} Check({Math.Max(errorCount, 0).ToString(CultureInfo.InvariantCulture)}\uAC74)";
            }

            if (IsTapDepthKind(issueKind))
            {
                return $"{(string.IsNullOrWhiteSpace(itemBase) ? "Tap, Saddle \uBAA8\uB378\uB9C1 \uAC80\uD1A0 (\uBB3B\uD798)" : itemBase)} Check({Math.Max(errorCount, 0).ToString(CultureInfo.InvariantCulture)}\uAC74)";
            }

            return $"{(string.IsNullOrWhiteSpace(itemBase) ? "Connector" : itemBase)} \uBBF8\uC5F0\uACB0 Check({Math.Max(errorCount, 0).ToString(CultureInfo.InvariantCulture)}\uAC74)";
        }

        private static bool ShouldExportRow(ReviewRow row)
        {
            return row != null && (row.IsInformational || IsIssueRow(row));
        }

        private static bool IsIssueRow(ReviewRow row)
        {
            if (row == null || row.IsInformational) return false;
            return row.UnconnectedCount > 0 || IsCenterAxisKind(row.IssueKind) || IsTapDepthKind(row.IssueKind);
        }

        private static bool IsCenterAxisKind(string issueKind)
        {
            return string.Equals(issueKind, "centeraxis", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTapDepthKind(string issueKind)
        {
            return string.Equals(issueKind, "tapdepth", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildGroupKey(ReviewRow row)
        {
            if (row == null) return string.Empty;
            string itemBase = string.IsNullOrWhiteSpace(row.ItemBase) ? "Connector" : row.ItemBase.Trim();
            string issueKind = string.IsNullOrWhiteSpace(row.IssueKind) ? "unconnected" : row.IssueKind.Trim();
            return issueKind + "\u001F" + itemBase;
        }

        private static string BuildContent(string categoryName, int totalConnectors, int unconnectedConnectors)
        {
            string category = string.IsNullOrWhiteSpace(categoryName) ? "Connector" : categoryName.Trim();
            return $"[Category]: [{category}] Connector {totalConnectors.ToString(CultureInfo.InvariantCulture)}ea \uC911 {unconnectedConnectors.ToString(CultureInfo.InvariantCulture)}ea\uAC00 \uBBF8\uC5F0\uACB0 \uC0C1\uD0DC\uC785\uB2C8\uB2E4. ({unconnectedConnectors.ToString(CultureInfo.InvariantCulture)} of {totalConnectors.ToString(CultureInfo.InvariantCulture)})";
        }

        private static string BuildSummaryReason(ReviewResult result, bool hasCommonScopeFilter)
        {
            if (result == null) return "\uAC80\uD1A0 \uACB0\uACFC\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            if (result.TargetElementCount <= 0) return "\uCEE4\uB125\uD130\uAC00 \uC788\uB294 \uAC80\uD1A0 \uB300\uC0C1 \uAC1D\uCCB4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.";
            if (result.ErrorCount <= 0) return "\uBAA8\uB4E0 \uCEE4\uB125\uD130\uAC00 \uC5F0\uACB0\uB41C \uC0C1\uD0DC\uC785\uB2C8\uB2E4.";

            string scopeText = hasCommonScopeFilter ? " / \uACF5\uD1B5 \uD544\uD130 \uC801\uC6A9" : string.Empty;
            return $"\uBBF8\uC5F0\uACB0 \uCEE4\uB125\uD130\uAC00 \uC788\uB294 \uAC1D\uCCB4 {result.ErrorCount.ToString(CultureInfo.InvariantCulture)}\uAC74\uC744 \uCC3E\uC558\uC2B5\uB2C8\uB2E4.{scopeText}";
        }

        private static string FormatExcelId(string id)
        {
            string value = (id ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return value.EndsWith(",", StringComparison.Ordinal) ? value : value + ",";
        }

        private static int GetElementIdValue(Element element)
        {
            try
            {
                return element?.Id == null ? 0 : element.Id.CompatIntegerValue();
            }
            catch
            {
                return 0;
            }
        }

        private static Category SafeCategory(Element element)
        {
            try
            {
                return element?.Category;
            }
            catch
            {
                return null;
            }
        }

        private static bool MatchesCategory(Category category, params string[] builtInCategoryNames)
        {
            if (category == null || builtInCategoryNames == null || builtInCategoryNames.Length == 0) return false;

            Category current = category;
            var visited = new HashSet<int>();
            while (current != null && current.Id != null)
            {
                int id;
                try
                {
                    id = current.Id.CompatIntegerValue();
                }
                catch
                {
                    break;
                }

                if (id == 0 || !visited.Add(id)) break;
                if (MatchesCategoryId(id, builtInCategoryNames)) return true;

                try
                {
                    current = current.Parent;
                }
                catch
                {
                    current = null;
                }
            }

            return false;
        }

        private static bool MatchesCategoryId(int categoryId, IEnumerable<string> builtInCategoryNames)
        {
            foreach (string name in builtInCategoryNames ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;

                try
                {
                    BuiltInCategory builtInCategory;
                    if (Enum.TryParse(name, out builtInCategory) && categoryId == (int)builtInCategory)
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }
    }
}
