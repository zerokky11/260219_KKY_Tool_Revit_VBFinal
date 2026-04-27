using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    public static class ProjectParameterDuplicationReviewService
    {
        public const string ItemLabel = "Parameter duplication check";
        public const string ScopeAll = "all";
        public const string ScopeSelected = "selected";

        public sealed class Settings
        {
            public string Scope { get; set; } = ScopeAll;
            public List<string> ParameterNames { get; set; } = new List<string>();
        }

        public sealed class ReviewRow
        {
            public string File { get; set; } = string.Empty;
            public string Item { get; set; } = ItemLabel;
            public string Id { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Result { get; set; } = string.Empty;
            public string Content { get; set; } = string.Empty;
            public string Etc { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Family { get; set; } = string.Empty;
            public string Solutions { get; set; } = string.Empty;
        }

        public sealed class FileSummary
        {
            public string File { get; set; } = string.Empty;
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public string Scope { get; set; } = ScopeAll;
            public int SelectedNameCount { get; set; }
            public string Status { get; set; } = "pending";
            public string Reason { get; set; } = string.Empty;
        }

        public sealed class ReviewResult
        {
            public string File { get; set; } = string.Empty;
            public string Scope { get; set; } = ScopeAll;
            public int SelectedNameCount { get; set; }
            public int TotalReviewed { get; set; }
            public int ErrorCount { get; set; }
            public int OkCount { get; set; }
            public List<ReviewRow> Rows { get; set; } = new List<ReviewRow>();
            public List<FileSummary> FileSummaries { get; set; } = new List<FileSummary>();
        }

        private sealed class ProjectParameterInfo
        {
            public ElementId Id { get; set; } = ElementId.InvalidElementId;
            public string Name { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
        }

        private sealed class ParameterElementLookup
        {
            public Dictionary<Guid, ElementId> SharedByGuid { get; } = new Dictionary<Guid, ElementId>();
            public Dictionary<string, List<ElementId>> ByName { get; } = new Dictionary<string, List<ElementId>>(StringComparer.OrdinalIgnoreCase);
        }

        public static ReviewResult RunOnDocument(Document doc, string fileLabel, Settings settings, Action<double, string> progress = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            string safeFileLabel = string.IsNullOrWhiteSpace(fileLabel) ? (doc.Title ?? string.Empty) : fileLabel.Trim();
            string scope = NormalizeScope(settings.Scope);
            HashSet<string> selectedNames = BuildSelectedNameSet(settings.ParameterNames);

            progress?.Invoke(5d, "Collecting project parameters");
            List<ProjectParameterInfo> allParameters = CollectProjectParameters(doc, progress);

            List<ProjectParameterInfo> reviewedParameters = allParameters
                .Where(info => ShouldReview(info, scope, selectedNames))
                .OrderBy(info => info.Name ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(info => info.Id != null ? info.Id.IntegerValue : int.MaxValue)
                .ToList();

            var duplicateCounts = reviewedParameters
                .GroupBy(info => NormalizeParameterKey(info.Name), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

            var result = new ReviewResult
            {
                File = safeFileLabel,
                Scope = scope,
                SelectedNameCount = selectedNames.Count,
                TotalReviewed = reviewedParameters.Count
            };

            for (int index = 0; index < reviewedParameters.Count; index++)
            {
                ProjectParameterInfo info = reviewedParameters[index];
                string key = NormalizeParameterKey(info.Name);
                int duplicateCount = duplicateCounts.TryGetValue(key, out int count) ? count : 0;
                bool isDuplicate = duplicateCount > 1;

                if (isDuplicate)
                {
                    result.ErrorCount++;
                }
                else
                {
                    result.OkCount++;
                }

                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = ItemLabel,
                    Id = ToParameterIdText(info.Id),
                    Name = isDuplicate ? (info.Name ?? string.Empty) : "Parameter",
                    Result = isDuplicate ? "Error" : "OK",
                    Content = isDuplicate
                        ? $"[Parameter]: The parameter [{info.Name}] is duplicated."
                        : $"[Parameter]: The parameter [{info.Name}] was created successfully.",
                    Etc = isDuplicate ? $"Duplicate count: {duplicateCount.ToString(CultureInfo.InvariantCulture)}" : string.Empty,
                    Category = info.Category ?? string.Empty,
                    Family = string.Empty,
                    Solutions = isDuplicate ? "Remove or rename the duplicated project parameter and keep only one binding." : string.Empty
                });

                if (index == 0 || index == reviewedParameters.Count - 1 || index % 20 == 0)
                {
                    progress?.Invoke(20d + (((double)(index + 1) / Math.Max(reviewedParameters.Count, 1)) * 80d), $"Reviewing project parameters ({index + 1}/{reviewedParameters.Count})");
                }
            }

            if (result.TotalReviewed == 0)
            {
                result.Rows.Add(new ReviewRow
                {
                    File = safeFileLabel,
                    Item = ItemLabel,
                    Id = string.Empty,
                    Name = "Parameter",
                    Result = "OK",
                    Content = scope == ScopeSelected
                        ? "No matching project parameters were found for the selected review scope."
                        : "No reviewable project parameters were found.",
                    Etc = string.Empty,
                    Category = string.Empty,
                    Family = string.Empty,
                    Solutions = string.Empty
                });
            }

            progress?.Invoke(100d, "Project parameter duplication review complete");

            result.FileSummaries.Add(new FileSummary
            {
                File = safeFileLabel,
                TotalReviewed = result.TotalReviewed,
                ErrorCount = result.ErrorCount,
                OkCount = result.OkCount,
                Scope = scope,
                SelectedNameCount = selectedNames.Count,
                Status = "success",
                Reason = BuildSummaryReason(result.TotalReviewed, result.ErrorCount, scope, selectedNames.Count)
            });

            return result;
        }

        public static DataTable BuildExportTable(IEnumerable<ReviewRow> rows)
        {
            var table = new DataTable("ProjectParameterDuplicationReview");
            table.Columns.Add("Item");
            table.Columns.Add("ID");
            table.Columns.Add("Name");
            table.Columns.Add("Result");
            table.Columns.Add("Content");
            table.Columns.Add("Etc");
            table.Columns.Add("Solutions");

            List<ReviewRow> source = (rows ?? Enumerable.Empty<ReviewRow>())
                .Where(row => row != null)
                .ToList();

            foreach (ReviewRow row in source)
            {
                DataRow dataRow = table.NewRow();
                dataRow["Item"] = row.Item ?? ItemLabel;
                dataRow["ID"] = row.Id ?? string.Empty;
                dataRow["Name"] = row.Name ?? string.Empty;
                dataRow["Result"] = row.Result ?? string.Empty;
                dataRow["Content"] = row.Content ?? string.Empty;
                dataRow["Etc"] = row.Etc ?? string.Empty;
                dataRow["Solutions"] = row.Solutions ?? string.Empty;
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static string BuildSummaryReason(int totalReviewed, int errorCount, string scope, int selectedNameCount)
        {
            if (totalReviewed <= 0)
            {
                return scope == ScopeSelected
                    ? $"No matching project parameters were found in the selected {selectedNameCount.ToString(CultureInfo.InvariantCulture)} target names."
                    : "No reviewable project parameters were found.";
            }

            if (errorCount <= 0)
            {
                return $"All {totalReviewed.ToString(CultureInfo.InvariantCulture)} reviewed project parameters are unique.";
            }

            return $"{errorCount.ToString(CultureInfo.InvariantCulture)} reviewed project parameters are duplicated.";
        }

        private static bool ShouldReview(ProjectParameterInfo info, string scope, HashSet<string> selectedNames)
        {
            if (info == null || string.IsNullOrWhiteSpace(info.Name))
            {
                return false;
            }

            if (!string.Equals(scope, ScopeSelected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return selectedNames.Contains(NormalizeParameterKey(info.Name));
        }

        private static List<ProjectParameterInfo> CollectProjectParameters(Document doc, Action<double, string> progress)
        {
            var result = new List<ProjectParameterInfo>();
            if (doc == null) return result;

            ParameterElementLookup lookup = BuildParameterElementLookup(doc);
            BindingMap map = doc.ParameterBindings;
            if (map == null) return result;

            DefinitionBindingMapIterator iterator = map.ForwardIterator();
            iterator.Reset();

            int index = 0;
            while (iterator.MoveNext())
            {
                index++;

                Definition definition = iterator.Key as Definition;
                if (definition == null) continue;

                Binding binding = iterator.Current as Binding;
                if (!(binding is ElementBinding elementBinding)) continue;

                string name = (definition.Name ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(name)) continue;

                result.Add(new ProjectParameterInfo
                {
                    Id = ResolveParameterId(doc, definition, lookup),
                    Name = name,
                    Category = BuildCategoryText(elementBinding)
                });

                if (index == 1 || index % 25 == 0)
                {
                    progress?.Invoke(5d + Math.Min(index, 100) * 0.1d, $"Scanning project parameters ({index})");
                }
            }

            return result;
        }

        private static ParameterElementLookup BuildParameterElementLookup(Document doc)
        {
            var lookup = new ParameterElementLookup();
            if (doc == null) return lookup;

            IEnumerable<ParameterElement> elements;
            try
            {
                elements = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterElement))
                    .Cast<ParameterElement>()
                    .ToList();
            }
            catch
            {
                return lookup;
            }

            foreach (ParameterElement parameterElement in elements)
            {
                if (parameterElement == null) continue;

                if (parameterElement is SharedParameterElement sharedParameterElement)
                {
                    try
                    {
                        Guid guid = sharedParameterElement.GuidValue;
                        if (guid != Guid.Empty && !lookup.SharedByGuid.ContainsKey(guid))
                        {
                            lookup.SharedByGuid.Add(guid, sharedParameterElement.Id);
                        }
                    }
                    catch
                    {
                    }
                }

                string name = string.Empty;
                try
                {
                    Definition definition = parameterElement.GetDefinition();
                    name = (definition != null ? definition.Name : string.Empty) ?? string.Empty;
                }
                catch
                {
                    name = string.Empty;
                }

                name = name.Trim();
                if (string.IsNullOrWhiteSpace(name)) continue;

                string key = NormalizeParameterKey(name);
                if (!lookup.ByName.TryGetValue(key, out List<ElementId> ids))
                {
                    ids = new List<ElementId>();
                    lookup.ByName[key] = ids;
                }

                ids.Add(parameterElement.Id);
            }

            return lookup;
        }

        private static ElementId ResolveParameterId(Document doc, Definition definition, ParameterElementLookup lookup)
        {
            if (definition == null) return ElementId.InvalidElementId;

            ExternalDefinition externalDefinition = definition as ExternalDefinition;
            if (externalDefinition != null)
            {
                try
                {
                    SharedParameterElement sharedParameterElement = SharedParameterElement.Lookup(doc, externalDefinition.GUID);
                    if (sharedParameterElement != null)
                    {
                        return sharedParameterElement.Id;
                    }
                }
                catch
                {
                }

                if (lookup != null && lookup.SharedByGuid.TryGetValue(externalDefinition.GUID, out ElementId sharedId))
                {
                    return sharedId ?? ElementId.InvalidElementId;
                }
            }

            ElementId internalId = TryGetInternalDefinitionId(definition);
            if (internalId != null && internalId != ElementId.InvalidElementId)
            {
                return internalId;
            }

            string key = NormalizeParameterKey(definition.Name);
            if (lookup != null && lookup.ByName.TryGetValue(key, out List<ElementId> ids) && ids.Count > 0)
            {
                return ids[0] ?? ElementId.InvalidElementId;
            }

            return ElementId.InvalidElementId;
        }

        private static ElementId TryGetInternalDefinitionId(Definition definition)
        {
            InternalDefinition internalDefinition = definition as InternalDefinition;
            if (internalDefinition == null) return ElementId.InvalidElementId;

            try
            {
                PropertyInfo property = typeof(InternalDefinition).GetProperty("Id", BindingFlags.Instance | BindingFlags.Public);
                if (property == null) return ElementId.InvalidElementId;
                return property.GetValue(internalDefinition, null) as ElementId ?? ElementId.InvalidElementId;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        private static string BuildCategoryText(ElementBinding binding)
        {
            if (binding == null) return string.Empty;

            var names = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                CategorySet categories = binding.Categories;
                if (categories == null) return string.Empty;

                foreach (Category category in categories)
                {
                    string name = (category?.Name ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return names.Count == 0 ? string.Empty : string.Join(", ", names);
        }

        private static HashSet<string> BuildSelectedNameSet(IEnumerable<string> names)
        {
            return new HashSet<string>(
                (names ?? Enumerable.Empty<string>())
                    .Select(name => NormalizeParameterKey(name))
                    .Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);
        }

        private static string NormalizeScope(string value)
        {
            string normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
            return normalized == ScopeSelected ? ScopeSelected : ScopeAll;
        }

        private static string NormalizeParameterKey(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static string ToParameterIdText(ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId) return string.Empty;
            return id.IntegerValue.ToString(CultureInfo.InvariantCulture);
        }
    }
}
