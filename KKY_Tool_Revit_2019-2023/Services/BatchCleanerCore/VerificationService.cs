using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class VerificationService
    {
        public static string VerifyCleanedOutputs(Autodesk.Revit.UI.UIApplication uiapp, BatchPrepareSession session, BatchCleanSettings settings, Action<string> log)
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            return VerifyPaths(uiapp, session.CleanedOutputPaths, !string.IsNullOrWhiteSpace(session.OutputFolder) ? session.OutputFolder : (settings != null ? settings.OutputFolder : null), settings, log);
        }

        public static string VerifyPaths(Autodesk.Revit.UI.UIApplication uiapp, IEnumerable<string> targetPaths, string outputFolder, BatchCleanSettings settings, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            List<string> paths = (targetPaths ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (paths.Count == 0) throw new InvalidOperationException("검토할 파일이 없습니다.");

            if (string.IsNullOrWhiteSpace(outputFolder)) outputFolder = settings.OutputFolder;
            if (string.IsNullOrWhiteSpace(outputFolder)) outputFolder = Path.GetDirectoryName(paths[0]);
            Directory.CreateDirectory(outputFolder);

            string csvPath = Path.Combine(outputFolder, "CleanVerification_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
            var lines = new List<string>();
            lines.Add(string.Join(",", new[]
            {
                "파일명","결과파일경로",
                "대상3D뷰존재","대상3D뷰이름","일반뷰1개만남음","일반뷰개수","3D뷰개수","뷰템플릿없음","대상뷰템플릿미사용",
                "상세수준Fine","표시형식ShadingWithEdges","모델카테고리표시","주석카테고리숨김","해석모델카테고리숨김","가져온카테고리표시",
                "페이즈신축","페이즈필터ShowAll","시작뷰일치",
                "Lines숨김","Mass숨김","Parts숨김","Site숨김","EndCut대상개수","EndCut숨김개수","EndCut전체숨김",
                "필터설정사용","필터문서존재","대상뷰필터적용","대상뷰필터표시상태",
                "뷰파라미터기대개수","뷰파라미터일치개수",
                "그룹없음","그룹타입없음","어셈블리없음","어셈블리타입없음",
                "디자인옵션없음","외부참조없음",
                "RevitLinkType개수","ImportInstance개수","CADLinkType개수","ImageType개수","PointCloudType개수",
                "최종판정"
            }));

            foreach (string path in paths)
            {
                Document doc = null;
                try
                {
                    log?.Invoke("검토 파일 열기: " + path);
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                    doc = uiapp.Application.OpenDocumentFile(modelPath, new OpenOptions());

                    string targetName = string.IsNullOrWhiteSpace(settings.Target3DViewName) ? "KKY_CLEAN_3D" : settings.Target3DViewName.Trim();
                    var nonTemplateViews = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => v != null && !v.IsTemplate)
                        .ToList();
                    var all3dViews = nonTemplateViews.OfType<View3D>().ToList();
                    View3D target3D = all3dViews.FirstOrDefault(v => string.Equals(v.Name, targetName, StringComparison.OrdinalIgnoreCase));

                    int nonTemplateViewCount = nonTemplateViews.Count;
                    int all3dViewCount = all3dViews.Count;
                    int viewTemplateCount = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Count(v => v != null && v.IsTemplate);

                    bool hasTarget3D = target3D != null;
                    bool onlyTarget3DRemains = hasTarget3D && nonTemplateViewCount == 1 && all3dViewCount == 1;
                    bool viewTemplateCountZero = viewTemplateCount == 0;
                    bool targetViewNoTemplate = hasTarget3D && target3D.ViewTemplateId != null && target3D.ViewTemplateId.IntegerValue == ElementId.InvalidElementId.IntegerValue;

                    bool detailFine = hasTarget3D && target3D.DetailLevel == ViewDetailLevel.Fine;
                    bool displayShadedWithEdges = hasTarget3D && target3D.DisplayStyle == DisplayStyle.ShadingWithEdges;
                    bool modelCategoriesShown = hasTarget3D && !target3D.AreModelCategoriesHidden;
                    bool annotationHidden = hasTarget3D && target3D.AreAnnotationCategoriesHidden;
                    bool analyticalHidden = hasTarget3D && target3D.AreAnalyticalModelCategoriesHidden;
                    bool importShown = hasTarget3D && !target3D.AreImportCategoriesHidden;

                    bool phaseNewConstruction = false;
                    bool phaseFilterShowAll = false;
                    if (hasTarget3D)
                    {
                        phaseNewConstruction = MatchesPhase(target3D, BuiltInParameter.VIEW_PHASE, "New Construction", "신축");
                        phaseFilterShowAll = MatchesPhase(target3D, BuiltInParameter.VIEW_PHASE_FILTER, "Show All", "모두 표시");
                    }

                    bool startingViewMatch = false;
                    try
                    {
                        StartingViewSettings startingViewSettings = StartingViewSettings.GetStartingViewSettings(doc);
                        startingViewMatch = hasTarget3D && startingViewSettings != null && startingViewSettings.ViewId != null && startingViewSettings.ViewId.IntegerValue == target3D.Id.IntegerValue;
                    }
                    catch
                    {
                    }

                    bool linesHidden = hasTarget3D && IsTopLevelCategoryHidden(target3D, doc, BuiltInCategory.OST_Lines, "Lines", "선");
                    bool massHidden = hasTarget3D && IsTopLevelCategoryHidden(target3D, doc, BuiltInCategory.OST_Mass, "Mass", "매스");
                    bool partsHidden = hasTarget3D && IsTopLevelCategoryHidden(target3D, doc, BuiltInCategory.OST_Parts, "Parts", "파츠");
                    bool siteHidden = hasTarget3D && IsTopLevelCategoryHidden(target3D, doc, BuiltInCategory.OST_Site, "Site", "대지");

                    int endCutSubcategoryTotal = 0;
                    int endCutSubcategoryHidden = 0;
                    if (hasTarget3D)
                    {
                        foreach (Category category in doc.Settings.Categories)
                        {
                            if (category == null || !IsTargetFittingCategory(category)) continue;
                            foreach (Category subCategory in EnumerateAllSubCategories(category))
                            {
                                if (subCategory == null) continue;
                                if (!ShouldHideNamedSubCategory(category, subCategory)) continue;
                                endCutSubcategoryTotal++;
                                if (GetCategoryHiddenSafe(target3D, subCategory.Id)) endCutSubcategoryHidden++;
                            }
                        }
                    }
                    bool endCutSubcategoryAllHidden = endCutSubcategoryTotal == 0 || endCutSubcategoryTotal == endCutSubcategoryHidden;

                    bool filterConfigured = settings.UseFilter && settings.FilterProfile != null && settings.FilterProfile.IsConfigured();
                    bool filterExists = false;
                    bool filterApplied = false;
                    bool filterVisible = false;
                    if (hasTarget3D && filterConfigured)
                    {
                        Element filterElement = new FilteredElementCollector(doc)
                            .OfClass(typeof(ParameterFilterElement))
                            .FirstOrDefault(x => string.Equals(x.Name, settings.FilterProfile.FilterName, StringComparison.OrdinalIgnoreCase));
                        if (filterElement != null)
                        {
                            filterExists = true;
                            ICollection<ElementId> filters = target3D.GetFilters();
                            filterApplied = filters.Any(x => x != null && x.IntegerValue == filterElement.Id.IntegerValue);
                            if (filterApplied)
                            {
                                filterVisible = GetFilterVisibilitySafe(target3D, filterElement.Id, true);
                            }
                        }
                    }

                    int expectedViewParameterCount = 0;
                    int matchedViewParameterCount = 0;
                    if (hasTarget3D && settings.ViewParameters != null)
                    {
                        foreach (ViewParameterAssignment assignment in settings.ViewParameters.Where(x => x != null && x.Enabled && !string.IsNullOrWhiteSpace(x.ParameterName)))
                        {
                            expectedViewParameterCount++;
                            Parameter parameter = target3D.LookupParameter(assignment.ParameterName);
                            if (parameter != null && ParameterMatchesExpected(parameter, assignment.ParameterValue))
                            {
                                matchedViewParameterCount++;
                            }
                        }
                    }

                    int groupCount = new FilteredElementCollector(doc).OfClass(typeof(Group)).GetElementCount();
                    int groupTypeCount = new FilteredElementCollector(doc).OfClass(typeof(GroupType)).GetElementCount();
                    int assemblyCount = new FilteredElementCollector(doc).OfClass(typeof(AssemblyInstance)).GetElementCount();
                    int assemblyTypeCount = GetElementCountByTypeName(doc, "Autodesk.Revit.DB.AssemblyType");
                    int designOptionCount = new FilteredElementCollector(doc).OfClass(typeof(DesignOption)).GetElementCount();

                    int externalReferenceCount = 0;
                    try { externalReferenceCount = ExternalFileUtils.GetAllExternalFileReferences(doc).Count; } catch { }

                    int revitLinkTypeCount = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).GetElementCount();
                    int importInstanceCount = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).GetElementCount();
                    int cadLinkTypeCount = new FilteredElementCollector(doc).OfClass(typeof(CADLinkType)).GetElementCount();
                    int imageTypeCount = GetElementCountByTypeName(doc, "Autodesk.Revit.DB.ImageType");
                    int pointCloudTypeCount = GetElementCountByTypeName(doc, "Autodesk.Revit.DB.PointClouds.PointCloudType");

                    bool passed = hasTarget3D
                        && onlyTarget3DRemains
                        && viewTemplateCountZero
                        && targetViewNoTemplate
                        && detailFine
                        && displayShadedWithEdges
                        && modelCategoriesShown
                        && annotationHidden
                        && analyticalHidden
                        && importShown
                        && phaseNewConstruction
                        && phaseFilterShowAll
                        && startingViewMatch
                        && linesHidden
                        && massHidden
                        && partsHidden
                        && siteHidden
                        && endCutSubcategoryAllHidden
                        && (!filterConfigured || (filterExists && filterApplied))
                        && matchedViewParameterCount == expectedViewParameterCount
                        && groupCount == 0
                        && groupTypeCount == 0
                        && assemblyCount == 0
                        && assemblyTypeCount == 0
                        && designOptionCount == 0
                        && externalReferenceCount == 0
                        && revitLinkTypeCount == 0
                        && importInstanceCount == 0
                        && cadLinkTypeCount == 0
                        && imageTypeCount == 0
                        && pointCloudTypeCount == 0;

                    lines.Add(string.Join(",", new[]
                    {
                        Csv(Path.GetFileName(path)),
                        Csv(path),
                        Csv(ToYesNo(hasTarget3D)),
                        Csv(targetName),
                        Csv(ToYesNo(onlyTarget3DRemains)),
                        nonTemplateViewCount.ToString(),
                        all3dViewCount.ToString(),
                        Csv(ToYesNo(viewTemplateCountZero)),
                        Csv(ToYesNo(targetViewNoTemplate)),
                        Csv(ToYesNo(detailFine)),
                        Csv(ToYesNo(displayShadedWithEdges)),
                        Csv(ToYesNo(modelCategoriesShown)),
                        Csv(ToYesNo(annotationHidden)),
                        Csv(ToYesNo(analyticalHidden)),
                        Csv(ToYesNo(importShown)),
                        Csv(ToYesNo(phaseNewConstruction)),
                        Csv(ToYesNo(phaseFilterShowAll)),
                        Csv(ToYesNo(startingViewMatch)),
                        Csv(ToYesNo(linesHidden)),
                        Csv(ToYesNo(massHidden)),
                        Csv(ToYesNo(partsHidden)),
                        Csv(ToYesNo(siteHidden)),
                        endCutSubcategoryTotal.ToString(),
                        endCutSubcategoryHidden.ToString(),
                        Csv(ToYesNo(endCutSubcategoryAllHidden)),
                        Csv(ToYesNo(filterConfigured)),
                        Csv(ToYesNo(filterExists)),
                        Csv(ToYesNo(filterApplied)),
                        Csv(ToYesNo(filterVisible)),
                        expectedViewParameterCount.ToString(),
                        matchedViewParameterCount.ToString(),
                        Csv(ToYesNo(groupCount == 0)),
                        Csv(ToYesNo(groupTypeCount == 0)),
                        Csv(ToYesNo(assemblyCount == 0)),
                        Csv(ToYesNo(assemblyTypeCount == 0)),
                        Csv(ToYesNo(designOptionCount == 0)),
                        Csv(ToYesNo(externalReferenceCount == 0)),
                        revitLinkTypeCount.ToString(),
                        importInstanceCount.ToString(),
                        cadLinkTypeCount.ToString(),
                        imageTypeCount.ToString(),
                        pointCloudTypeCount.ToString(),
                        Csv(passed ? "통과" : "확인필요")
                    }));

                    log?.Invoke("검토 완료: " + Path.GetFileName(path) + " / " + (passed ? "PASS" : "CHECK"));
                }
                finally
                {
                    try
                    {
                        if (doc != null && doc.IsValidObject)
                        {
                            doc.Close(false);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            File.WriteAllText(csvPath, string.Join(Environment.NewLine, lines) + Environment.NewLine, new UTF8Encoding(true));
            log?.Invoke("검토 CSV 저장: " + csvPath);
            return csvPath;
        }

        private static bool MatchesPhase(View view, BuiltInParameter parameterId, string englishName, string koreanName)
        {
            if (view == null) return false;
            try
            {
                Parameter parameter = view.get_Parameter(parameterId);
                if (parameter == null || parameter.StorageType != StorageType.ElementId) return false;
                ElementId valueId = parameter.AsElementId();
                if (valueId == null || valueId == ElementId.InvalidElementId) return false;
                Element element = view.Document.GetElement(valueId);
                string name = element != null ? element.Name : string.Empty;
                return string.Equals(name, englishName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, koreanName, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsTopLevelCategoryHidden(View view, Document doc, BuiltInCategory builtInCategory, params string[] fallbackNames)
        {
            if (view == null || doc == null) return false;
            Category category = null;
            try
            {
                category = Category.GetCategory(doc, builtInCategory);
            }
            catch
            {
            }

            if (category == null)
            {
                foreach (Category item in doc.Settings.Categories)
                {
                    if (item == null) continue;
                    if (fallbackNames.Any(name => EqualsNormalizedCategoryName(item, name)))
                    {
                        category = item;
                        break;
                    }
                }
            }

            return category != null && GetCategoryHiddenSafe(view, category.Id);
        }

        private static bool GetCategoryHiddenSafe(View view, ElementId categoryId)
        {
            if (view == null || categoryId == null || categoryId == ElementId.InvalidElementId) return false;
            try
            {
                var method = typeof(View).GetMethod("GetCategoryHidden", new[] { typeof(ElementId) });
                if (method != null)
                {
                    object result = method.Invoke(view, new object[] { categoryId });
                    return result is bool && (bool)result;
                }
            }
            catch
            {
            }
            return false;
        }

        private static bool GetFilterVisibilitySafe(View view, ElementId filterId, bool defaultValue)
        {
            if (view == null || filterId == null || filterId == ElementId.InvalidElementId) return defaultValue;
            try
            {
                var enabledMethod = typeof(View).GetMethod("GetIsFilterEnabled", new[] { typeof(ElementId) });
                if (enabledMethod != null)
                {
                    object enabled = enabledMethod.Invoke(view, new object[] { filterId });
                    if (enabled is bool)
                    {
                        return (bool)enabled;
                    }
                }
            }
            catch
            {
            }

            try
            {
                return view.GetFilterVisibility(filterId);
            }
            catch
            {
                return defaultValue;
            }
        }

        private static bool ParameterMatchesExpected(Parameter parameter, string expectedValue)
        {
            if (parameter == null) return false;
            string actual = GetComparableParameterText(parameter);
            return string.Equals(actual ?? string.Empty, expectedValue ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetComparableParameterText(Parameter parameter)
        {
            if (parameter == null) return string.Empty;
            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        return parameter.AsString() ?? parameter.AsValueString() ?? string.Empty;
                    case StorageType.Integer:
                        return parameter.AsValueString() ?? parameter.AsInteger().ToString();
                    case StorageType.Double:
                        return parameter.AsValueString() ?? parameter.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        return id != null && id != ElementId.InvalidElementId ? id.IntegerValue.ToString() : string.Empty;
                    default:
                        return parameter.AsValueString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int GetElementCountByTypeName(Document doc, string fullTypeName)
        {
            try
            {
                Type type = typeof(Document).Assembly.GetType(fullTypeName, false);
                if (type == null) return 0;
                return new FilteredElementCollector(doc).OfClass(type).GetElementCount();
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsTargetFittingCategory(Category category)
        {
            if (category == null) return false;
            int id = category.Id.IntegerValue;
            return id == (int)BuiltInCategory.OST_CableTrayFitting
                   || id == (int)BuiltInCategory.OST_ConduitFitting
                   || id == (int)BuiltInCategory.OST_DuctFitting
                   || id == (int)BuiltInCategory.OST_PipeFitting
                   || ContainsNormalizedCategoryName(category, "cablefitting")
                   || ContainsNormalizedCategoryName(category, "conduitfitting")
                   || ContainsNormalizedCategoryName(category, "ductfitting")
                   || ContainsNormalizedCategoryName(category, "pipefitting");
        }

        private static bool ShouldHideNamedSubCategory(Category parentCategory, Category subCategory)
        {
            if (!IsTargetFittingCategory(parentCategory)) return false;
            string name = NormalizeCategoryName(subCategory != null ? subCategory.Name : string.Empty);
            return !string.IsNullOrWhiteSpace(name) && (name.Contains("end") || name.Contains("cut"));
        }

        private static IEnumerable<Category> EnumerateAllSubCategories(Category category)
        {
            var visited = new HashSet<int>();
            foreach (Category child in EnumerateAllSubCategoriesRecursive(category, visited))
            {
                yield return child;
            }
        }

        private static IEnumerable<Category> EnumerateAllSubCategoriesRecursive(Category category, HashSet<int> visited)
        {
            if (category == null || category.SubCategories == null) yield break;
            foreach (Category child in category.SubCategories)
            {
                if (child == null) continue;
                if (!visited.Add(child.Id.IntegerValue)) continue;
                yield return child;
                foreach (Category nested in EnumerateAllSubCategoriesRecursive(child, visited))
                {
                    yield return nested;
                }
            }
        }

        private static bool EqualsNormalizedCategoryName(Category category, string expected)
        {
            return string.Equals(NormalizeCategoryName(category != null ? category.Name : string.Empty), NormalizeCategoryName(expected), StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsNormalizedCategoryName(Category category, string token)
        {
            string name = NormalizeCategoryName(category != null ? category.Name : string.Empty);
            string work = NormalizeCategoryName(token);
            return !string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(work) && name.Contains(work);
        }

        private static string NormalizeCategoryName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return new string(value.Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_').ToArray()).ToLowerInvariant();
        }

        private static string ToYesNo(bool value)
        {
            return value ? "예" : "아니오";
        }

        private static string Csv(string value)
        {
            string text = value ?? string.Empty;
            bool quoted = text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n");
            if (text.Contains("\"")) text = text.Replace("\"", "\"\"");
            return quoted ? "\"" + text + "\"" : text;
        }
    }
}
