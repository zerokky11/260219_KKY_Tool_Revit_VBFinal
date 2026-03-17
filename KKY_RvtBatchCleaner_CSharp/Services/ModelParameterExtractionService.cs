using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class ModelParameterExtractionService
    {
        public static string ExportModelParameters(UIApplication uiapp, BatchPrepareSession session, string outputFolder, string parameterNamesCsv, Action<string> log)
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            return ExportModelParameters(uiapp, session.CleanedOutputPaths, string.IsNullOrWhiteSpace(outputFolder) ? session.OutputFolder : outputFolder, parameterNamesCsv, log);
        }

        public static string ExportModelParameters(UIApplication uiapp, IEnumerable<string> targetPaths, string outputFolder, string parameterNamesCsv, Action<string> log)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));

            List<string> paths = (targetPaths ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (paths.Count == 0) throw new InvalidOperationException("추출할 파일이 없습니다.");

            List<string> parameterNames = SplitParameterNames(parameterNamesCsv);
            if (parameterNames.Count == 0) throw new InvalidOperationException("추출할 파라미터 이름을 하나 이상 입력해야 합니다.");

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                outputFolder = Path.GetDirectoryName(paths[0]);
            }
            Directory.CreateDirectory(outputFolder);

            string csvPath = Path.Combine(outputFolder, "ModelParameterExport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
            var sb = new StringBuilder();
            var header = new List<string> { "파일명", "요소ID", "카테고리", "패밀리이름", "타입이름" };
            header.AddRange(parameterNames);
            sb.AppendLine(string.Join(",", header.Select(Csv)));

            foreach (string path in paths)
            {
                Document doc = null;
                try
                {
                    log?.Invoke("속성값 추출 파일 열기: " + path);
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                    doc = uiapp.Application.OpenDocumentFile(modelPath, new OpenOptions());

                    IList<Element> elements = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Cast<Element>()
                        .Where(x => x != null)
                        .Where(IsEligibleModelElement)
                        .OrderBy(x => x.Id.IntegerValue)
                        .ToList();

                    int rowCount = 0;
                    string fileName = Path.GetFileName(path);
                    foreach (Element element in elements)
                    {
                        var row = new List<string>();
                        row.Add(fileName);
                        row.Add(element.Id.IntegerValue.ToString());
                        row.Add(GetCategoryName(element));
                        row.Add(GetFamilyName(doc, element));
                        row.Add(GetTypeName(doc, element));
                        foreach (string parameterName in parameterNames)
                        {
                            row.Add(GetParameterValue(doc, element, parameterName));
                        }
                        sb.AppendLine(string.Join(",", row.Select(Csv)));
                        rowCount++;
                    }

                    log?.Invoke("속성값 추출 완료: " + fileName + " / 행 " + rowCount);
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

            File.WriteAllText(csvPath, sb.ToString(), new UTF8Encoding(true));
            log?.Invoke("속성값 추출 CSV 저장: " + csvPath);
            return csvPath;
        }

        private static bool IsEligibleModelElement(Element element)
        {
            if (element == null) return false;
            if (element.ViewSpecific) return false;
            if (element.Category == null) return false;
            if (element.Category.CategoryType != CategoryType.Model) return false;
            if (element is View) return false;
            if (element is CurveElement) return false;
            if (element is ReferencePlane) return false;
            if (element is Group) return false;
            if (element is AssemblyInstance) return false;
            if (element is Level) return false;
            if (element is Grid) return false;
            if (element is Room) return false;
            if (element is Area) return false;

            try
            {
                int categoryId = element.Category.Id != null ? element.Category.Id.IntegerValue : 0;
                if (categoryId == (int)BuiltInCategory.OST_Lines) return false;
                if (categoryId == (int)BuiltInCategory.OST_CLines) return false;
                if (categoryId == (int)BuiltInCategory.OST_IOSModelGroups) return false;
                if (categoryId == (int)BuiltInCategory.OST_Assemblies) return false;
                if (categoryId == (int)BuiltInCategory.OST_Levels) return false;
                if (categoryId == (int)BuiltInCategory.OST_Grids) return false;
                if (categoryId == (int)BuiltInCategory.OST_Rooms) return false;
                if (categoryId == (int)BuiltInCategory.OST_Areas) return false;
            }
            catch
            {
            }

            string categoryName = string.Empty;
            try { categoryName = element.Category.Name ?? string.Empty; } catch { }
            if (string.Equals(categoryName, "Lines", StringComparison.OrdinalIgnoreCase)) return false;
            if (string.Equals(categoryName, "Reference Planes", StringComparison.OrdinalIgnoreCase)) return false;
            if (string.Equals(categoryName, "참조 평면", StringComparison.OrdinalIgnoreCase)) return false;
            return true;
        }

        private static List<string> SplitParameterNames(string csv)
        {
            if (string.IsNullOrWhiteSpace(csv)) return new List<string>();
            return csv.Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string GetCategoryName(Element element)
        {
            try { return element.Category != null ? element.Category.Name : string.Empty; } catch { return string.Empty; }
        }

        private static string GetFamilyName(Document doc, Element element)
        {
            if (doc == null || element == null) return string.Empty;
            try
            {
                FamilyInstance familyInstance = element as FamilyInstance;
                if (familyInstance != null && familyInstance.Symbol != null && familyInstance.Symbol.Family != null)
                {
                    return familyInstance.Symbol.Family.Name ?? string.Empty;
                }
            }
            catch
            {
            }

            try
            {
                ElementType type = GetElementType(doc, element);
                if (type != null)
                {
                    string familyName = type.FamilyName;
                    if (!string.IsNullOrWhiteSpace(familyName)) return familyName;
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static string GetTypeName(Document doc, Element element)
        {
            if (doc == null || element == null) return string.Empty;
            try
            {
                ElementType type = GetElementType(doc, element);
                if (type != null)
                {
                    return type.Name ?? string.Empty;
                }
            }
            catch
            {
            }
            return string.Empty;
        }

        private static ElementType GetElementType(Document doc, Element element)
        {
            if (doc == null || element == null) return null;
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return null;
                return doc.GetElement(typeId) as ElementType;
            }
            catch
            {
                return null;
            }
        }

        private static string GetParameterValue(Document doc, Element element, string parameterName)
        {
            Parameter parameter = FindParameterOnElementOrType(doc, element, parameterName);
            if (parameter == null) return string.Empty;

            try
            {
                string valueString = parameter.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString)) return valueString;
            }
            catch
            {
            }

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        return parameter.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return parameter.AsInteger().ToString();
                    case StorageType.Double:
                        return parameter.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        ElementId id = parameter.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId) return string.Empty;
                        Element refElement = doc.GetElement(id);
                        return refElement != null ? (refElement.Name ?? id.IntegerValue.ToString()) : id.IntegerValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Parameter FindParameterOnElementOrType(Document doc, Element element, string parameterName)
        {
            if (doc == null || element == null || string.IsNullOrWhiteSpace(parameterName)) return null;
            Parameter parameter = FindParameterByName(element, parameterName);
            if (parameter != null) return parameter;

            ElementType type = GetElementType(doc, element);
            if (type != null)
            {
                parameter = FindParameterByName(type, parameterName);
                if (parameter != null) return parameter;
            }

            return null;
        }

        private static Parameter FindParameterByName(Element owner, string parameterName)
        {
            if (owner == null || string.IsNullOrWhiteSpace(parameterName)) return null;
            try
            {
                Parameter direct = owner.LookupParameter(parameterName);
                if (IsUsableNamedParameter(direct, parameterName)) return direct;
            }
            catch
            {
            }

            try
            {
                foreach (Parameter parameter in owner.Parameters)
                {
                    if (IsUsableNamedParameter(parameter, parameterName)) return parameter;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool IsUsableNamedParameter(Parameter parameter, string expectedName)
        {
            if (parameter == null || parameter.Definition == null) return false;
            string actualName = parameter.Definition.Name;
            return !string.IsNullOrWhiteSpace(actualName)
                   && string.Equals(actualName.Trim(), expectedName.Trim(), StringComparison.OrdinalIgnoreCase);
        }
        private static string Csv(string value)
        {
            string text = value ?? string.Empty;
            bool mustQuote = text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n");
            if (text.Contains("\""))
            {
                text = text.Replace("\"", "\"\"");
            }

            return mustQuote ? "\"" + text + "\"" : text;
        }
    }
}
