using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public static class GuidAuditService
    {
        public sealed class RunResult
        {
            public int Mode { get; set; }
            public DataTable Project { get; set; }
            public DataTable FamilyDetail { get; set; }
            public DataTable FamilyIndex { get; set; }
            public string RunId { get; set; }
        }

        public static RunResult Run(
            UIApplication uiApp,
            int mode,
            IEnumerable<string> rvtPaths,
            Action<double, string> progress,
            bool includeAnnotation)
        {
            var map = SharedParamReader.ReadSharedParamNameGuidMap(uiApp.Application);
            if (map.Count == 0)
            {
                throw new InvalidOperationException("Shared Parameter 파일을 읽을 수 없습니다.");
            }

            var doc = uiApp.ActiveUIDocument?.Document ?? throw new InvalidOperationException("활성 문서가 없습니다.");

            var project = CreateProjectTable();
            var familyDetail = CreateFamilyDetailTable();
            var familyIndex = CreateFamilyIndexTable();

            progress?.Invoke(5, "프로젝트 파라미터 검사 시작");
            AuditProject(doc, map, project);

            if (mode == 2)
            {
                progress?.Invoke(35, "패밀리 검사 시작");
                AuditFamily(doc, map, familyDetail, familyIndex, includeAnnotation);
            }

            progress?.Invoke(100, "GUID 검사 완료");
            return new RunResult
            {
                Mode = mode,
                Project = project,
                FamilyDetail = familyDetail,
                FamilyIndex = familyIndex,
                RunId = Guid.NewGuid().ToString("N")
            };
        }

        private static void AuditProject(Document doc, Dictionary<string, List<Guid>> map, DataTable table)
        {
            var bindingMap = doc.ParameterBindings;
            var it = bindingMap.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                if (!(it.Key is Definition def)) continue;
                var name = def.Name ?? string.Empty;

                var isShared = def is ExternalDefinition ext;
                var projectGuid = isShared ? ext.GUID.ToString() : string.Empty;
                var fileGuid = map.TryGetValue(name, out var guids) ? guids.FirstOrDefault().ToString() : string.Empty;

                string result;
                if (isShared && string.IsNullOrWhiteSpace(fileGuid))
                    result = "MissingInSharedParamFile";
                else if (isShared && !string.Equals(projectGuid, fileGuid, StringComparison.OrdinalIgnoreCase))
                    result = "GuidMismatch";
                else if (!isShared && map.ContainsKey(name))
                    result = "ProjectParamNameExistsInShared";
                else
                    continue;

                table.Rows.Add(
                    doc.PathName,
                    doc.Title,
                    "Project",
                    name,
                    isShared ? "Shared" : "Project",
                    projectGuid,
                    fileGuid,
                    result,
                    string.Empty);
            }
        }

        private static void AuditFamily(Document doc, Dictionary<string, List<Guid>> map, DataTable detail, DataTable index, bool includeAnnotation)
        {
            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .OrderBy(f => f.Name)
                .ToList();

            foreach (var fam in families)
            {
                if (!includeAnnotation && fam.FamilyCategory != null && fam.FamilyCategory.CategoryType == CategoryType.Annotation)
                    continue;

                index.Rows.Add(doc.PathName, doc.Title, fam.Name);
                Document famDoc = null;
                try
                {
                    famDoc = doc.EditFamily(fam);
                    var fm = famDoc.FamilyManager;
                    foreach (FamilyParameter fp in fm.Parameters)
                    {
                        var name = fp.Definition?.Name ?? string.Empty;
                        var isShared = fp.IsShared;
                        var pg = isShared ? fp.GUID.ToString() : string.Empty;
                        var fg = map.TryGetValue(name, out var guids) ? guids.FirstOrDefault().ToString() : string.Empty;

                        var result = string.Empty;
                        var kind = isShared ? "Shared" : (fp.Definition != null ? "Family" : "BuiltIn");
                        if (isShared && string.IsNullOrWhiteSpace(fg)) result = "MissingInSharedParamFile";
                        else if (isShared && !string.Equals(pg, fg, StringComparison.OrdinalIgnoreCase)) result = "GuidMismatch";
                        else if (!isShared && map.ContainsKey(name)) result = "FamilyParamNameExistsInShared";
                        else continue;

                        detail.Rows.Add(doc.PathName, doc.Title, fam.Name, fam.FamilyCategory?.Name ?? string.Empty, name, kind, isShared, pg, fg, result, string.Empty);
                    }
                }
                catch (Exception ex)
                {
                    detail.Rows.Add(doc.PathName, doc.Title, fam.Name, fam.FamilyCategory?.Name ?? string.Empty, "", "", false, "", "", "Error", ex.Message);
                }
                finally
                {
                    famDoc?.Close(false);
                }
            }
        }

        public static DataTable CreateProjectTable()
        {
            var dt = new DataTable("Project");
            dt.Columns.Add("RvtPath");
            dt.Columns.Add("RvtName");
            dt.Columns.Add("Scope");
            dt.Columns.Add("ParamName");
            dt.Columns.Add("ParamKind");
            dt.Columns.Add("ProjectGuid");
            dt.Columns.Add("FileGuid");
            dt.Columns.Add("Result");
            dt.Columns.Add("Notes");
            return dt;
        }

        public static DataTable CreateFamilyDetailTable()
        {
            var dt = new DataTable("FamilyDetail");
            dt.Columns.Add("RvtPath");
            dt.Columns.Add("RvtName");
            dt.Columns.Add("FamilyName");
            dt.Columns.Add("Category");
            dt.Columns.Add("ParamName");
            dt.Columns.Add("ParamKind");
            dt.Columns.Add("IsShared", typeof(bool));
            dt.Columns.Add("ProjectGuid");
            dt.Columns.Add("FileGuid");
            dt.Columns.Add("Result");
            dt.Columns.Add("Notes");
            return dt;
        }

        public static DataTable CreateFamilyIndexTable()
        {
            var dt = new DataTable("FamilyIndex");
            dt.Columns.Add("RvtPath");
            dt.Columns.Add("RvtName");
            dt.Columns.Add("FamilyName");
            return dt;
        }
    }
}
