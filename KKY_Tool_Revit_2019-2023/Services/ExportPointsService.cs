using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Infrastructure;

namespace KKY_Tool_Revit.Services
{
    public class ExportPointsService
    {
        public class Row
        {
            public string File { get; set; }
            public double ProjectE { get; set; }
            public double ProjectN { get; set; }
            public double ProjectZ { get; set; }
            public double SurveyE { get; set; }
            public double SurveyN { get; set; }
            public double SurveyZ { get; set; }
            public double TrueNorth { get; set; }
        }

        public class ProgressInfo
        {
            public string Phase { get; set; }
            public string Message { get; set; }
            public int Current { get; set; }
            public int Total { get; set; }
            public double PhaseProgress { get; set; }
        }

        public static IList<Row> Run(UIApplication uiapp, object files, Action<ProgressInfo> progress = null)
        {
            var app = uiapp.Application;
            var list = new List<Row>();

            var paths = new List<string>();
            if (files is IEnumerable<object> objs)
            {
                foreach (var o in objs)
                {
                    var s = o as string;
                    if (!string.IsNullOrWhiteSpace(s) && File.Exists(s)) paths.Add(s);
                }
            }
            else if (files is string one && File.Exists(one))
            {
                paths.Add(one);
            }

            paths = paths.Distinct().ToList();
            var total = paths.Count;
            ReportProgress(progress, "COLLECT", "파일 목록 준비 중", 0, total, 0.0);

            if (paths.Count == 0)
            {
                ReportProgress(progress, "DONE", "대상 파일이 없습니다.", 0, 0, 1.0);
                return list;
            }

            for (var i = 0; i < paths.Count; i++)
            {
                var p = paths[i];
                var stageProgress = total > 0 ? (double)i / total : 0.0;
                ReportProgress(progress, "EXTRACT", $"파일 열기: {Path.GetFileName(p)}", i, total, stageProgress);

                Document doc = null;
                try
                {
                    var opt = BuildOpenOptions(p);
                    var mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(p);
                    doc = app.OpenDocumentFile(mp, opt);

                    var row = new Row { File = Path.GetFileName(p) };
                    Extract(doc, row);
                    list.Add(row);
                    var afterProgress = total > 0 ? (double)(i + 1) / total : 1.0;
                    ReportProgress(progress, "EXTRACT", $"포인트 추출: {Path.GetFileName(p)}", i + 1, total, afterProgress);
                }
                catch
                {
                    var afterProgress = total > 0 ? (double)(i + 1) / total : 1.0;
                    ReportProgress(progress, "EXTRACT", $"오류로 건너뜀: {Path.GetFileName(p)}", i + 1, total, afterProgress);
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.Close(false); } catch { }
                    }
                }
            }

            ReportProgress(progress, "DONE", "포인트 추출 완료", total, total, 1.0);
            return list;
        }

        public static IList<Row> RunOnDocument(Document doc, string fileName, Action<ProgressInfo> progress = null)
        {
            var list = new List<Row>();
            if (doc == null) return list;
            var row = new Row { File = string.IsNullOrWhiteSpace(fileName) ? doc.Title : fileName };
            try
            {
                ReportProgress(progress, "EXTRACT", $"포인트 추출: {row.File}", 0, 1, 0.0);
                Extract(doc, row);
                list.Add(row);
                ReportProgress(progress, "DONE", "포인트 추출 완료", 1, 1, 1.0);
            }
            catch
            {
                ReportProgress(progress, "ERROR", "포인트 추출 실패", 0, 1, 1.0);
            }

            return list;
        }

        public static string ExportToExcel(UIApplication uiapp, object files, string unit = "ft", bool doAutoFit = false)
        {
            var rows = Run(uiapp, files);

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var outPath = Path.Combine(desktop, $"ExportPoints_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            var normalizedUnit = NormalizeUnit(unit);
            var headers = BuildHeaders(normalizedUnit);
            var data = rows.Select(r => new object[]
            {
                r.File,
                RoundCoord(ToUnitValue(r.ProjectE, normalizedUnit)),
                RoundCoord(ToUnitValue(r.ProjectN, normalizedUnit)),
                RoundCoord(ToUnitValue(r.ProjectZ, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyE, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyN, normalizedUnit)),
                RoundCoord(ToUnitValue(r.SurveyZ, normalizedUnit)),
                Math.Round(r.TrueNorth, 3)
            });

            var dt = BuildTable(headers, data);
            ExcelCore.EnsureNoDataRow(dt, "추출 결과가 없습니다.");
            ExcelCore.SaveXlsx(outPath, "Points", dt, doAutoFit, sheetKey: "Points", exportKind: "points");
            ExcelExportStyleRegistry.ApplyStylesForKey("points", outPath, autoFit: doAutoFit, excelMode: doAutoFit ? "normal" : "fast");

            return outPath;
        }

        private static DataTable BuildTable(IEnumerable<string> headers, IEnumerable<IEnumerable<object>> rows)
        {
            var dt = new DataTable("ExportedPoints");

            var headArr = (headers ?? Enumerable.Empty<string>())
                .Select((h, i) => string.IsNullOrWhiteSpace(h) ? $"Col{i + 1}" : h.Trim())
                .ToArray();

            foreach (var h in headArr)
            {
                dt.Columns.Add(h);
            }

            if (rows != null)
            {
                foreach (var r in rows)
                {
                    var vals = (r ?? Enumerable.Empty<object>()).ToArray();
                    var dr = dt.NewRow();
                    for (var i = 0; i <= Math.Min(vals.Length, dt.Columns.Count) - 1; i++)
                    {
                        dr[i] = (vals[i] ?? string.Empty).ToString();
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        private static void Extract(Document doc, Row row)
        {
            var basePt = new FilteredElementCollector(doc)
                .OfClass(typeof(BasePoint))
                .Cast<BasePoint>()
                .FirstOrDefault(bp => bp.IsShared == false);

            var surveyPt = new FilteredElementCollector(doc)
                .OfClass(typeof(BasePoint))
                .Cast<BasePoint>()
                .FirstOrDefault(bp => bp.IsShared == true);

            var project = basePt != null ? basePt.Position : XYZ.Zero;
            var survey = surveyPt != null ? surveyPt.Position : XYZ.Zero;

            row.ProjectE = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_EASTWEST_PARAM, project.X);
            row.ProjectN = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM, project.Y);
            row.ProjectZ = TryGetParamDouble(basePt, BuiltInParameter.BASEPOINT_ELEVATION_PARAM, project.Z);

            row.SurveyE = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_EASTWEST_PARAM, survey.X);
            row.SurveyN = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM, survey.Y);
            row.SurveyZ = TryGetParamDouble(surveyPt, BuiltInParameter.BASEPOINT_ELEVATION_PARAM, survey.Z);

            row.TrueNorth = GetTrueNorthDegrees(doc, basePt);
        }

        private static double GetTrueNorthDegrees(Document doc, BasePoint basePt)
        {
            double deg;
            if (TryGetBasePointAngle(basePt, out deg)) return NormalizeAngleDegrees(deg);
            if (TryGetProjectLocationAngle(doc, out deg)) return NormalizeAngleDegrees(deg);
            return 0.0;
        }

        private static bool TryGetBasePointAngle(BasePoint basePt, out double deg)
        {
            deg = 0.0;
            if (basePt == null) return false;

            try
            {
                var p = basePt.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM);
                if (p != null)
                {
                    deg = p.AsDouble() * (180.0 / Math.PI);
                    return true;
                }
            }
            catch { }

            try
            {
                var p = basePt.LookupParameter("Angle to True North");
                if (p != null)
                {
                    deg = p.AsDouble() * (180.0 / Math.PI);
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static bool TryGetProjectLocationAngle(Document doc, out double deg)
        {
            deg = 0.0;
            try
            {
                var pl = doc.ActiveProjectLocation;
                if (pl == null) return false;
                var pp = pl.GetProjectPosition(XYZ.Zero);
                if (pp != null)
                {
                    deg = pp.Angle * (180.0 / Math.PI);
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static double NormalizeAngleDegrees(double deg)
        {
            if (double.IsNaN(deg) || double.IsInfinity(deg)) return 0.0;
            var v = deg % 360.0;
            if (v < 0.0) v += 360.0;
            return v;
        }

        private static double TryGetParamDouble(Element el, BuiltInParameter bip, double fallback)
        {
            if (el == null) return fallback;
            try
            {
                var p = el.get_Parameter(bip);
                if (p != null)
                {
                    return p.AsDouble();
                }
            }
            catch { }
            return fallback;
        }

        private static OpenOptions BuildOpenOptions(string path)
        {
            BasicFileInfo info = null;
            try
            {
                info = BasicFileInfo.Extract(path);
            }
            catch { }

            var ws = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            var opt = new OpenOptions
            {
                Audit = false,
                DetachFromCentralOption = (info != null && info.IsCentral)
                    ? DetachFromCentralOption.DetachAndPreserveWorksets
                    : DetachFromCentralOption.DoNotDetach
            };
            opt.SetOpenWorksetsConfiguration(ws);
            return opt;
        }

        private static string NormalizeUnit(string unit)
        {
            var u = (unit ?? string.Empty).Trim().ToLowerInvariant();
            if (u == "m" || u == "meter" || u == "meters") return "m";
            if (u == "mm" || u == "millimeter" || u == "millimeters") return "mm";
            return "ft";
        }

        private static string[] BuildHeaders(string unit)
        {
            var suffix = "(ft)";
            if (unit == "m") suffix = "(m)";
            else if (unit == "mm") suffix = "(mm)";

            return new[]
            {
                "File",
                $"ProjectPoint_E{suffix}", $"ProjectPoint_N{suffix}", $"ProjectPoint_Z{suffix}",
                $"SurveyPoint_E{suffix}", $"SurveyPoint_N{suffix}", $"SurveyPoint_Z{suffix}",
                "TrueNorthAngle(deg)"
            };
        }

        private static double ToUnitValue(double valueFt, string unit)
        {
            if (unit == "m") return valueFt * 0.3048;
            if (unit == "mm") return valueFt * 304.8;
            return valueFt;
        }

        private static double RoundCoord(double v)
        {
            return Math.Round(v, 4);
        }

        private static void ReportProgress(Action<ProgressInfo> cb, string phase, string message, int current, int total, double phaseProgress)
        {
            if (cb == null) return;
            try
            {
                cb(new ProgressInfo
                {
                    Phase = phase,
                    Message = message,
                    Current = current,
                    Total = total,
                    PhaseProgress = Clamp01(phaseProgress)
                });
            }
            catch { }
        }

        private static double Clamp01(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return 0.0;
            if (v < 0.0) return 0.0;
            if (v > 1.0) return 1.0;
            return v;
        }
    }
}
