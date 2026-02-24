using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KKY_Tool_Revit.Services
{
    public sealed class SharedParameterStatus
    {
        public string Path { get; set; } = string.Empty;
        public bool IsSet { get; set; }
        public bool ExistsOnDisk { get; set; }
        public bool CanOpen { get; set; }
        public string Status { get; set; } = "warn";
        public string StatusLabel { get; set; } = "설정 필요";
        public string WarningMessage { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }

    public sealed class SharedParameterDefinitionItem
    {
        public string Name { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;
        public string GroupName { get; set; } = string.Empty;
        public string DataTypeToken { get; set; } = string.Empty;
    }

    public static class SharedParameterStatusService
    {
        public static SharedParameterStatus GetStatus(UIApplication app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var status = new SharedParameterStatus();
            var path = app.Application.SharedParametersFilename;

            status.Path = path ?? string.Empty;
            status.IsSet = !string.IsNullOrWhiteSpace(path);
            status.ExistsOnDisk = status.IsSet && File.Exists(path);

            if (!status.IsSet)
            {
                status.Status = "warn";
                status.StatusLabel = "설정 필요";
                status.WarningMessage = "Shared Parameter 파일 경로가 설정되지 않았습니다.";
                return status;
            }

            if (!status.ExistsOnDisk)
            {
                status.Status = "error";
                status.StatusLabel = "파일 없음";
                status.ErrorMessage = "Shared Parameter 파일을 찾을 수 없습니다.";
                return status;
            }

            DefinitionFile defFile;
            try
            {
                defFile = app.Application.OpenSharedParameterFile();
            }
            catch (Exception ex)
            {
                status.Status = "error";
                status.StatusLabel = "열기 실패";
                status.ErrorMessage = ex.Message;
                return status;
            }

            status.CanOpen = defFile != null;
            if (!status.CanOpen)
            {
                status.Status = "error";
                status.StatusLabel = "열기 실패";
                status.ErrorMessage = "Shared Parameter 파일을 열 수 없습니다.";
                return status;
            }

            status.Status = "ok";
            status.StatusLabel = "정상";
            return status;
        }

        public static List<SharedParameterDefinitionItem> ListDefinitions(UIApplication app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var items = new List<SharedParameterDefinitionItem>();
            DefinitionFile defFile;

            try
            {
                defFile = app.Application.OpenSharedParameterFile();
            }
            catch
            {
                return items;
            }

            if (defFile == null)
            {
                return items;
            }

            foreach (DefinitionGroup grp in defFile.Groups)
            {
                if (grp == null) continue;

                foreach (Definition defn in grp.Definitions)
                {
                    if (defn == null) continue;

                    var guidValue = string.Empty;
                    if (defn is ExternalDefinition ext)
                    {
                        try
                        {
                            guidValue = ext.GUID.ToString("D");
                        }
                        catch
                        {
                            guidValue = string.Empty;
                        }
                    }

                    items.Add(new SharedParameterDefinitionItem
                    {
                        Name = defn.Name,
                        Guid = guidValue,
                        GroupName = grp.Name,
                        DataTypeToken = TryGetDefinitionDataTypeToken(defn)
                    });
                }
            }

            return items;
        }

        private static string TryGetDefinitionDataTypeToken(Definition defn)
        {
            if (defn == null) return string.Empty;

            try
            {
                var m = defn.GetType().GetMethod("GetDataType", Type.EmptyTypes);
                if (m != null)
                {
                    var dtObj = m.Invoke(defn, null);
                    if (dtObj != null)
                    {
                        var pTypeId = dtObj.GetType().GetProperty("TypeId");
                        if (pTypeId != null)
                        {
                            var v = pTypeId.GetValue(dtObj, null);
                            if (v is string s && !string.IsNullOrWhiteSpace(s))
                            {
                                return s;
                            }
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                var p = defn.GetType().GetProperty("ParameterType");
                if (p != null)
                {
                    var v = p.GetValue(defn, null);
                    if (v != null) return v.ToString();
                }
            }
            catch
            {
                // ignore
            }

            return string.Empty;
        }
    }
}
