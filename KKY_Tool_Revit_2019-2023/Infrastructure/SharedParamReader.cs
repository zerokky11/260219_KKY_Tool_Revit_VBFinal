using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.ApplicationServices;

namespace KKY_Tool_Revit.Infrastructure
{
    public sealed class SharedParameterStatus
    {
        public string Status { get; set; } = "warn";
        public string Path { get; set; } = string.Empty;
        public bool ExistsOnDisk { get; set; }
        public bool CanOpen { get; set; }
        public bool IsSet { get; set; }
        public string Warning { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }

    public static class SharedParamReader
    {
        public static SharedParameterStatus ReadStatus(Application app)
        {
            var status = new SharedParameterStatus();
            try
            {
                var path = app?.SharedParametersFilename ?? string.Empty;
                status.Path = path;
                status.IsSet = !string.IsNullOrWhiteSpace(path);
                status.ExistsOnDisk = status.IsSet && File.Exists(path);

                if (!status.IsSet)
                {
                    status.Status = "warn";
                    status.Warning = "Shared Parameter 파일 경로가 설정되지 않았습니다.";
                    return status;
                }

                if (!status.ExistsOnDisk)
                {
                    status.Status = "error";
                    status.ErrorMessage = "Shared Parameter 파일이 디스크에 존재하지 않습니다.";
                    return status;
                }

                using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    status.CanOpen = true;
                }

                status.Status = "ok";
                return status;
            }
            catch (Exception ex)
            {
                status.Status = "error";
                status.ErrorMessage = ex.Message;
                return status;
            }
        }

        public static Dictionary<string, List<Guid>> ReadSharedParamNameGuidMap(Application app)
        {
            var result = new Dictionary<string, List<Guid>>(StringComparer.OrdinalIgnoreCase);
            var status = ReadStatus(app);
            if (!string.Equals(status.Status, "ok", StringComparison.OrdinalIgnoreCase))
            {
                return result;
            }

            foreach (var line in File.ReadLines(status.Path))
            {
                if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("PARAM\t", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var cells = line.Split('\t');
                if (cells.Length < 3 || !Guid.TryParse(cells[1], out var guid))
                {
                    continue;
                }

                var name = (cells[2] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!result.TryGetValue(name, out var list))
                {
                    list = new List<Guid>();
                    result[name] = list;
                }

                if (!list.Contains(guid))
                {
                    list.Add(guid);
                }
            }

            return result;
        }
    }
}
