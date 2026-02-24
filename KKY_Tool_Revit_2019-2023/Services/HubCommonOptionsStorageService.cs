using System;
using System.IO;
using System.Web.Script.Serialization;

namespace KKY_Tool_Revit.Services
{
    public static class HubCommonOptionsStorageService
    {
        public sealed class HubCommonOptions
        {
            public string ExtraParamsText { get; set; } = string.Empty;
            public string TargetFilterText { get; set; } = string.Empty;
            public bool ExcludeEndDummy { get; set; }
        }

        public static HubCommonOptions Load()
        {
            var result = new HubCommonOptions();
            try
            {
                var path = GetOptionsPath();
                if (!File.Exists(path)) return result;

                var json = File.ReadAllText(path);
                if (string.IsNullOrWhiteSpace(json)) return result;

                var serializer = new JavaScriptSerializer();
                var loaded = serializer.Deserialize<HubCommonOptions>(json);
                if (loaded != null)
                {
                    result = loaded;
                }
            }
            catch
            {
                return result;
            }

            return result;
        }

        public static bool Save(HubCommonOptions options)
        {
            if (options == null) return false;
            try
            {
                var path = GetOptionsPath();
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var serializer = new JavaScriptSerializer();
                var json = serializer.Serialize(options);
                File.WriteAllText(path, json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetOptionsPath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            if (string.IsNullOrWhiteSpace(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }

            var dir = Path.Combine(root, "KKY_Tool_Revit");
            return Path.Combine(dir, "hub-common-options.json");
        }
    }
}
