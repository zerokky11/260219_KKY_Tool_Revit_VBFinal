using System;
using System.Globalization;
using System.IO;
using System.Text;

public static class FamilyBrowserUniqueJsonReportStore
{
    public static string Save(string outputDirectory, string fileNameStem, object report)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new ArgumentException("An output directory is required.", "outputDirectory");
        }
        if (report == null)
        {
            throw new ArgumentNullException("report");
        }

        Directory.CreateDirectory(outputDirectory);
        string safeStem = string.IsNullOrWhiteSpace(fileNameStem) ? "family-browser-report" : fileNameStem.Trim();
        string fileName = safeStem + "-" +
            DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff", CultureInfo.InvariantCulture) + "-" +
            Guid.NewGuid().ToString("N").Substring(0, 8) + ".json";
        string path = Path.Combine(outputDirectory, fileName);
        string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
        try
        {
            byte[] payload = new UTF8Encoding(false).GetBytes(PlainJsonReportWriter.Serialize(report));
            using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(payload, 0, payload.Length);
                stream.Flush(true);
            }
            FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
            return path;
        }
        finally
        {
            try
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
            catch
            {
            }
        }
    }
}
