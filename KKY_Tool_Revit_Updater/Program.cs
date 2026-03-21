using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;

namespace KKY_Tool_Revit_Updater
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var options = UpdaterOptions.Parse(args);
            using (var logger = new UpdateLogger(options.LogPath))
            {
                try
                {
                    logger.Log("Updater started.");
                    logger.Log("Mode: " + options.Mode);
                    logger.Log("Source: " + options.SourcePath);
                    logger.Log("Wait PID: " + options.WaitPid);
                    logger.Log("Is admin: " + IsAdministrator());

                    if (string.IsNullOrWhiteSpace(options.SourcePath) || !File.Exists(options.SourcePath))
                    {
                        logger.Log("Source file not found.");
                        return 2;
                    }

                    if (!IsAdministrator())
                    {
                        logger.Log("Updater is not running with administrator privileges.");
                        return 5;
                    }

                    WaitForProcessExit(options.WaitPid, logger);
                    WaitForAllRevitExit(logger);

                    if (string.Equals(options.Mode, "installer", StringComparison.OrdinalIgnoreCase))
                    {
                        LaunchInstaller(options.SourcePath, logger);
                    }
                    else
                    {
                        ApplyZipPackage(options.SourcePath, logger);
                        ShowSuccessMessage();
                    }

                    logger.Log("Updater completed successfully.");
                    return 0;
                }
                catch (Exception ex)
                {
                    logger.Log("Updater failed: " + ex);
                    ShowFailureMessage(options.LogPath, ex);
                    return 1;
                }
            }
        }

        private static void ShowSuccessMessage()
        {
            MessageBox.Show(
                "업데이트가 완료되었습니다." + Environment.NewLine +
                "이제 Revit을 다시 실행해 주세요.",
                "KKY Tool Revit 업데이트",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
        }

        private static void ShowFailureMessage(string logPath, Exception ex)
        {
            var message =
                "업데이트 적용에 실패했습니다." + Environment.NewLine +
                "자세한 내용은 로그를 확인해 주세요.";

            if (!string.IsNullOrWhiteSpace(logPath))
            {
                message += Environment.NewLine + Environment.NewLine + "로그 위치:" + Environment.NewLine + logPath;
            }

            if (ex != null && !string.IsNullOrWhiteSpace(ex.Message))
            {
                message += Environment.NewLine + Environment.NewLine + "오류:" + Environment.NewLine + ex.Message;
            }

            MessageBox.Show(
                message,
                "KKY Tool Revit 업데이트",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1);
        }

        private static void WaitForProcessExit(int waitPid, UpdateLogger logger)
        {
            if (waitPid <= 0)
            {
                return;
            }

            while (true)
            {
                try
                {
                    using (var process = Process.GetProcessById(waitPid))
                    {
                        logger.Log("Waiting for process exit: " + waitPid);
                        if (!process.HasExited)
                        {
                            Thread.Sleep(TimeSpan.FromSeconds(2));
                            continue;
                        }
                    }
                }
                catch (ArgumentException)
                {
                    break;
                }

                break;
            }
        }

        private static void WaitForAllRevitExit(UpdateLogger logger)
        {
            while (Process.GetProcessesByName("Revit").Any())
            {
                logger.Log("Waiting for all Revit processes to exit.");
                Thread.Sleep(TimeSpan.FromSeconds(2));
            }

            Thread.Sleep(800);
        }

        private static void LaunchInstaller(string installerPath, UpdateLogger logger)
        {
            logger.Log("Launching installer: " + installerPath);

            var psi = new ProcessStartInfo();
            var extension = Path.GetExtension(installerPath);

            if (string.Equals(extension, ".msi", StringComparison.OrdinalIgnoreCase))
            {
                psi.FileName = "msiexec.exe";
                psi.Arguments = "/i " + QuoteArgument(installerPath);
            }
            else
            {
                psi.FileName = installerPath;
            }

            psi.UseShellExecute = true;
            psi.WorkingDirectory = Path.GetDirectoryName(installerPath) ?? Environment.CurrentDirectory;
            Process.Start(psi);
        }

        private static void ApplyZipPackage(string zipPath, UpdateLogger logger)
        {
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var extractRoot = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "UpdaterExtract", stamp);
            var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var addinsRoot = Path.Combine(programData, "Autodesk", "Revit", "Addins");

            logger.Log("Extract root: " + extractRoot);

            if (Directory.Exists(extractRoot))
            {
                ForceDeleteDirectory(extractRoot, logger);
            }

            Directory.CreateDirectory(extractRoot);
            ZipFile.ExtractToDirectory(zipPath, extractRoot);
            logger.Log("ZIP extracted.");

            foreach (var year in new[] { "2019", "2021", "2023", "2025" })
            {
                if (!IsRevitInstalled(year))
                {
                    logger.Log("Skipping Revit " + year + " because it is not installed.");
                    continue;
                }

                var sourceYearDir = Path.Combine(extractRoot, year);
                if (!Directory.Exists(sourceYearDir))
                {
                    logger.Log("Skipping Revit " + year + " because package content is missing.");
                    continue;
                }

                var destinationYearDir = Path.Combine(addinsRoot, year);
                var sourceAddin = Path.Combine(sourceYearDir, "KKY_Tool_Revit.addin");
                var sourcePayload = Path.Combine(sourceYearDir, "KKY_Tool_Revit");
                var destinationPayload = Path.Combine(destinationYearDir, "KKY_Tool_Revit");

                logger.Log("Updating Revit " + year + " at " + destinationYearDir);
                Directory.CreateDirectory(destinationYearDir);

                if (Directory.Exists(destinationPayload))
                {
                    ForceDeleteDirectory(destinationPayload, logger);
                }

                if (Directory.Exists(sourcePayload))
                {
                    CopyDirectoryContents(sourcePayload, destinationPayload, logger);
                }

                if (File.Exists(sourceAddin))
                {
                    var destinationAddin = Path.Combine(destinationYearDir, "KKY_Tool_Revit.addin");
                    SafeCopyFile(sourceAddin, destinationAddin, logger);
                }
            }
        }

        private static bool IsRevitInstalled(string year)
        {
            var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var addinPath = Path.Combine(programData, "Autodesk", "Revit", "Addins", year);
            var pf64 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Autodesk", "Revit " + year);
            var pf86 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Autodesk", "Revit " + year);
            return Directory.Exists(addinPath) || Directory.Exists(pf64) || Directory.Exists(pf86);
        }

        private static bool IsAdministrator()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private static void CopyDirectoryContents(string sourceDir, string destinationDir, UpdateLogger logger)
        {
            Directory.CreateDirectory(destinationDir);

            foreach (var directory in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
            {
                var relative = directory.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                Directory.CreateDirectory(Path.Combine(destinationDir, relative));
            }

            foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
            {
                var relative = file.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var destinationFile = Path.Combine(destinationDir, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(destinationFile) ?? destinationDir);
                SafeCopyFile(file, destinationFile, logger);
            }
        }

        private static void SafeCopyFile(string sourceFile, string destinationFile, UpdateLogger logger)
        {
            Retry(
                () =>
                {
                    if (File.Exists(destinationFile))
                    {
                        File.SetAttributes(destinationFile, FileAttributes.Normal);
                    }

                    File.Copy(sourceFile, destinationFile, true);
                },
                logger,
                "Copy file " + destinationFile);
        }

        private static void ForceDeleteDirectory(string directoryPath, UpdateLogger logger)
        {
            Retry(
                () =>
                {
                    if (!Directory.Exists(directoryPath))
                    {
                        return;
                    }

                    foreach (var file in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }

                    Directory.Delete(directoryPath, true);
                },
                logger,
                "Delete directory " + directoryPath);
        }

        private static void Retry(Action action, UpdateLogger logger, string operationName)
        {
            Exception lastException = null;

            for (var attempt = 1; attempt <= 4; attempt++)
            {
                try
                {
                    action();
                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    logger.Log(operationName + " failed on attempt " + attempt + ": " + ex.Message);
                    Thread.Sleep(TimeSpan.FromMilliseconds(700));
                }
            }

            throw new InvalidOperationException(operationName + " failed after multiple attempts.", lastException);
        }

        private static string QuoteArgument(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "\"\"";
            }

            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }
    }

    internal sealed class UpdaterOptions
    {
        public string Mode { get; private set; } = "zip";
        public string SourcePath { get; private set; }
        public int WaitPid { get; private set; }
        public string LogPath { get; private set; }

        public static UpdaterOptions Parse(IReadOnlyList<string> args)
        {
            var options = new UpdaterOptions();

            for (var i = 0; i < args.Count; i++)
            {
                var arg = args[i];
                var next = i + 1 < args.Count ? args[i + 1] : null;

                if (string.Equals(arg, "--mode", StringComparison.OrdinalIgnoreCase) && next != null)
                {
                    options.Mode = next;
                    i++;
                }
                else if (string.Equals(arg, "--source", StringComparison.OrdinalIgnoreCase) && next != null)
                {
                    options.SourcePath = next;
                    i++;
                }
                else if (string.Equals(arg, "--wait-pid", StringComparison.OrdinalIgnoreCase) && next != null)
                {
                    int waitPid;
                    if (int.TryParse(next, out waitPid))
                    {
                        options.WaitPid = waitPid;
                    }

                    i++;
                }
                else if (string.Equals(arg, "--log", StringComparison.OrdinalIgnoreCase) && next != null)
                {
                    options.LogPath = next;
                    i++;
                }
            }

            if (string.IsNullOrWhiteSpace(options.LogPath))
            {
                var logRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KKY_Tool_Revit",
                    "UpdateLogs");
                Directory.CreateDirectory(logRoot);
                options.LogPath = Path.Combine(logRoot, "updater-" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log");
            }

            return options;
        }
    }

    internal sealed class UpdateLogger : IDisposable
    {
        private readonly string _logPath;
        private readonly object _sync = new object();

        public UpdateLogger(string logPath)
        {
            _logPath = logPath;
            var logDir = Path.GetDirectoryName(_logPath);
            if (!string.IsNullOrWhiteSpace(logDir))
            {
                Directory.CreateDirectory(logDir);
            }
        }

        public void Log(string message)
        {
            lock (_sync)
            {
                File.AppendAllText(
                    _logPath,
                    "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message + Environment.NewLine);
            }
        }

        public void Dispose()
        {
        }
    }
}
