using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var options = UpdaterOptions.Parse(args);
            using (var logger = new UpdateLogger(options.LogPath))
            using (var progressWindow = string.Equals(options.Mode, "zip", StringComparison.OrdinalIgnoreCase)
                ? new UpdateProgressWindow()
                : null)
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
                        progressWindow?.ShowWindow();
                        progressWindow?.UpdateProgress(6, "Revit 종료 확인됨", "업데이트 준비를 시작합니다.");
                        ApplyZipPackage(options.SourcePath, logger, progressWindow);
                        progressWindow?.UpdateProgress(100, "업데이트 적용 완료", "완료 창을 확인한 뒤 Revit을 다시 실행해 주세요.");
                        Thread.Sleep(450);
                        ShowSuccessMessage();
                    }

                    logger.Log("Updater completed successfully.");
                    return 0;
                }
                catch (Exception ex)
                {
                    progressWindow?.CloseWindow();
                    logger.Log("Updater failed: " + ex);
                    ShowFailureMessage(options.LogPath, ex);
                    return 1;
                }
            }
        }

        private static void ShowSuccessMessage()
        {
            ShowStatusDialog(
                "업데이트 완료",
                "Revit 업데이트가 적용되었습니다.",
                "이제 Revit을 다시 실행하면 최신 버전이 로드됩니다.",
                null,
                false);
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

            ShowStatusDialog(
                "업데이트 실패",
                "업데이트 적용 중 문제가 발생했습니다.",
                message,
                logPath,
                true);
        }

        private static void ShowStatusDialog(string badgeText, string title, string message, string detailText, bool isError)
        {
            using (var form = new Form())
            using (var root = new TableLayoutPanel())
            using (var badge = new Label())
            using (var titleLabel = new Label())
            using (var messageLabel = new Label())
            using (var hintPanel = new Panel())
            using (var hintLabel = new Label())
            using (var detailBox = new TextBox())
            using (var buttonPanel = new Panel())
            using (var okButton = new Button())
            {
                var borderColor = isError ? Color.FromArgb(220, 120, 120) : Color.FromArgb(113, 181, 142);
                var badgeBack = isError ? Color.FromArgb(254, 235, 235) : Color.FromArgb(232, 248, 237);
                var badgeFore = isError ? Color.FromArgb(176, 62, 62) : Color.FromArgb(39, 124, 78);

                form.Text = "KKY Tool Revit 업데이트";
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ShowInTaskbar = true;
                form.TopMost = true;
                form.BackColor = Color.White;
                form.ClientSize = new Size(540, string.IsNullOrWhiteSpace(detailText) ? 290 : 390);
                form.Padding = new Padding(18);
                form.Font = new Font("Malgun Gothic", 9.5F, FontStyle.Regular);
                form.Paint += (sender, args) =>
                {
                    using (var pen = new Pen(borderColor, 2))
                    {
                        args.Graphics.DrawRectangle(pen, 1, 1, form.ClientSize.Width - 3, form.ClientSize.Height - 3);
                    }
                };

                root.Dock = DockStyle.Fill;
                root.ColumnCount = 1;
                root.RowCount = string.IsNullOrWhiteSpace(detailText) ? 5 : 6;
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                IfNotEmptyAddDetailRow(root, detailText);
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                badge.AutoSize = true;
                badge.Text = badgeText;
                badge.BackColor = badgeBack;
                badge.ForeColor = badgeFore;
                badge.Font = new Font("Malgun Gothic", 9.75F, FontStyle.Bold);
                badge.Padding = new Padding(12, 6, 12, 6);
                badge.Margin = new Padding(0, 0, 0, 12);

                titleLabel.AutoSize = true;
                titleLabel.Text = title;
                titleLabel.Font = new Font("Malgun Gothic", 17F, FontStyle.Bold);
                titleLabel.ForeColor = Color.FromArgb(24, 38, 58);
                titleLabel.Margin = new Padding(0, 0, 0, 10);

                messageLabel.Dock = DockStyle.Fill;
                messageLabel.Text = message;
                messageLabel.Font = new Font("Malgun Gothic", 10.5F, FontStyle.Regular);
                messageLabel.ForeColor = Color.FromArgb(66, 78, 96);
                messageLabel.Margin = new Padding(0);
                messageLabel.Padding = new Padding(0, 4, 0, 0);

                hintPanel.Dock = DockStyle.Fill;
                hintPanel.BackColor = isError ? Color.FromArgb(255, 245, 245) : Color.FromArgb(242, 249, 245);
                hintPanel.Margin = new Padding(0, 14, 0, 0);
                hintPanel.Padding = new Padding(14, 12, 14, 12);

                hintLabel.Dock = DockStyle.Fill;
                hintLabel.Font = new Font("Malgun Gothic", 9.5F, FontStyle.Bold);
                hintLabel.ForeColor = isError ? Color.FromArgb(141, 50, 50) : Color.FromArgb(35, 95, 63);
                hintLabel.Text = isError
                    ? "문제가 계속되면 로그 경로를 확인한 뒤 다시 시도해 주세요."
                    : "완료 창을 확인한 뒤 Revit을 다시 실행해 주세요.";
                hintPanel.Controls.Add(hintLabel);

                detailBox.ReadOnly = true;
                detailBox.Multiline = true;
                detailBox.ScrollBars = ScrollBars.Vertical;
                detailBox.BorderStyle = BorderStyle.FixedSingle;
                detailBox.BackColor = Color.FromArgb(248, 250, 252);
                detailBox.ForeColor = Color.FromArgb(60, 70, 84);
                detailBox.Font = new Font("Consolas", 9F, FontStyle.Regular);
                detailBox.Dock = DockStyle.Fill;
                detailBox.Text = detailText ?? string.Empty;
                detailBox.Margin = new Padding(0, 14, 0, 0);
                detailBox.Visible = !string.IsNullOrWhiteSpace(detailText);

                buttonPanel.Dock = DockStyle.Fill;
                buttonPanel.Height = 44;
                buttonPanel.Margin = new Padding(0, 16, 0, 0);

                okButton.Text = "확인";
                okButton.DialogResult = DialogResult.OK;
                okButton.Anchor = AnchorStyles.Right | AnchorStyles.Top;
                okButton.Size = new Size(112, 36);
                okButton.Location = new Point(form.ClientSize.Width - form.Padding.Horizontal - okButton.Width, 0);
                okButton.BackColor = isError ? Color.FromArgb(198, 78, 78) : Color.FromArgb(44, 122, 88);
                okButton.ForeColor = Color.White;
                okButton.FlatStyle = FlatStyle.Flat;
                okButton.FlatAppearance.BorderSize = 0;
                okButton.Font = new Font("Malgun Gothic", 10F, FontStyle.Bold);
                okButton.Text = isError ? "확인" : "Revit 다시 실행 준비";

                buttonPanel.Controls.Add(okButton);

                root.Controls.Add(badge, 0, 0);
                root.Controls.Add(titleLabel, 0, 1);
                root.Controls.Add(messageLabel, 0, 2);
                root.Controls.Add(hintPanel, 0, 3);
                if (detailBox.Visible)
                {
                    root.Controls.Add(detailBox, 0, 4);
                    root.Controls.Add(buttonPanel, 0, 5);
                }
                else
                {
                    root.Controls.Add(buttonPanel, 0, 4);
                }

                form.Controls.Add(root);
                form.AcceptButton = okButton;
                form.CancelButton = okButton;
                form.Shown += (sender, args) =>
                {
                    form.Activate();
                    form.BringToFront();
                    okButton.Focus();
                };
                form.ShowDialog();
            }
        }

        private static void IfNotEmptyAddDetailRow(TableLayoutPanel root, string detailText)
        {
            if (string.IsNullOrWhiteSpace(detailText))
            {
                return;
            }

            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 108F));
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

        private static void ApplyZipPackage(string zipPath, UpdateLogger logger, UpdateProgressWindow progressWindow)
        {
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var extractRoot = Path.Combine(Path.GetTempPath(), "KKY_Tool_Revit", "UpdaterExtract", stamp);
            var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var addinsRoot = Path.Combine(programData, "Autodesk", "Revit", "Addins");

            logger.Log("Extract root: " + extractRoot);
            progressWindow?.UpdateProgress(12, "업데이트 패키지 준비 중", "기존 임시 폴더를 정리하고 있습니다.");

            if (Directory.Exists(extractRoot))
            {
                ForceDeleteDirectory(extractRoot, logger);
            }

            Directory.CreateDirectory(extractRoot);
            progressWindow?.UpdateProgress(24, "압축 해제 중", "업데이트 파일을 임시 폴더에 풀고 있습니다.");
            ZipFile.ExtractToDirectory(zipPath, extractRoot);
            logger.Log("ZIP extracted.");
            progressWindow?.UpdateProgress(34, "압축 해제 완료", "설치된 Revit 버전을 확인하는 중입니다.");

            var installedYears = new[] { "2019", "2021", "2023", "2025" }
                .Where(IsRevitInstalled)
                .ToList();
            var totalTargets = Math.Max(1, installedYears.Count);
            var appliedCount = 0;

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
                appliedCount += 1;
                var percent = 34 + (int)Math.Round((58.0 * appliedCount) / totalTargets);
                progressWindow?.UpdateProgress(percent, "업데이트 적용 중", "Revit " + year + " 경로에 파일을 적용하고 있습니다.");
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

            progressWindow?.UpdateProgress(96, "업데이트 적용 확인 중", "모든 설치 경로 반영을 마무리하고 있습니다.");
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

    internal sealed class UpdateProgressWindow : IDisposable
    {
        private readonly Form _form;
        private readonly Label _title;
        private readonly Label _detail;
        private readonly ProgressBar _progressBar;
        private readonly Label _progressText;

        public UpdateProgressWindow()
        {
            _form = new Form
            {
                Text = "KKY Tool Revit 업데이트",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = true,
                TopMost = true,
                BackColor = Color.White,
                ClientSize = new Size(520, 240),
                Padding = new Padding(18),
                Font = new Font("Malgun Gothic", 9.5F, FontStyle.Regular)
            };

            _form.Paint += (sender, args) =>
            {
                using (var pen = new Pen(Color.FromArgb(106, 177, 141), 2))
                {
                    args.Graphics.DrawRectangle(pen, 1, 1, _form.ClientSize.Width - 3, _form.ClientSize.Height - 3);
                }
            };

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var badge = new Label
            {
                AutoSize = true,
                Text = "업데이트 적용 중",
                BackColor = Color.FromArgb(232, 248, 237),
                ForeColor = Color.FromArgb(39, 124, 78),
                Font = new Font("Malgun Gothic", 9.75F, FontStyle.Bold),
                Padding = new Padding(12, 6, 12, 6),
                Margin = new Padding(0, 0, 0, 12)
            };

            _title = new Label
            {
                AutoSize = true,
                Text = "업데이트를 준비하는 중입니다.",
                Font = new Font("Malgun Gothic", 17F, FontStyle.Bold),
                ForeColor = Color.FromArgb(24, 38, 58),
                Margin = new Padding(0, 0, 0, 10)
            };

            _detail = new Label
            {
                Dock = DockStyle.Fill,
                Text = "잠시만 기다려 주세요.",
                Font = new Font("Malgun Gothic", 10.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(66, 78, 96),
                Margin = new Padding(0, 0, 0, 16)
            };

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Height = 20,
                Style = ProgressBarStyle.Continuous,
                Maximum = 100,
                Minimum = 0,
                Value = 0,
                Margin = new Padding(0, 0, 0, 8)
            };

            _progressText = new Label
            {
                AutoSize = true,
                Text = "0%",
                Font = new Font("Consolas", 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(44, 122, 88),
                Margin = new Padding(0)
            };

            root.Controls.Add(badge, 0, 0);
            root.Controls.Add(_title, 0, 1);
            root.Controls.Add(_detail, 0, 2);
            root.Controls.Add(_progressBar, 0, 3);
            root.Controls.Add(_progressText, 0, 4);

            _form.Controls.Add(root);
        }

        public void ShowWindow()
        {
            if (_form.Visible)
            {
                return;
            }

            _form.Show();
            _form.Activate();
            _form.BringToFront();
            Application.DoEvents();
        }

        public void UpdateProgress(int percent, string title, string detail)
        {
            var safePercent = Math.Max(0, Math.Min(100, percent));
            _title.Text = title ?? string.Empty;
            _detail.Text = detail ?? string.Empty;
            _progressBar.Value = safePercent;
            _progressText.Text = safePercent + "%";
            _form.Activate();
            _form.BringToFront();
            _form.Refresh();
            Application.DoEvents();
        }

        public void CloseWindow()
        {
            if (_form.IsDisposed)
            {
                return;
            }

            _form.Hide();
            Application.DoEvents();
        }

        public void Dispose()
        {
            if (_form != null && !_form.IsDisposed)
            {
                _form.Dispose();
            }
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
