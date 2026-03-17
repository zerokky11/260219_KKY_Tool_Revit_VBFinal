using System;
using System.Collections.Generic;
using System.Linq;
using Drawing = System.Drawing;
using WinForms = System.Windows.Forms;
using KKY_Tool_Revit.Models;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI
{
    public sealed class PurgeProgressForm : WinForms.Form
    {
        private readonly WinForms.Label _stateLabel = new WinForms.Label();
        private readonly WinForms.Label _fileLabel = new WinForms.Label();
        private readonly WinForms.Label _iterationLabel = new WinForms.Label();
        private readonly WinForms.ProgressBar _fileProgress = new WinForms.ProgressBar();
        private readonly WinForms.ProgressBar _iterationProgress = new WinForms.ProgressBar();
        private readonly WinForms.TextBox _logText = new WinForms.TextBox();
        private readonly WinForms.Button _closeButton = new WinForms.Button();
        private readonly WinForms.Timer _timer = new WinForms.Timer();
        private int _lastLogCount = -1;

        public PurgeProgressForm()
        {
            Text = "KKY RVT Cleaner - Purge 진행";
            StartPosition = WinForms.FormStartPosition.CenterScreen;
            Width = 760;
            Height = 520;
            MinimumSize = new Drawing.Size(700, 480);
            Font = new Drawing.Font("Segoe UI", 9F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point);
            TopMost = false;
            ShowInTaskbar = true;

            var root = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new WinForms.Padding(12)
            };
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 26F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 26F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 26F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 28F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 28F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 42F));
            Controls.Add(root);

            _stateLabel.Dock = WinForms.DockStyle.Fill;
            _fileLabel.Dock = WinForms.DockStyle.Fill;
            _iterationLabel.Dock = WinForms.DockStyle.Fill;
            _fileProgress.Dock = WinForms.DockStyle.Fill;
            _iterationProgress.Dock = WinForms.DockStyle.Fill;
            _logText.Dock = WinForms.DockStyle.Fill;
            _logText.Multiline = true;
            _logText.ReadOnly = true;
            _logText.ScrollBars = WinForms.ScrollBars.Vertical;

            _closeButton.Text = "닫기";
            _closeButton.Width = 100;
            _closeButton.Height = 30;
            _closeButton.Click += (_, __) => Close();

            var buttonPanel = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.RightToLeft, WrapContents = false };
            buttonPanel.Controls.Add(_closeButton);

            root.Controls.Add(_stateLabel, 0, 0);
            root.Controls.Add(_fileLabel, 0, 1);
            root.Controls.Add(_iterationLabel, 0, 2);
            root.Controls.Add(_fileProgress, 0, 3);
            root.Controls.Add(_iterationProgress, 0, 4);
            root.Controls.Add(_logText, 0, 5);
            root.Controls.Add(buttonPanel, 0, 6);

            _timer.Interval = 500;
            _timer.Tick += (_, __) => RefreshUi();
            Load += (_, __) =>
            {
                RefreshUi();
                _timer.Start();
            };
            FormClosed += (_, __) => _timer.Stop();
        }

        private void RefreshUi()
        {
            PurgeBatchProgressSnapshot snapshot = PurgeUiBatchService.GetProgressSnapshot();
            IList<string> logs = App.GetSharedLogLinesSnapshot();

            _stateLabel.Text = "상태: " + (snapshot.StateName ?? "대기") + (!string.IsNullOrWhiteSpace(snapshot.Message) ? " / " + snapshot.Message : string.Empty);
            _fileLabel.Text = "파일 진행: " + Math.Max(snapshot.CurrentFileIndex, 0) + " / " + Math.Max(snapshot.TotalFiles, 0) + (string.IsNullOrWhiteSpace(snapshot.CurrentFileName) ? string.Empty : " / " + snapshot.CurrentFileName);
            _iterationLabel.Text = "반복 진행: " + Math.Max(snapshot.CurrentIteration, 0) + " / " + Math.Max(snapshot.TotalIterations, 0);

            _fileProgress.Minimum = 0;
            _fileProgress.Maximum = Math.Max(snapshot.TotalFiles, 1);
            _fileProgress.Value = Math.Max(0, Math.Min(_fileProgress.Maximum, snapshot.CurrentFileIndex));

            _iterationProgress.Minimum = 0;
            _iterationProgress.Maximum = Math.Max(snapshot.TotalIterations, 1);
            _iterationProgress.Value = Math.Max(0, Math.Min(_iterationProgress.Maximum, snapshot.CurrentIteration));

            if (logs.Count != _lastLogCount)
            {
                IEnumerable<string> recent = logs.Skip(Math.Max(0, logs.Count - 120));
                _logText.Lines = recent.ToArray();
                _logText.SelectionStart = _logText.TextLength;
                _logText.ScrollToCaret();
                _lastLogCount = logs.Count;
            }
        }
    }
}
