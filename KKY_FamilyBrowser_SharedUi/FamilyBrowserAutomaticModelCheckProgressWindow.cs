using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Autodesk.Revit.UI;

public sealed class FamilyBrowserAutomaticModelCheckProgressWindow : Form
{
	private sealed class WindowHandleOwner : IWin32Window
	{
		private readonly IntPtr _handle;

		public IntPtr Handle => _handle;

		public WindowHandleOwner(IntPtr handle)
		{
			_handle = handle;
		}
	}

	private readonly Label _messageLabel;

	private readonly Label _countLabel;

	private readonly ProgressBar _progressBar;

	private FamilyBrowserAutomaticModelCheckProgressWindow(string projectTitle, string disciplineLabel)
	{
		bool isKorean = FamilyBrowserLanguageService.IsKorean();
		Text = isKorean ? "자동 모델 검사" : "Automatic Model Check";
		FormBorderStyle = FormBorderStyle.FixedDialog;
		StartPosition = FormStartPosition.CenterParent;
		ShowInTaskbar = false;
		ControlBox = false;
		MaximizeBox = false;
		MinimizeBox = false;
		TopMost = false;
		ClientSize = new Size(560, 188);
		BackColor = Color.FromArgb(245, 247, 251);
		Font = new Font(isKorean ? "Malgun Gothic" : "Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);

		Panel header = new Panel
		{
			Dock = DockStyle.Top,
			Height = 58,
			BackColor = Color.FromArgb(11, 18, 32)
		};
		Label titleLabel = new Label
		{
			AutoSize = false,
			Location = new Point(20, 9),
			Size = new Size(520, 24),
			ForeColor = Color.White,
			Font = new Font(Font.FontFamily, 12f, FontStyle.Bold),
			Text = isKorean ? "자동 모델 검사 진행 중" : "Automatic Model Check in progress"
		};
		Label contextLabel = new Label
		{
			AutoSize = false,
			Location = new Point(20, 34),
			Size = new Size(520, 18),
			ForeColor = Color.FromArgb(191, 219, 254),
			Text = BuildContext(projectTitle, disciplineLabel, isKorean)
		};
		header.Controls.Add(titleLabel);
		header.Controls.Add(contextLabel);

		_messageLabel = new Label
		{
			AutoEllipsis = true,
			Location = new Point(22, 78),
			Size = new Size(516, 38),
			ForeColor = Color.FromArgb(30, 41, 59),
			Text = isKorean ? "현재 프로젝트 데이터를 준비하는 중..." : "Preparing current project data..."
		};
		_progressBar = new ProgressBar
		{
			Location = new Point(22, 124),
			Size = new Size(438, 18),
			Minimum = 0,
			Maximum = 100,
			Style = ProgressBarStyle.Continuous
		};
		_countLabel = new Label
		{
			AutoSize = false,
			Location = new Point(468, 120),
			Size = new Size(70, 26),
			TextAlign = ContentAlignment.MiddleRight,
			ForeColor = Color.FromArgb(47, 107, 255),
			Font = new Font(Font.FontFamily, 9f, FontStyle.Bold),
			Text = "0%"
		};
		Label noteLabel = new Label
		{
			AutoSize = false,
			Location = new Point(22, 151),
			Size = new Size(516, 24),
			ForeColor = Color.FromArgb(100, 116, 139),
			Text = isKorean
				? "Revit 문서 검사는 API 안전을 위해 현재 화면에서 순차 처리됩니다."
				: "Revit document inspection runs sequentially on the current UI context for API safety."
		};

		Controls.Add(noteLabel);
		Controls.Add(_countLabel);
		Controls.Add(_progressBar);
		Controls.Add(_messageLabel);
		Controls.Add(header);
	}

	public static FamilyBrowserAutomaticModelCheckProgressWindow Begin(UIApplication application, string projectTitle, string disciplineLabel)
	{
		try
		{
			FamilyBrowserAutomaticModelCheckProgressWindow window = new FamilyBrowserAutomaticModelCheckProgressWindow(projectTitle, disciplineLabel);
			IWin32Window owner = ResolveOwner(application);
			if (owner == null)
			{
				window.StartPosition = FormStartPosition.CenterScreen;
				window.Show();
			}
			else
			{
				window.Show(owner);
			}
			window.RefreshVisibleSurface();
			return window;
		}
		catch
		{
			return null;
		}
	}

	public void Report(int current, int total, string message)
	{
		if (IsDisposed)
		{
			return;
		}
		int safeTotal = Math.Max(1, total);
		int safeCurrent = Math.Max(0, Math.Min(current, safeTotal));
		int percent = (int)Math.Round((double)safeCurrent / safeTotal * 100.0);
		_messageLabel.Text = string.IsNullOrWhiteSpace(message)
			? (FamilyBrowserLanguageService.IsKorean() ? "프로젝트 데이터를 검사하는 중..." : "Inspecting project data...")
			: message.Trim();
		_messageLabel.Tag = _messageLabel.Text;
		_progressBar.Value = Math.Max(_progressBar.Minimum, Math.Min(percent, _progressBar.Maximum));
		_countLabel.Text = percent.ToString() + "%";
		RefreshVisibleSurface();
	}

	protected override void OnFormClosing(FormClosingEventArgs e)
	{
		if (e.CloseReason == CloseReason.UserClosing)
		{
			e.Cancel = true;
			return;
		}
		base.OnFormClosing(e);
	}

	private void RefreshVisibleSurface()
	{
		try
		{
			Refresh();
			_progressBar.Refresh();
			_messageLabel.Refresh();
			_countLabel.Refresh();
			Update();
		}
		catch
		{
		}
	}

	private static string BuildContext(string projectTitle, string disciplineLabel, bool isKorean)
	{
		string project = string.IsNullOrWhiteSpace(projectTitle) ? "-" : projectTitle.Trim();
		string discipline = string.IsNullOrWhiteSpace(disciplineLabel) ? "-" : disciplineLabel.Trim();
		return isKorean
			? "프로젝트: " + project + "  |  검사 기준: " + discipline
			: "Project: " + project + "  |  Standard: " + discipline;
	}

	private static IWin32Window ResolveOwner(UIApplication application)
	{
		try
		{
			PropertyInfo property = application == null ? null : application.GetType().GetProperty("MainWindowHandle");
			object value = property == null ? null : property.GetValue(application, null);
			if (value is IntPtr && (IntPtr)value != IntPtr.Zero)
			{
				return new WindowHandleOwner((IntPtr)value);
			}
		}
		catch
		{
		}
		return null;
	}
}
