using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

internal sealed class StandardListSheetSelectionHtmlForm : Form
{
	private readonly bool _isKorean;
	private readonly List<string> _sheetNames;
	private readonly string _currentSheetName;
	private readonly WebBrowser _browser;

	public string SelectedSheetName { get; private set; }

	public StandardListSheetSelectionHtmlForm(bool isKorean, string titleText, IEnumerable<string> sheetNames, string currentSheetName)
	{
		_isKorean = isKorean;
		_sheetNames = (sheetNames ?? Enumerable.Empty<string>()).Where(name => !string.IsNullOrWhiteSpace(name)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
		_currentSheetName = (currentSheetName ?? string.Empty).Trim();
		SelectedSheetName = _sheetNames.FirstOrDefault(name => string.Equals(name, _currentSheetName, StringComparison.OrdinalIgnoreCase)) ?? _sheetNames.FirstOrDefault() ?? string.Empty;
		Text = titleText ?? string.Empty;
		AutoScaleMode = AutoScaleMode.Dpi;
		AutoScaleDimensions = new SizeF(96f, 96f);
		Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
		StartPosition = FormStartPosition.CenterParent;
		FormBorderStyle = FormBorderStyle.Sizable;
		ShowInTaskbar = false;
		MinimizeBox = false;
		MaximizeBox = true;
		Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
		ClientSize = new Size(Math.Max(560, Math.Min(760, workingArea.Width - 160)), Math.Max(420, Math.Min(620, workingArea.Height - 160)));
		MinimumSize = new Size(500, 380);
		_browser = new WebBrowser
		{
			Dock = DockStyle.Fill,
			ScriptErrorsSuppressed = true,
			AllowNavigation = true,
			AllowWebBrowserDrop = false,
			IsWebBrowserContextMenuEnabled = false,
			WebBrowserShortcutsEnabled = false
		};
		_browser.Navigating += BrowserNavigating;
		Controls.Add(_browser);
	}

	protected override void OnShown(EventArgs e)
	{
		base.OnShown(e);
		_browser.DocumentText = BuildHtml();
	}

	private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
	{
		if (e == null || e.Url == null)
		{
			return;
		}
		string command = string.Empty;
		string rawIndex = string.Empty;
		if (string.Equals(e.Url.Scheme, "kkysheet", StringComparison.OrdinalIgnoreCase))
		{
			command = e.Url.Host ?? string.Empty;
			rawIndex = (e.Url.AbsolutePath ?? string.Empty).Trim('/');
		}
		else
		{
			string raw = e.Url.AbsoluteUri ?? string.Empty;
			if (raw.StartsWith("about:kkysheet://", StringComparison.OrdinalIgnoreCase))
			{
				string payload = raw.Substring("about:kkysheet://".Length).Trim('/');
				int slash = payload.IndexOf('/');
				command = (slash >= 0) ? payload.Substring(0, slash) : payload;
				rawIndex = (slash >= 0) ? payload.Substring(slash + 1).Trim('/') : string.Empty;
			}
		}
		if (string.IsNullOrWhiteSpace(command))
		{
			return;
		}
		command = Uri.UnescapeDataString(command.Trim('/', ' '));
		rawIndex = Uri.UnescapeDataString((rawIndex ?? string.Empty).Trim('/', ' '));
		e.Cancel = true;
		if (string.Equals(command, "select", StringComparison.OrdinalIgnoreCase))
		{
			if (int.TryParse(rawIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) && index >= 0 && index < _sheetNames.Count)
			{
				SelectedSheetName = _sheetNames[index];
				DialogResult = DialogResult.OK;
				Close();
			}
			return;
		}
		DialogResult = DialogResult.Cancel;
		Close();
	}

	private string BuildHtml()
	{
		FamilyBrowserUiTheme theme = FamilyBrowserUiThemeService.Load();
		StringBuilder builder = new StringBuilder();
		builder.Append("<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'><style>");
		builder.Append("html,body{margin:0;width:100%;height:100%;overflow:hidden;font-family:'Malgun Gothic','Segoe UI',Arial,sans-serif;background:#eef3f0;color:#17231f;}*{box-sizing:border-box}.shell{position:absolute;left:0;right:0;top:0;bottom:0;background:#fff;}.head{position:absolute;left:0;right:0;top:0;height:78px;background:#20362f;color:#fff;padding:16px 24px;border-left:7px solid #22b47f;}.head h1{font-size:22px;line-height:1.2;margin:0 110px 7px 0;font-weight:850;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.head p{font-size:13px;line-height:1.35;margin:0;color:#d7e8df;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.close{position:absolute;right:22px;top:22px;text-decoration:none;color:#fff;border:1px solid rgba(255,255,255,.45);border-radius:6px;padding:8px 14px;font-weight:850}.body{position:absolute;left:0;right:0;top:78px;bottom:54px;overflow:auto;padding:20px 24px;background:#f8fbf9}.sheet{display:block;text-decoration:none;color:#173a30;background:#fff;border:1px solid #d5e3dd;border-radius:8px;padding:13px 15px;margin:0 0 10px 0;font-size:14px;font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.sheet:hover{border-color:#1c9765;background:#f0fbf5}.sheet.active{border-color:#1c9765;background:#e9f8f0;box-shadow:inset 4px 0 0 #1c9765}.meta{font-size:12px;color:#61756d;margin:0 0 14px 0}.foot{position:absolute;left:0;right:0;bottom:0;height:54px;background:#fff;border-top:1px solid #d8e1dc;text-align:right;padding:9px 20px}.btn{display:inline-block;text-decoration:none;border:1px solid #becfc7;border-radius:7px;background:#fff;color:#1d372e;font-weight:850;font-size:13px;line-height:32px;min-width:86px;text-align:center;padding:0 14px}.empty{padding:34px 16px;text-align:center;color:#63766e;background:#fff;border:1px dashed #cad8d2;border-radius:8px}");
		builder.Append(FamilyBrowserUiThemeService.ThemeCss());
		builder.Append("</style><script>").Append(FamilyBrowserOverflowTitleScript.Script()).Append("</script></head><body data-theme='").Append(Attr(FamilyBrowserUiThemeService.Code(theme))).Append("' class='fb-sheet-selection ").Append(Attr(FamilyBrowserUiThemeService.BodyClass(theme))).Append("'><div class='shell'><div class='head'><h1>").Append(Html(L("Select worksheet", "\uC2DC\uD2B8 \uC120\uD0DD"))).Append("</h1><p>")
			.Append(Html(L("Choose the sheet that contains the standard family list.", "\uD45C\uC900 \uD328\uBC00\uB9AC \uBAA9\uB85D\uC774 \uB4E4\uC5B4\uC788\uB294 \uC2DC\uD2B8\uB97C \uC120\uD0DD\uD558\uC138\uC694.")))
			.Append("</p><a class='close' href='kkysheet://cancel/'>")
			.Append(Html(L("Cancel", "\uCDE8\uC18C")))
			.Append("</a></div><div class='body'>");
		builder.Append("<div class='meta'>").Append(Html(string.Format(CultureInfo.InvariantCulture, L("{0} worksheet(s) found.", "{0}\uAC1C \uC2DC\uD2B8\uB97C \uCC3E\uC558\uC2B5\uB2C8\uB2E4."), _sheetNames.Count))).Append("</div>");
		if (_sheetNames.Count == 0)
		{
			builder.Append("<div class='empty'>").Append(Html(L("No worksheets were found in this workbook.", "\uC774 Excel \uD30C\uC77C\uC5D0\uC11C \uC2DC\uD2B8\uB97C \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4."))).Append("</div>");
		}
		else
		{
			for (int index = 0; index < _sheetNames.Count; index++)
			{
				string sheetName = _sheetNames[index];
				bool active = string.Equals(sheetName, SelectedSheetName, StringComparison.OrdinalIgnoreCase);
				builder.Append("<a class='sheet");
				if (active)
				{
					builder.Append(" active");
				}
				builder.Append("' title='").Append(Attr(sheetName)).Append("' href='kkysheet://select/")
					.Append(index.ToString(CultureInfo.InvariantCulture))
					.Append("'>")
					.Append(Html(sheetName))
					.Append("</a>");
			}
		}
		builder.Append("</div><div class='foot'><a class='btn' href='kkysheet://cancel/'>").Append(Html(L("Cancel", "\uCDE8\uC18C"))).Append("</a></div></div></body></html>");
		return builder.ToString();
	}

	private string L(string englishText, string koreanText)
	{
		return _isKorean ? koreanText : englishText;
	}

	private static string Html(string value)
	{
		return (value ?? string.Empty).Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
	}

	private static string Attr(string value)
	{
		return Html(value).Replace("'", "&#39;");
	}
}
