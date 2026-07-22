using System;
using System.Drawing;
using System.Windows.Forms;

internal sealed class FamilyBrowserDetachedDetailWindow : Form, IFamilyBrowserThemeAware
{
	private readonly WebBrowser _browser;
	private readonly Action<string> _measurementUnitChanged;

	private bool _isKorean;

	private FamilyBrowserUiTheme _theme;

	public FamilyBrowserDetachedDetailWindow(bool isKorean, Action<string> measurementUnitChanged = null)
	{
		_isKorean = isKorean;
		_measurementUnitChanged = measurementUnitChanged;
		_theme = FamilyBrowserUiThemeService.Load();
		Text = DefaultTitle();
		AutoScaleMode = AutoScaleMode.Dpi;
		AutoScaleDimensions = new SizeF(96f, 96f);
		Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
		StartPosition = FormStartPosition.CenterScreen;
		FormBorderStyle = FormBorderStyle.Sizable;
		ShowInTaskbar = true;
		MinimizeBox = true;
		MaximizeBox = true;
		Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
		ClientSize = new Size(Math.Max(680, Math.Min(980, workingArea.Width - 140)), Math.Max(640, Math.Min(900, workingArea.Height - 140)));
		MinimumSize = new Size(640, 560);
		_browser = new WebBrowser
		{
			Dock = DockStyle.Fill,
			ScriptErrorsSuppressed = true,
			AllowNavigation = true,
			AllowWebBrowserDrop = false,
			IsWebBrowserContextMenuEnabled = true,
			WebBrowserShortcutsEnabled = true
		};
		_browser.Navigating += BrowserNavigating;
		Controls.Add(_browser);
		ApplyWindowPalette();
	}

	public void SetLanguage(bool isKorean)
	{
		_isKorean = isKorean;
		Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
		if (string.IsNullOrWhiteSpace(Text) || string.Equals(Text, "Selected Item Detail", StringComparison.OrdinalIgnoreCase) || Text.Contains("?"))
		{
			Text = DefaultTitle();
		}
	}

	public void SetDetailHtml(string title, string html)
	{
		Text = (string.IsNullOrWhiteSpace(title) || title == "-") ? DefaultTitle() : title;
		_browser.DocumentText = html ?? string.Empty;
	}

	public void ApplyTheme(FamilyBrowserUiTheme theme)
	{
		_theme = theme;
		ApplyWindowPalette();
		try
		{
			if (_browser.Document != null)
			{
				string code = FamilyBrowserUiThemeService.Code(theme);
				_browser.Document.InvokeScript("eval", new object[] { "(function(){var b=document.body;if(!b)return;b.className=(b.className||'').replace(/(^|\\s)theme-light(?=\\s|$)/g,' ').replace(/(^|\\s)theme-dark(?=\\s|$)/g,' ').replace(/^\\s+|\\s+$/g,'').replace(/\\s+/g,' ')+' theme-" + code + "';b.setAttribute('data-theme','" + code + "');})()" });
			}
		}
		catch
		{
		}
	}

	public void SyncMeasurementDisplayUnit(string unit)
	{
		try
		{
			if (_browser.Document != null)
			{
				_browser.Document.InvokeScript("setSystemDisplayUnitFromHost", new object[] { FamilyBrowserMeasurementUnitPreferenceService.Normalize(unit) });
			}
		}
		catch
		{
		}
	}

	private void ApplyWindowPalette()
	{
		BackColor = _theme == FamilyBrowserUiTheme.Dark ? Color.FromArgb(11, 18, 32) : Color.FromArgb(245, 247, 251);
	}

	private string DefaultTitle()
	{
		return _isKorean ? "선택 항목 상세" : "Selected Item Detail";
	}

	private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
	{
		if ((object)e.Url == null)
		{
			return;
		}
		string raw = e.Url.ToString();
		if (raw.StartsWith("kkyfb:", StringComparison.OrdinalIgnoreCase) || raw.StartsWith("about:kkyfb:", StringComparison.OrdinalIgnoreCase))
		{
			e.Cancel = true;
			string prefix = raw.StartsWith("about:kkyfb:", StringComparison.OrdinalIgnoreCase) ? "about:kkyfb:" : "kkyfb:";
			string action = raw.Substring(prefix.Length).Trim('/', ' ');
			if (action.StartsWith("measurement-unit/", StringComparison.OrdinalIgnoreCase))
			{
				string unit = FamilyBrowserMeasurementUnitPreferenceService.Normalize(Uri.UnescapeDataString(action.Substring("measurement-unit/".Length)));
				FamilyBrowserMeasurementUnitPreferenceService.Save(unit);
				if (_measurementUnitChanged != null)
				{
					_measurementUnitChanged(unit);
				}
				SyncMeasurementDisplayUnit(unit);
			}
		}
	}
}
