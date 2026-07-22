using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

internal sealed class FamilyBrowserHtmlDialogHost : Form
{
	private const int WmNcLButtonDown = 0x00A1;
	private const int HtCaption = 0x0002;

	private readonly bool _isKorean;
	private readonly MessageBoxButtons _buttons;
	private readonly MessageBoxDefaultButton _defaultButton;
	private readonly DialogResult _closeResult;
	private readonly string _copyDetailsText;
	private readonly string _reportPath;
	private readonly WebBrowser _browser;

	public event EventHandler AuxiliaryActionRequested;

	public FamilyBrowserHtmlDialogHost(
		bool isKorean,
		string caption,
		string contentDocument,
		MessageBoxButtons buttons,
		MessageBoxIcon icon,
		MessageBoxDefaultButton defaultButton,
		string positiveButtonText,
		string negativeButtonText,
		string copyDetailsText,
		string reportPath,
		Size preferredClientSize,
		Size minimumSize,
		bool showMaximize)
		: this(isKorean, caption, contentDocument, buttons, icon, defaultButton, positiveButtonText, negativeButtonText, copyDetailsText, reportPath, preferredClientSize, minimumSize, showMaximize, string.Empty)
	{
	}

	public FamilyBrowserHtmlDialogHost(
		bool isKorean,
		string caption,
		string contentDocument,
		MessageBoxButtons buttons,
		MessageBoxIcon icon,
		MessageBoxDefaultButton defaultButton,
		string positiveButtonText,
		string negativeButtonText,
		string copyDetailsText,
		string reportPath,
		Size preferredClientSize,
		Size minimumSize,
		bool showMaximize,
		string auxiliaryActionText)
	{
		_isKorean = isKorean;
		_buttons = buttons;
		_defaultButton = defaultButton;
		_closeResult = buttons == MessageBoxButtons.YesNo ? DialogResult.No : DialogResult.OK;
		_copyDetailsText = copyDetailsText ?? string.Empty;
		_reportPath = reportPath ?? string.Empty;

		Text = caption ?? string.Empty;
		AutoScaleMode = AutoScaleMode.Dpi;
		AutoScaleDimensions = new SizeF(96f, 96f);
		StartPosition = FormStartPosition.CenterParent;
		FormBorderStyle = FormBorderStyle.None;
		ShowInTaskbar = false;
		MinimizeBox = false;
		MaximizeBox = false;
		KeyPreview = true;
		BackColor = Color.FromArgb(193, 204, 223);
		Padding = new Padding(1);

		Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
		int width = Math.Min(Math.Max(minimumSize.Width, preferredClientSize.Width), Math.Max(520, workingArea.Width - 80));
		int height = Math.Min(Math.Max(minimumSize.Height, preferredClientSize.Height), Math.Max(360, workingArea.Height - 80));
		ClientSize = new Size(width, height);
		MinimumSize = new Size(Math.Min(minimumSize.Width, width), Math.Min(minimumSize.Height, height));

		_browser = new WebBrowser
		{
			Dock = DockStyle.Fill,
			ScriptErrorsSuppressed = true,
			AllowWebBrowserDrop = false,
			IsWebBrowserContextMenuEnabled = true,
			WebBrowserShortcutsEnabled = true,
			TabStop = true,
			Margin = new Padding(0)
		};
		_browser.Navigating += BrowserNavigating;
		_browser.DocumentCompleted += BrowserDocumentCompleted;
		Controls.Add(_browser);
		_browser.DocumentText = BuildDocument(
			_isKorean,
			caption,
			contentDocument,
			buttons,
			icon,
			defaultButton,
			positiveButtonText,
			negativeButtonText,
			!string.IsNullOrWhiteSpace(_copyDetailsText),
			!string.IsNullOrWhiteSpace(_reportPath),
			showMaximize,
			auxiliaryActionText);
	}

	public static string BuildDocument(
		bool isKorean,
		string caption,
		string contentDocument,
		MessageBoxButtons buttons,
		MessageBoxIcon icon,
		MessageBoxDefaultButton defaultButton,
		string positiveButtonText,
		string negativeButtonText,
		bool showCopyDetails,
		bool showReportActions,
		bool showMaximize)
	{
		return BuildDocument(isKorean, caption, contentDocument, buttons, icon, defaultButton, positiveButtonText, negativeButtonText, showCopyDetails, showReportActions, showMaximize, string.Empty);
	}

	public static string BuildDocument(
		bool isKorean,
		string caption,
		string contentDocument,
		MessageBoxButtons buttons,
		MessageBoxIcon icon,
		MessageBoxDefaultButton defaultButton,
		string positiveButtonText,
		string negativeButtonText,
		bool showCopyDetails,
		bool showReportActions,
		bool showMaximize,
		string auxiliaryActionText)
	{
		GeneratedDocumentParts parts = ExtractGeneratedDocument(contentDocument);
		string positiveText = string.IsNullOrWhiteSpace(positiveButtonText)
			? (buttons == MessageBoxButtons.YesNo ? (isKorean ? "예" : "Yes") : (isKorean ? "확인" : "OK"))
			: positiveButtonText;
		string negativeText = string.IsNullOrWhiteSpace(negativeButtonText) ? (isKorean ? "아니오" : "No") : negativeButtonText;
		string defaultAction = buttons == MessageBoxButtons.YesNo && defaultButton == MessageBoxDefaultButton.Button2 ? "decline" : "accept";
		string kind = ResolveKind(icon);
		string accent = ResolveAccent(icon);
		string hint = buttons == MessageBoxButtons.YesNo
			? (isKorean ? "내용을 확인한 뒤 작업을 선택합니다." : "Review the details, then choose an action.")
			: (isKorean ? "확인하면 작업 화면으로 돌아갑니다." : "Confirm to return to the dashboard.");

		string bodyOpen = string.IsNullOrWhiteSpace(parts.BodyOpenTag) ? "<body>" : parts.BodyOpenTag;
		bodyOpen = AddBodyAttribute(bodyOpen, "data-dialog-shell", "full-html");
		bodyOpen = AddBodyAttribute(bodyOpen, "data-dialog-kind", kind);

		StringBuilder html = new StringBuilder();
		html.AppendLine("<!doctype html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"/><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/><meta charset=\"utf-8\"/>");
		html.AppendLine("<style>");
		if (!string.IsNullOrWhiteSpace(parts.Style))
		{
			html.AppendLine(parts.Style);
		}
		html.AppendLine("*{box-sizing:border-box}html,body{width:100%;height:100%;margin:0;padding:0;overflow:hidden;background:#f5f7fb;color:#111827;font-family:'Malgun Gothic','Segoe UI',sans-serif}.fb-dialog-shell{position:relative;width:100%;height:100%;overflow:hidden;background:#f5f7fb}.fb-dialog-header{position:absolute;left:0;right:0;top:0;height:72px;background:#0b1220;color:#fff;border-bottom:1px solid #243654;cursor:default}.fb-dialog-accent{position:absolute;left:0;top:0;bottom:0;width:6px;background:" + accent + "}.fb-dialog-title-stack{position:absolute;left:25px;right:112px;top:13px;height:48px;overflow:hidden}.fb-dialog-title{font-size:17px;line-height:25px;font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.fb-dialog-subtitle{margin-top:2px;color:#a9bfdf;font-size:11px;line-height:17px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.fb-dialog-window-actions{position:absolute;right:12px;top:14px;height:42px;white-space:nowrap}.fb-dialog-window-button{display:inline-block;width:40px;height:40px;margin-left:4px;border:1px solid #314462;color:#dbe7fb;background:#121c33;text-align:center;line-height:37px;text-decoration:none;font-size:16px;font-weight:700;vertical-align:top}.fb-dialog-window-button:hover,.fb-dialog-window-button:focus{background:#20304d;border-color:#5d78a4;outline:none}.fb-dialog-body{position:absolute;left:0;right:0;top:72px;bottom:74px;overflow:auto;background:#f5f7fb}.fb-dialog-content{min-height:100%;overflow:hidden}.fb-dialog-footer{position:absolute;left:0;right:0;bottom:0;height:74px;padding:12px 18px;background:#fff;border-top:1px solid #d6dde9}.fb-dialog-footer-table{width:100%;height:48px;border-collapse:collapse;table-layout:auto}.fb-dialog-status-cell{width:100%;padding:0 14px 0 0;vertical-align:middle}.fb-dialog-status{display:block;max-height:38px;overflow:hidden;color:#64748b;font-size:11px;line-height:18px}.fb-dialog-actions-cell{white-space:nowrap;text-align:right;vertical-align:middle}.fb-dialog-button{display:inline-block;min-width:92px;height:38px;margin-left:8px;padding:0 16px;border:1px solid #c1ccdf;background:#fff;color:#1f2937;text-decoration:none;text-align:center;line-height:36px;font-size:13px;font-weight:700;white-space:nowrap;vertical-align:middle}.fb-dialog-button:hover,.fb-dialog-button:focus{border-color:#7f9bcc;background:#f1f5fb;outline:none}.fb-dialog-button.primary{border-color:#2f6bff;background:#2f6bff;color:#fff}.fb-dialog-button.primary:hover,.fb-dialog-button.primary:focus{border-color:#1f55d1;background:#1f55d1}.fb-dialog-button.default{box-shadow:0 0 0 2px #dbe8ff}.fb-dialog-button[aria-disabled=true]{border-color:#dce3ee;background:#f5f7fb;color:#94a3b8;cursor:default}.fb-dialog-body .wrap{min-height:100%}@media(max-width:760px){.fb-dialog-footer{padding-left:12px;padding-right:12px}.fb-dialog-status-cell{display:none}.fb-dialog-button{min-width:82px;padding-left:12px;padding-right:12px}.fb-dialog-title-stack{right:100px}}");
		html.AppendLine("</style></head>");
		html.Append(bodyOpen).Append("<div id=\"dialogShell\" class=\"fb-dialog-shell\">");
		html.Append("<div id=\"dialogHeader\" class=\"fb-dialog-header\" onmousedown=\"return kkyDialogDrag(event);\"><span class=\"fb-dialog-accent\"></span><div class=\"fb-dialog-title-stack\"><div id=\"dialogTitle\" class=\"fb-dialog-title\">").Append(Html(string.IsNullOrWhiteSpace(caption) ? (isKorean ? "KKY 패밀리 브라우저" : "KKY Family Browser") : caption)).Append("</div><div id=\"dialogSubtitle\" class=\"fb-dialog-subtitle\">").Append(Html(ResolveKindText(isKorean, icon))).Append("</div></div><div class=\"fb-dialog-window-actions\">");
		if (showMaximize)
		{
			html.Append("<a id=\"dialogMaximize\" class=\"fb-dialog-window-button\" href=\"kkymsg:maximize\" onclick=\"return kkyDialogAction('maximize',event);\" title=\"").Append(Html(isKorean ? "최대화/복원" : "Maximize/restore")).Append("\">&#9633;</a>");
		}
		html.Append("<a id=\"dialogClose\" class=\"fb-dialog-window-button\" href=\"kkymsg:close\" onclick=\"return kkyDialogAction('close',event);\" title=\"").Append(Html(isKorean ? "닫기" : "Close")).Append("\">&#10005;</a></div></div>");
		html.Append("<div id=\"dialogBody\" class=\"fb-dialog-body\"><div class=\"fb-dialog-content\">").Append(parts.BodyHtml).Append("</div></div>");
		html.Append("<div id=\"dialogFooter\" class=\"fb-dialog-footer\"><table class=\"fb-dialog-footer-table\" role=\"presentation\"><tr><td class=\"fb-dialog-status-cell\"><span id=\"dialogStatus\" class=\"fb-dialog-status\">").Append(Html(hint)).Append("</span></td><td class=\"fb-dialog-actions-cell\">");
		if (showCopyDetails)
		{
			html.Append("<a id=\"dialogCopyDetails\" class=\"fb-dialog-button\" href=\"kkymsg:copy-details\" onclick=\"return kkyDialogAction('copy-details',event);\">").Append(Html(isKorean ? "내용 복사" : "Copy details")).Append("</a>");
		}
		if (showReportActions)
		{
			html.Append("<a id=\"dialogOpenFolder\" class=\"fb-dialog-button\" href=\"kkymsg:open-folder\" onclick=\"return kkyDialogAction('open-folder',event);\">").Append(Html(isKorean ? "폴더 열기" : "Open folder")).Append("</a>");
			html.Append("<a id=\"dialogCopyPath\" class=\"fb-dialog-button\" href=\"kkymsg:copy-path\" onclick=\"return kkyDialogAction('copy-path',event);\">").Append(Html(isKorean ? "경로 복사" : "Copy path")).Append("</a>");
		}
		if (!string.IsNullOrWhiteSpace(auxiliaryActionText))
		{
			html.Append("<a id=\"dialogAuxiliary\" class=\"fb-dialog-button\" href=\"kkymsg:auxiliary\" onclick=\"return kkyDialogAction('auxiliary',event);\">").Append(Html(auxiliaryActionText)).Append("</a>");
		}
		if (buttons == MessageBoxButtons.YesNo)
		{
			html.Append("<a id=\"dialogDecline\" class=\"fb-dialog-button").Append(defaultAction == "decline" ? " default" : string.Empty).Append("\" href=\"kkymsg:decline\" onclick=\"return kkyDialogAction('decline',event);\">").Append(Html(negativeText)).Append("</a>");
		}
		html.Append("<a id=\"dialogAccept\" class=\"fb-dialog-button primary").Append(defaultAction == "accept" ? " default" : string.Empty).Append("\" href=\"kkymsg:accept\" onclick=\"return kkyDialogAction('accept',event);\">").Append(Html(positiveText)).Append("</a>");
		html.Append("</td></tr></table></div></div>");
		html.Append("<script>").Append(FamilyBrowserOverflowTitleScript.Script()).Append("</script>");
		html.Append("<script>function kkyDialogAction(a,e){e=e||window.event;if(e)e.cancelBubble=true;window.location.href='kkymsg:'+a;return false;}function kkyDialogDrag(e){e=e||window.event;var s=e?(e.srcElement||e.target):null;if(s&&s.tagName&&String(s.tagName).toLowerCase()=='a')return true;if(e&&typeof e.button!='undefined'&&e.button!==1&&e.button!==0)return true;window.location.href='kkymsg:drag';return false;}document.onkeydown=function(){var e=window.event;if(!e)return true;if(e.keyCode==27)return kkyDialogAction('close',e);if(e.keyCode==13){var s=e.srcElement||e.target;var t=s&&s.tagName?String(s.tagName).toLowerCase():'';if(t!='textarea'&&t!='select')return kkyDialogAction('").Append(defaultAction).Append("',e);}return true;};window.onload=function(){var e=document.getElementById('").Append(defaultAction == "decline" ? "dialogDecline" : "dialogAccept").Append("');if(e&&e.focus)e.focus();};</script></body></html>");
		return html.ToString();
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
	{
		if (keyData == Keys.Escape)
		{
			CloseWith(_closeResult);
			return true;
		}
		if (keyData == Keys.Enter)
		{
			CloseWith(_buttons == MessageBoxButtons.YesNo && _defaultButton == MessageBoxDefaultButton.Button2 ? DialogResult.No : (_buttons == MessageBoxButtons.YesNo ? DialogResult.Yes : DialogResult.OK));
			return true;
		}
		return base.ProcessCmdKey(ref msg, keyData);
	}

	protected override void OnFormClosing(FormClosingEventArgs e)
	{
		if (DialogResult == DialogResult.None)
		{
			DialogResult = _closeResult;
		}
		base.OnFormClosing(e);
	}

	private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
	{
		string target = e.Url == null ? string.Empty : e.Url.AbsoluteUri;
		if (string.IsNullOrWhiteSpace(target) || target.StartsWith("about:blank", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}
		e.Cancel = true;
		if (!target.StartsWith("kkymsg:", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}

		string action = target.Substring("kkymsg:".Length).Trim().Trim('/');
		int delimiter = action.IndexOfAny(new char[] { '?', '#' });
		if (delimiter >= 0)
		{
			action = action.Substring(0, delimiter);
		}
		switch (action.ToLowerInvariant())
		{
			case "accept":
				CloseWith(_buttons == MessageBoxButtons.YesNo ? DialogResult.Yes : DialogResult.OK);
				break;
			case "decline":
			case "close":
				CloseWith(_closeResult);
				break;
			case "copy-details":
				CopyToClipboard(_copyDetailsText, _isKorean ? "메시지 내용을 복사했습니다." : "Message details copied.");
				break;
			case "copy-path":
				CopyToClipboard(_reportPath, _isKorean ? "리포트 경로를 복사했습니다." : "Report path copied.");
				break;
			case "open-folder":
				OpenReportFolder();
				break;
			case "maximize":
				WindowState = WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized;
				break;
			case "auxiliary":
				AuxiliaryActionRequested?.Invoke(this, EventArgs.Empty);
				break;
			case "drag":
				if (WindowState == FormWindowState.Normal)
				{
					ReleaseCapture();
					SendMessage(Handle, WmNcLButtonDown, new IntPtr(HtCaption), IntPtr.Zero);
				}
				break;
		}
	}

	private void BrowserDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
	{
		try
		{
			string id = _buttons == MessageBoxButtons.YesNo && _defaultButton == MessageBoxDefaultButton.Button2 ? "dialogDecline" : "dialogAccept";
			HtmlElement element = _browser.Document == null ? null : _browser.Document.GetElementById(id);
			if (element != null)
			{
				element.Focus();
			}
		}
		catch
		{
		}
	}

	private void CloseWith(DialogResult result)
	{
		DialogResult = result;
		Close();
	}

	private void CopyToClipboard(string value, string successText)
	{
		try
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				Clipboard.SetText(value);
				SetStatus(successText);
			}
		}
		catch (Exception ex)
		{
			SetStatus(ex.Message);
		}
	}

	private void OpenReportFolder()
	{
		try
		{
			string folder = string.IsNullOrWhiteSpace(_reportPath) ? string.Empty : Path.GetDirectoryName(_reportPath);
			if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
			{
				Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
				SetStatus(_isKorean ? "리포트 폴더를 열었습니다." : "Report folder opened.");
			}
		}
		catch (Exception ex)
		{
			SetStatus(ex.Message);
		}
	}

	public void SetStatusMessage(string text)
	{
		SetStatus(text);
	}

	private void SetStatus(string text)
	{
		try
		{
			HtmlElement element = _browser.Document == null ? null : _browser.Document.GetElementById("dialogStatus");
			if (element != null)
			{
				element.InnerText = text ?? string.Empty;
			}
		}
		catch
		{
		}
	}

	private static GeneratedDocumentParts ExtractGeneratedDocument(string document)
	{
		string source = document ?? string.Empty;
		int styleOpen = source.IndexOf("<style", StringComparison.OrdinalIgnoreCase);
		int styleOpenEnd = styleOpen < 0 ? -1 : source.IndexOf('>', styleOpen);
		int styleClose = styleOpenEnd < 0 ? -1 : source.IndexOf("</style>", styleOpenEnd + 1, StringComparison.OrdinalIgnoreCase);
		string style = styleOpenEnd >= 0 && styleClose > styleOpenEnd ? source.Substring(styleOpenEnd + 1, styleClose - styleOpenEnd - 1) : string.Empty;

		int bodyOpen = source.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
		int bodyOpenEnd = bodyOpen < 0 ? -1 : source.IndexOf('>', bodyOpen);
		int bodyClose = bodyOpenEnd < 0 ? -1 : source.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
		if (bodyOpen >= 0 && bodyOpenEnd > bodyOpen && bodyClose > bodyOpenEnd)
		{
			return new GeneratedDocumentParts(source.Substring(bodyOpen, bodyOpenEnd - bodyOpen + 1), style, source.Substring(bodyOpenEnd + 1, bodyClose - bodyOpenEnd - 1));
		}
		return new GeneratedDocumentParts("<body>", style, "<div class=\"wrap\"><pre>" + Html(source) + "</pre></div>");
	}

	private static string AddBodyAttribute(string bodyOpenTag, string name, string value)
	{
		string tag = string.IsNullOrWhiteSpace(bodyOpenTag) ? "<body>" : bodyOpenTag;
		if (tag.IndexOf(name + "=", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return tag;
		}
		int end = tag.LastIndexOf('>');
		if (end < 0)
		{
			return "<body " + name + "=\"" + Html(value) + "\">";
		}
		return tag.Insert(end, " " + name + "=\"" + Html(value) + "\"");
	}

	private static string ResolveKind(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand) return "error";
		if (icon == MessageBoxIcon.Exclamation) return "warning";
		if (icon == MessageBoxIcon.Question) return "question";
		return "information";
	}

	private static string ResolveKindText(bool isKorean, MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Exclamation) return isKorean ? "주의가 필요한 작업 결과" : "Action result needs attention";
		if (icon == MessageBoxIcon.Hand) return isKorean ? "작업 중 오류가 발생했습니다" : "An error occurred";
		if (icon == MessageBoxIcon.Question) return isKorean ? "작업을 계속할지 확인합니다" : "Confirm before continuing";
		return isKorean ? "작업 결과" : "Action result";
	}

	private static string ResolveAccent(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Exclamation) return "#d39a1b";
		if (icon == MessageBoxIcon.Hand) return "#c94d3e";
		return "#2f6bff";
	}

	private static string Html(string value)
	{
		return (value ?? string.Empty).Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&#39;");
	}

	[DllImport("user32.dll")]
	private static extern bool ReleaseCapture();

	[DllImport("user32.dll")]
	private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

	private sealed class GeneratedDocumentParts
	{
		public readonly string BodyOpenTag;
		public readonly string Style;
		public readonly string BodyHtml;

		public GeneratedDocumentParts(string bodyOpenTag, string style, string bodyHtml)
		{
			BodyOpenTag = bodyOpenTag ?? "<body>";
			Style = style ?? string.Empty;
			BodyHtml = bodyHtml ?? string.Empty;
		}
	}
}
