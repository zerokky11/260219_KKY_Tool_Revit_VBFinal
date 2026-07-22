using System;
using System.Drawing;
using System.Windows.Forms;

internal static class FamilyBrowserModernMessageDialog
{
	public static DialogResult Show(IWin32Window owner, bool isKorean, string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
	{
		return Show(owner, isKorean, message, caption, buttons, icon, MessageBoxDefaultButton.Button1);
	}

	public static DialogResult Show(IWin32Window owner, bool isKorean, string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
	{
		return Show(owner, isKorean, message, caption, buttons, icon, defaultButton, string.Empty, string.Empty);
	}

	public static DialogResult Show(IWin32Window owner, bool isKorean, string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, string positiveButtonText, string negativeButtonText)
	{
		try
		{
			if (buttons == MessageBoxButtons.OK || buttons == MessageBoxButtons.YesNo)
			{
				Size preferredSize = ResolvePreferredSize(isKorean, message);
				string reportPath = FamilyBrowserMessageHtmlRenderer.FindPrimaryOutputPath(message);
				using (FamilyBrowserHtmlDialogHost dialog = new FamilyBrowserHtmlDialogHost(
					isKorean,
					caption,
					FamilyBrowserMessageHtmlRenderer.Build(isKorean, message, caption, icon, FamilyBrowserUiTheme.Light),
					buttons,
					icon,
					defaultButton,
					positiveButtonText,
					negativeButtonText,
					message,
					reportPath,
					preferredSize,
					new Size(680, 430),
					false))
				{
					return dialog.ShowDialog(owner);
				}
			}
		}
		catch
		{
		}
		return MessageBox.Show(owner, message, caption, buttons, icon, defaultButton);
	}

	public static string BuildHtmlForAudit(bool isKorean, string message, string caption, MessageBoxIcon icon)
	{
		return BuildDialogHtmlForAudit(isKorean, message, caption, icon, FamilyBrowserUiTheme.Light);
	}

	public static string BuildHtmlForThemeAudit(bool isKorean, string message, string caption, MessageBoxIcon icon, string themeCode)
	{
		return BuildDialogHtmlForAudit(isKorean, message, caption, icon, FamilyBrowserUiTheme.Light);
	}

	private static string BuildDialogHtmlForAudit(bool isKorean, string message, string caption, MessageBoxIcon icon, FamilyBrowserUiTheme theme)
	{
		bool showReportActions = !string.IsNullOrWhiteSpace(FamilyBrowserMessageHtmlRenderer.FindPrimaryOutputPath(message));
		return FamilyBrowserHtmlDialogHost.BuildDocument(
			isKorean,
			caption,
			FamilyBrowserMessageHtmlRenderer.Build(isKorean, message, caption, icon, theme),
			MessageBoxButtons.OK,
			icon,
			MessageBoxDefaultButton.Button1,
			string.Empty,
			string.Empty,
			true,
			showReportActions,
			false);
	}

	private static Size ResolvePreferredSize(bool isKorean, string message)
	{
		int messageLength = (message ?? string.Empty).Length;
		bool structured = FamilyBrowserMessageHtmlRenderer.ContainsStructuredSections(message);
		int width = isKorean ? 860 : 820;
		int height = structured ? (messageLength > 1200 ? 720 : 640) : (messageLength > 700 ? 590 : (messageLength > 260 ? 510 : 440));
		return new Size(width, height);
	}
}
