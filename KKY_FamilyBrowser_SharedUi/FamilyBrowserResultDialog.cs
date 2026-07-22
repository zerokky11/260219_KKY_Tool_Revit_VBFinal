using System.Windows.Forms;

internal static class FamilyBrowserResultDialog
{
	public static DialogResult Show(string caption, string message)
	{
		return Show(null, caption, message, MessageBoxIcon.Asterisk);
	}

	public static DialogResult Show(IWin32Window owner, string caption, string message, MessageBoxIcon icon)
	{
		return FamilyBrowserModernMessageDialog.Show(
			owner,
			FamilyBrowserLanguageService.IsKorean(),
			message,
			caption,
			MessageBoxButtons.OK,
			icon,
			MessageBoxDefaultButton.Button1);
	}

	public static bool Confirm(IWin32Window owner, string caption, string headline, string details, string positiveText, string negativeText, bool defaultToPositive)
	{
		string message = (headline ?? string.Empty).Trim();
		if (!string.IsNullOrWhiteSpace(details))
		{
			message += "\r\n\r\n" + details.Trim();
		}
		DialogResult result = FamilyBrowserModernMessageDialog.Show(
			owner,
			FamilyBrowserLanguageService.IsKorean(),
			message,
			caption,
			MessageBoxButtons.YesNo,
			MessageBoxIcon.Question,
			defaultToPositive ? MessageBoxDefaultButton.Button1 : MessageBoxDefaultButton.Button2,
			positiveText,
			negativeText);
		return result == DialogResult.Yes;
	}
}
