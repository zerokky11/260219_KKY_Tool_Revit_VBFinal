using System.Drawing;
using System.Windows.Forms;

public static class FamilyBrowserWinFormsTheme
{
	public static void Apply(Form form, FamilyBrowserUiTheme theme)
	{
		if (form == null)
		{
			return;
		}
		Palette palette = new Palette(theme);
		form.BackColor = palette.Window;
		form.ForeColor = palette.Text;
		ApplyControlTree(form, form, palette, false);
	}

	private static void ApplyControlTree(Form owner, Control root, Palette palette, bool parentBrandHeader)
	{
		foreach (Control control in root.Controls)
		{
			if (control is WebBrowser)
			{
				continue;
			}
			bool brandHeader = parentBrandHeader || IsLegacyBrandHeader(control);
			bool accentStrip = control is Panel && IsLegacyGreenAccent(control.BackColor);
			control.ForeColor = brandHeader ? palette.OnBrandHeader : palette.Text;
			if (control is Button)
			{
				Button button = (Button)control;
				if (brandHeader)
				{
					button.FlatStyle = FlatStyle.Flat;
					button.FlatAppearance.BorderSize = 0;
					button.BackColor = palette.BrandHeader;
					button.ForeColor = palette.OnBrandHeader;
					button.UseVisualStyleBackColor = false;
					ApplyControlTree(owner, control, palette, true);
					continue;
				}
				bool primary = ReferenceEquals(owner.AcceptButton, button) || button.DialogResult == DialogResult.OK || button.DialogResult == DialogResult.Yes;
				button.FlatStyle = FlatStyle.Flat;
				button.FlatAppearance.BorderSize = 1;
				button.FlatAppearance.BorderColor = primary ? palette.Accent : palette.Border;
				button.BackColor = primary ? palette.Accent : palette.Control;
				button.ForeColor = primary ? palette.OnAccent : palette.Text;
				button.UseVisualStyleBackColor = false;
			}
			else if (control is TextBoxBase)
			{
				control.BackColor = palette.Control;
			}
			else if (control is ComboBox || control is ListBox || control is ListView || control is TreeView || control is CheckedListBox || control is NumericUpDown)
			{
				control.BackColor = palette.Control;
			}
			else if (control is DataGridView)
			{
				ApplyGrid((DataGridView)control, palette);
			}
			else if (accentStrip)
			{
				control.BackColor = palette.Accent;
				control.ForeColor = palette.OnAccent;
			}
			else if (control is Panel || control is TableLayoutPanel || control is FlowLayoutPanel || control is GroupBox || control is TabControl || control is TabPage)
			{
				control.BackColor = brandHeader ? palette.BrandHeader : palette.Surface;
			}
			ApplyControlTree(owner, control, palette, brandHeader && !accentStrip);
		}
	}

	private static bool IsLegacyBrandHeader(Control control)
	{
		if (!(control is Panel || control is TableLayoutPanel || control is FlowLayoutPanel))
		{
			return false;
		}
		Color color = control.BackColor;
		return color.A > 0 && color.R <= 60 && color.G <= 85 && color.B <= 80;
	}

	private static bool IsLegacyGreenAccent(Color color)
	{
		return color.A > 0 && color.G >= 105 && color.G >= color.R + 28 && color.G >= color.B + 18;
	}

	private static void ApplyGrid(DataGridView grid, Palette palette)
	{
		grid.BackgroundColor = palette.Surface;
		grid.BorderStyle = BorderStyle.FixedSingle;
		grid.GridColor = palette.Border;
		grid.EnableHeadersVisualStyles = false;
		grid.ColumnHeadersDefaultCellStyle.BackColor = palette.Header;
		grid.ColumnHeadersDefaultCellStyle.ForeColor = palette.Text;
		grid.DefaultCellStyle.BackColor = palette.Control;
		grid.DefaultCellStyle.ForeColor = palette.Text;
		grid.DefaultCellStyle.SelectionBackColor = palette.Selection;
		grid.DefaultCellStyle.SelectionForeColor = palette.Text;
	}

	private sealed class Palette
	{
		public readonly Color Window;
		public readonly Color Surface;
		public readonly Color Control;
		public readonly Color Header;
		public readonly Color Selection;
		public readonly Color Border;
		public readonly Color Text;
		public readonly Color Accent;
		public readonly Color OnAccent;
		public readonly Color BrandHeader;
		public readonly Color OnBrandHeader;

		public Palette(FamilyBrowserUiTheme theme)
		{
			bool dark = theme == FamilyBrowserUiTheme.Dark;
			Window = dark ? Color.FromArgb(11, 18, 32) : Color.FromArgb(245, 247, 251);
			Surface = dark ? Color.FromArgb(15, 26, 43) : Color.White;
			Control = dark ? Color.FromArgb(18, 28, 51) : Color.White;
			Header = dark ? Color.FromArgb(23, 36, 59) : Color.FromArgb(238, 244, 255);
			Selection = dark ? Color.FromArgb(24, 39, 70) : Color.FromArgb(234, 242, 255);
			Border = dark ? Color.FromArgb(64, 84, 119) : Color.FromArgb(193, 204, 223);
			Text = dark ? Color.FromArgb(244, 247, 251) : Color.FromArgb(17, 24, 39);
			Accent = dark ? Color.FromArgb(47, 183, 255) : Color.FromArgb(47, 107, 255);
			OnAccent = dark ? Color.FromArgb(6, 18, 31) : Color.White;
			BrandHeader = Color.FromArgb(15, 26, 43);
			OnBrandHeader = Color.White;
		}
	}
}
