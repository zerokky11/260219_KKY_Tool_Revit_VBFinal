using System;
using Drawing = System.Drawing;
using WinForms = System.Windows.Forms;

namespace KKY_Tool_Revit.UI
{
    public sealed class ModelParameterExtractionForm : WinForms.Form
    {
        private readonly WinForms.TextBox _parameterNamesText = new WinForms.TextBox();
        public string ParameterNamesCsv { get; private set; }

        public ModelParameterExtractionForm(string initialValue)
        {
            Text = "모델 속성값 추출 설정";
            StartPosition = WinForms.FormStartPosition.CenterParent;
            Width = 620;
            Height = 360;
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            BuildLayout(initialValue);
        }

        private void BuildLayout(string initialValue)
        {
            var root = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new WinForms.Padding(12) };
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 52F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 58F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 42F));
            Controls.Add(root);

            root.Controls.Add(new WinForms.Label
            {
                Text = "쉼표(,)로 구분해서 추출할 파라미터 이름을 입력\r\n예: Comments, Mark, KKY_CODE",
                Dock = WinForms.DockStyle.Fill,
                TextAlign = Drawing.ContentAlignment.MiddleLeft
            }, 0, 0);

            _parameterNamesText.Dock = WinForms.DockStyle.Fill;
            _parameterNamesText.Multiline = true;
            _parameterNamesText.ScrollBars = WinForms.ScrollBars.Vertical;
            _parameterNamesText.Text = initialValue ?? string.Empty;
            root.Controls.Add(_parameterNamesText, 0, 1);

            root.Controls.Add(new WinForms.Label
            {
                Text = "추출 컬럼: 파일명, 요소ID, 카테고리, 패밀리이름, 타입이름, 입력한 파라미터들",
                Dock = WinForms.DockStyle.Fill,
                TextAlign = Drawing.ContentAlignment.MiddleLeft
            }, 0, 2);

            var buttons = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.RightToLeft, WrapContents = false };
            var okButton = new WinForms.Button { Text = "실행", Width = 100, Height = 30 };
            var cancelButton = new WinForms.Button { Text = "취소", Width = 100, Height = 30 };
            okButton.Click += (_, __) => OnOk();
            cancelButton.Click += (_, __) => { DialogResult = WinForms.DialogResult.Cancel; Close(); };
            buttons.Controls.Add(okButton);
            buttons.Controls.Add(cancelButton);
            root.Controls.Add(buttons, 0, 3);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        private void OnOk()
        {
            string value = _parameterNamesText.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(value))
            {
                WinForms.MessageBox.Show(this, "추출할 파라미터 이름을 하나 이상 입력해야 합니다.", "입력 필요", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            ParameterNamesCsv = value.Trim();
            DialogResult = WinForms.DialogResult.OK;
            Close();
        }
    }
}
