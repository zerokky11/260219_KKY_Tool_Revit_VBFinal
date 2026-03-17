using System;
using System.Linq;
using Drawing = System.Drawing;
using WinForms = System.Windows.Forms;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.UI
{
    public sealed class ElementParameterUpdateConfigForm : WinForms.Form
    {
        private readonly WinForms.ComboBox _combinationCombo = new WinForms.ComboBox();
        private readonly WinForms.DataGridView _conditionsGrid = new WinForms.DataGridView();
        private readonly WinForms.DataGridView _assignmentsGrid = new WinForms.DataGridView();

        public ElementParameterUpdateSettings ConfiguredSettings { get; private set; }

        public ElementParameterUpdateConfigForm(ElementParameterUpdateSettings seed)
        {
            ConfiguredSettings = seed?.Clone() ?? new ElementParameterUpdateSettings();
            InitializeForm();
            InitializeLayout();
            BindSeed();
        }

        private void InitializeForm()
        {
            Text = "모델 객체 파라미터 일괄 입력 설정";
            StartPosition = WinForms.FormStartPosition.CenterParent;
            Width = 1180;
            Height = 620;
            MinimumSize = new Drawing.Size(1080, 560);
            Font = new Drawing.Font("Segoe UI", 9F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point);
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
        }

        private void InitializeLayout()
        {
            var root = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new WinForms.Padding(12) };
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 42F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 46F));
            Controls.Add(root);

            var top = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 4, RowCount = 1 };
            top.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 140F));
            top.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 120F));
            top.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100F));
            top.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 200F));
            root.Controls.Add(top, 0, 0);

            top.Controls.Add(new WinForms.Label { Text = "조건 결합", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 0);
            _combinationCombo.Dock = WinForms.DockStyle.Fill;
            _combinationCombo.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
            _combinationCombo.Items.AddRange(Enum.GetNames(typeof(ParameterConditionCombination)));
            top.Controls.Add(_combinationCombo, 1, 0);
            top.Controls.Add(new WinForms.Label { Text = "왼쪽 필터 설정 4개 / 오른쪽 입력 파라미터 4개", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 0);
            top.Controls.Add(new WinForms.Label { Text = "조건에 맞는 모델 객체만 대상", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleRight }, 3, 0);

            var body = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            body.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 55F));
            body.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 45F));
            root.Controls.Add(body, 0, 1);

            var leftGroup = new WinForms.GroupBox { Text = "필터 설정", Dock = WinForms.DockStyle.Fill };
            var rightGroup = new WinForms.GroupBox { Text = "밀어넣을 파라미터 / 값", Dock = WinForms.DockStyle.Fill };
            body.Controls.Add(leftGroup, 0, 0);
            body.Controls.Add(rightGroup, 1, 0);

            _conditionsGrid.Dock = WinForms.DockStyle.Fill;
            _conditionsGrid.AllowUserToAddRows = false;
            _conditionsGrid.AllowUserToDeleteRows = false;
            _conditionsGrid.RowHeadersVisible = false;
            _conditionsGrid.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill;
            _conditionsGrid.Columns.Add(new WinForms.DataGridViewCheckBoxColumn { HeaderText = "사용", FillWeight = 16F });
            _conditionsGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "파라미터", FillWeight = 34F });
            var opCol = new WinForms.DataGridViewComboBoxColumn { HeaderText = "연산", FillWeight = 26F, DisplayStyle = WinForms.DataGridViewComboBoxDisplayStyle.DropDownButton };
            opCol.Items.AddRange(Enum.GetNames(typeof(FilterRuleOperator)));
            _conditionsGrid.Columns.Add(opCol);
            _conditionsGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "값", FillWeight = 24F });
            leftGroup.Controls.Add(_conditionsGrid);

            _assignmentsGrid.Dock = WinForms.DockStyle.Fill;
            _assignmentsGrid.AllowUserToAddRows = false;
            _assignmentsGrid.AllowUserToDeleteRows = false;
            _assignmentsGrid.RowHeadersVisible = false;
            _assignmentsGrid.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill;
            _assignmentsGrid.Columns.Add(new WinForms.DataGridViewCheckBoxColumn { HeaderText = "사용", FillWeight = 20F });
            _assignmentsGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "입력 파라미터", FillWeight = 40F });
            _assignmentsGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "입력 값", FillWeight = 40F });
            rightGroup.Controls.Add(_assignmentsGrid);

            for (int i = 0; i < 4; i++)
            {
                _conditionsGrid.Rows.Add(false, string.Empty, nameof(FilterRuleOperator.Equals), string.Empty);
                _assignmentsGrid.Rows.Add(false, string.Empty, string.Empty);
            }

            var bottom = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.RightToLeft, WrapContents = false };
            root.Controls.Add(bottom, 0, 2);
            var okButton = new WinForms.Button { Text = "확인", Width = 100, Height = 30 };
            var cancelButton = new WinForms.Button { Text = "취소", Width = 100, Height = 30 };
            okButton.Click += (_, __) => ConfirmAndClose();
            cancelButton.Click += (_, __) => { DialogResult = WinForms.DialogResult.Cancel; Close(); };
            bottom.Controls.Add(okButton);
            bottom.Controls.Add(cancelButton);
            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        private void BindSeed()
        {
            _combinationCombo.SelectedItem = ConfiguredSettings.CombinationMode.ToString();

            var conditions = ConfiguredSettings.Conditions.Where(x => x != null).Take(4).ToList();
            for (int i = 0; i < conditions.Count; i++)
            {
                var item = conditions[i];
                _conditionsGrid.Rows[i].Cells[0].Value = item.Enabled;
                _conditionsGrid.Rows[i].Cells[1].Value = item.ParameterName ?? string.Empty;
                _conditionsGrid.Rows[i].Cells[2].Value = item.Operator.ToString();
                _conditionsGrid.Rows[i].Cells[3].Value = item.Value ?? string.Empty;
            }

            var assignments = ConfiguredSettings.Assignments.Where(x => x != null).Take(4).ToList();
            for (int i = 0; i < assignments.Count; i++)
            {
                var item = assignments[i];
                _assignmentsGrid.Rows[i].Cells[0].Value = item.Enabled;
                _assignmentsGrid.Rows[i].Cells[1].Value = item.ParameterName ?? string.Empty;
                _assignmentsGrid.Rows[i].Cells[2].Value = item.Value ?? string.Empty;
            }
        }

        private void ConfirmAndClose()
        {
            var settings = new ElementParameterUpdateSettings
            {
                CombinationMode = (ParameterConditionCombination)Enum.Parse(typeof(ParameterConditionCombination), Convert.ToString(_combinationCombo.SelectedItem) ?? nameof(ParameterConditionCombination.And))
            };

            foreach (WinForms.DataGridViewRow row in _conditionsGrid.Rows)
            {
                settings.Conditions.Add(new ElementParameterCondition
                {
                    Enabled = ConvertToBool(row.Cells[0].Value),
                    ParameterName = Convert.ToString(row.Cells[1].Value),
                    Operator = (FilterRuleOperator)Enum.Parse(typeof(FilterRuleOperator), Convert.ToString(row.Cells[2].Value) ?? nameof(FilterRuleOperator.Equals)),
                    Value = Convert.ToString(row.Cells[3].Value) ?? string.Empty
                });
            }

            foreach (WinForms.DataGridViewRow row in _assignmentsGrid.Rows)
            {
                settings.Assignments.Add(new ElementParameterAssignment
                {
                    Enabled = ConvertToBool(row.Cells[0].Value),
                    ParameterName = Convert.ToString(row.Cells[1].Value),
                    Value = Convert.ToString(row.Cells[2].Value) ?? string.Empty
                });
            }

            ConfiguredSettings = settings;
            DialogResult = WinForms.DialogResult.OK;
            Close();
        }

        private static bool ConvertToBool(object value)
        {
            if (value == null) return false;
            bool result;
            return bool.TryParse(value.ToString(), out result) && result;
        }
    }
}
