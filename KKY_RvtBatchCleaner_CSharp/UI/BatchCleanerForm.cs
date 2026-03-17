using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using UIApplication = Autodesk.Revit.UI.UIApplication;
using Drawing = System.Drawing;
using WinForms = System.Windows.Forms;
using KKY_Tool_Revit.Models;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit.UI
{
    public sealed class BatchCleanerForm : WinForms.Form
    {
        private readonly UIApplication _uiapp;

        private readonly WinForms.ListBox _filesList = new WinForms.ListBox();
        private readonly WinForms.TextBox _outputFolderText = new WinForms.TextBox();
        private readonly WinForms.TextBox _viewNameText = new WinForms.TextBox();
        private readonly WinForms.DataGridView _viewParamGrid = new WinForms.DataGridView();

        private readonly WinForms.CheckBox _useElementUpdateCheck = new WinForms.CheckBox();
        private readonly WinForms.Button _configureElementUpdateButton = new WinForms.Button();
        private readonly WinForms.TextBox _elementUpdateSummaryText = new WinForms.TextBox();
        private ElementParameterUpdateSettings _elementUpdateSettings = new ElementParameterUpdateSettings();

        private readonly WinForms.CheckBox _useFilterCheck = new WinForms.CheckBox();
        private readonly WinForms.CheckBox _applyFilterInitiallyCheck = new WinForms.CheckBox();
        private readonly WinForms.CheckBox _autoEnableFilterIfEmptyCheck = new WinForms.CheckBox();
        private readonly WinForms.TextBox _filterNameText = new WinForms.TextBox();
        private readonly WinForms.TextBox _filterCategoriesText = new WinForms.TextBox();
        private readonly WinForms.TextBox _filterParameterText = new WinForms.TextBox();
        private readonly WinForms.ComboBox _filterOperatorCombo = new WinForms.ComboBox();
        private readonly WinForms.TextBox _filterValueText = new WinForms.TextBox();
        private readonly WinForms.TextBox _filterStructureText = new WinForms.TextBox();

        private ViewFilterProfile _boundFilterProfile;

        private readonly WinForms.TextBox _logText = new WinForms.TextBox();
        private readonly WinForms.Button _prepareButton = new WinForms.Button();
        private readonly WinForms.Button _purgeButton = new WinForms.Button();
        private readonly WinForms.Button _saveButton = new WinForms.Button();
        private readonly WinForms.Button _reviewButton = new WinForms.Button();
        private readonly WinForms.Button _extractButton = new WinForms.Button();
        private readonly WinForms.Button _openFolderButton = new WinForms.Button();

        private BatchPrepareSession _currentSession;
        private bool _closingForPurge;
        private string _lastExtractParameterNamesCsv = string.Empty;

        public BatchCleanerForm(UIApplication uiapp)
        {
            _uiapp = uiapp ?? throw new ArgumentNullException(nameof(uiapp));

            InitializeForm();
            InitializeLayout();
            InitializeEvents();
            RestoreSharedState();
        }

        private void InitializeForm()
        {
            Text = "KKY RVT Batch Cleaner";
            StartPosition = WinForms.FormStartPosition.CenterScreen;
            Width = 1280;
            Height = 940;
            MinimumSize = new Drawing.Size(1200, 840);
            Font = new Drawing.Font("Segoe UI", 9F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point);
            _applyFilterInitiallyCheck.Checked = true;
        }

        private void InitializeLayout()
        {
            var root = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new WinForms.Padding(12)
            };
            root.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 430F));
            root.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100F));
            Controls.Add(root);

            var leftPanel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, RowCount = 5, ColumnCount = 1 };
            leftPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 260F));
            leftPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 86F));
            leftPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 120F));
            leftPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 280F));
            leftPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            root.Controls.Add(leftPanel, 0, 0);

            var rightPanel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, RowCount = 4, ColumnCount = 1 };
            rightPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 260F));
            rightPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 180F));
            rightPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 290F));
            rightPanel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            root.Controls.Add(rightPanel, 1, 0);

            leftPanel.Controls.Add(BuildFilesGroup(), 0, 0);
            leftPanel.Controls.Add(BuildBasicOptionsGroup(), 0, 1);
            leftPanel.Controls.Add(BuildRunGroup(), 0, 2);
            leftPanel.Controls.Add(BuildHelpGroup(), 0, 3);
            leftPanel.Controls.Add(new WinForms.Panel { Dock = WinForms.DockStyle.Fill }, 0, 4);

            rightPanel.Controls.Add(BuildViewParametersGroup(), 0, 0);
            rightPanel.Controls.Add(BuildElementUpdateGroup(), 0, 1);
            rightPanel.Controls.Add(BuildFilterGroup(), 0, 2);
            rightPanel.Controls.Add(BuildLogGroup(), 0, 3);
        }

        private WinForms.Control BuildFilesGroup()
        {
            var group = new WinForms.GroupBox { Text = "1) 대상 RVT 파일", Dock = WinForms.DockStyle.Fill };
            var panel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 42F));
            group.Controls.Add(panel);

            _filesList.Dock = WinForms.DockStyle.Fill;
            _filesList.HorizontalScrollbar = true;
            panel.Controls.Add(_filesList, 0, 0);

            var buttons = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.LeftToRight, WrapContents = false };
            var addButton = new WinForms.Button { Text = "RVT 추가...", Width = 100, Height = 30 };
            var removeButton = new WinForms.Button { Text = "선택 제거", Width = 100, Height = 30 };
            var clearButton = new WinForms.Button { Text = "목록 비우기", Width = 100, Height = 30 };
            addButton.Click += (_, __) => AddFiles();
            removeButton.Click += (_, __) => RemoveSelectedFiles();
            clearButton.Click += (_, __) => { _filesList.Items.Clear(); UpdateActionButtons(); };
            buttons.Controls.Add(addButton);
            buttons.Controls.Add(removeButton);
            buttons.Controls.Add(clearButton);
            panel.Controls.Add(buttons, 0, 1);

            return group;
        }

        private WinForms.Control BuildBasicOptionsGroup()
        {
            var group = new WinForms.GroupBox { Text = "2) 기본 설정", Dock = WinForms.DockStyle.Fill };
            var panel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 3, RowCount = 2 };
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 120F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 140F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32F));
            group.Controls.Add(panel);

            panel.Controls.Add(new WinForms.Label { Text = "결과 폴더", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 0);
            _outputFolderText.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_outputFolderText, 1, 0);
            var browseOutputButton = new WinForms.Button { Text = "찾아보기...", Dock = WinForms.DockStyle.Fill };
            browseOutputButton.Click += (_, __) => BrowseOutputFolder();
            panel.Controls.Add(browseOutputButton, 2, 0);

            panel.Controls.Add(new WinForms.Label { Text = "3D 뷰 이름", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 1);
            _viewNameText.Dock = WinForms.DockStyle.Fill;
            _viewNameText.Text = "KKY_CLEAN_3D";
            panel.Controls.Add(_viewNameText, 1, 1);
            panel.Controls.Add(new WinForms.Label { Text = "기존 3D 뷰 삭제 후 이 이름 1개만 남김", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 1);

            return group;
        }

        private WinForms.Control BuildRunGroup()
        {
            var group = new WinForms.GroupBox { Text = "3) 실행", Dock = WinForms.DockStyle.Fill };
            var panel = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.LeftToRight, WrapContents = true };
            group.Controls.Add(panel);

            _prepareButton.Text = "정리 시작";
            _prepareButton.Width = 110;
            _prepareButton.Height = 34;
            panel.Controls.Add(_prepareButton);

            _purgeButton.Text = "Purge 일괄처리";
            _purgeButton.Width = 120;
            _purgeButton.Height = 34;
            _purgeButton.Enabled = false;
            panel.Controls.Add(_purgeButton);

            _reviewButton.Text = "검토";
            _reviewButton.Width = 90;
            _reviewButton.Height = 34;
            _reviewButton.Enabled = false;
            panel.Controls.Add(_reviewButton);

            _extractButton.Text = "속성값 추출";
            _extractButton.Width = 110;
            _extractButton.Height = 34;
            _extractButton.Enabled = false;
            panel.Controls.Add(_extractButton);

            _saveButton.Text = "저장 없음";
            _saveButton.Visible = false;
            _saveButton.Enabled = false;

            _openFolderButton.Text = "결과 폴더 열기";
            _openFolderButton.Width = 120;
            _openFolderButton.Height = 34;
            _openFolderButton.Enabled = false;
            panel.Controls.Add(_openFolderButton);

            var closeButton = new WinForms.Button { Text = "닫기", Width = 100, Height = 34 };
            closeButton.Click += (_, __) => Close();
            panel.Controls.Add(closeButton);

            return group;
        }

        private WinForms.Control BuildHelpGroup()
        {
            var group = new WinForms.GroupBox { Text = "동작 요약", Dock = WinForms.DockStyle.Fill };
            var text = new WinForms.TextBox
            {
                Dock = WinForms.DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = WinForms.ScrollBars.Vertical,
                Text = string.Join(Environment.NewLine, new[]
                {
                    "- 1단계: 선택 RVT를 하나씩 열어 정리 후 즉시 저장",
                    "- 2단계: 저장된 정리본을 하나씩 다시 열어 Purge를 반복 실행",
                    "- 3단계: Purge 저장 후 백업파일 삭제, 파일 닫기, 다음 파일 진행",
                    "- Manage Links/Imports/Image/PointCloud 제거",
                    "- 그룹 인스턴스 해제/삭제 + GroupType 삭제, 어셈블리 해제/삭제",
                    "- 기존 3D 뷰/동명 뷰 삭제 후 지정 이름으로 새 3D 뷰 1개 생성",
                    "- VG는 기본적으로 전부 켠 뒤 Lines/Mass/Parts/Site 및 End/Cut 서브카테고리 끄기",
                    "- 모델 객체 파라미터 일괄 입력은 별도 설정창에서 다중 조건/다중 입력 설정",
                    "- 정리된 파일의 모델 객체 속성값을 CSV로 추출 가능",
                    "- 뷰 필터 XML 불러오기/문서 필터 추출 가능",
                    "- 파일명의 _Detached 제거 후 결과 폴더에 저장"
                })
            };
            group.Controls.Add(text);
            return group;
        }

        private WinForms.Control BuildViewParametersGroup()
        {
            var group = new WinForms.GroupBox { Text = "4) 뷰 파라미터 지정 (최대 5개)", Dock = WinForms.DockStyle.Fill };
            _viewParamGrid.Dock = WinForms.DockStyle.Fill;
            _viewParamGrid.AllowUserToAddRows = false;
            _viewParamGrid.AllowUserToDeleteRows = false;
            _viewParamGrid.RowHeadersVisible = false;
            _viewParamGrid.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill;
            _viewParamGrid.Columns.Add(new WinForms.DataGridViewCheckBoxColumn { HeaderText = "사용", Width = 50 });
            _viewParamGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "파라미터명" });
            _viewParamGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn { HeaderText = "값" });

            for (int i = 0; i < 5; i++)
            {
                _viewParamGrid.Rows.Add(false, string.Empty, string.Empty);
            }

            group.Controls.Add(_viewParamGrid);
            return group;
        }

        private WinForms.Control BuildElementUpdateGroup()
        {
            var group = new WinForms.GroupBox { Text = "5) 모델 객체 파라미터 일괄 입력", Dock = WinForms.DockStyle.Fill };
            var panel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 3, RowCount = 3 };
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 120F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 140F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 34F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 34F));
            panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
            group.Controls.Add(panel);

            _useElementUpdateCheck.Text = "기능 사용";
            _useElementUpdateCheck.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_useElementUpdateCheck, 0, 0);

            _configureElementUpdateButton.Text = "필터 설정...";
            _configureElementUpdateButton.Dock = WinForms.DockStyle.Fill;
            _configureElementUpdateButton.Click += (_, __) => ConfigureElementUpdate();
            panel.Controls.Add(_configureElementUpdateButton, 1, 0);
            panel.Controls.Add(new WinForms.Label { Text = "왼쪽은 조건 4개, 오른쪽은 입력 파라미터/값 4개까지", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 0);

            panel.Controls.Add(new WinForms.Label { Text = "설정 요약", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 1);
            panel.Controls.Add(new WinForms.Label { Text = string.Empty, Dock = WinForms.DockStyle.Fill }, 1, 1);
            panel.Controls.Add(new WinForms.Label { Text = "조건 결합(AND/OR) + 다중 파라미터 입력 지원", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 1);

            _elementUpdateSummaryText.Dock = WinForms.DockStyle.Fill;
            _elementUpdateSummaryText.Multiline = true;
            _elementUpdateSummaryText.ReadOnly = true;
            _elementUpdateSummaryText.ScrollBars = WinForms.ScrollBars.Vertical;
            panel.Controls.Add(_elementUpdateSummaryText, 0, 2);
            panel.SetColumnSpan(_elementUpdateSummaryText, 3);
            UpdateElementUpdateSummary();

            return group;
        }

        private WinForms.Control BuildFilterGroup()
        {
            var group = new WinForms.GroupBox { Text = "6) 뷰 필터", Dock = WinForms.DockStyle.Fill };
            var panel = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 3, RowCount = 9 };
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 130F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100F));
            panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 150F));
            for (int i = 0; i < 9; i++) panel.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32F));
            group.Controls.Add(panel);

            _useFilterCheck.Text = "필터 사용";
            _useFilterCheck.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_useFilterCheck, 0, 0);
            panel.Controls.Add(new WinForms.Label { Text = string.Empty }, 1, 0);
            panel.Controls.Add(new WinForms.Label { Text = string.Empty }, 2, 0);

            _applyFilterInitiallyCheck.Text = "처음부터 적용";
            _applyFilterInitiallyCheck.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_applyFilterInitiallyCheck, 0, 1);
            _autoEnableFilterIfEmptyCheck.Text = "초기 미적용 + 뷰 비면 자동 적용";
            _autoEnableFilterIfEmptyCheck.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_autoEnableFilterIfEmptyCheck, 1, 1);

            var importXmlButton = new WinForms.Button { Text = "XML 불러오기...", Dock = WinForms.DockStyle.Fill };
            importXmlButton.Click += (_, __) => ImportFilterXml();
            panel.Controls.Add(importXmlButton, 2, 1);

            panel.Controls.Add(new WinForms.Label { Text = "필터 이름", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 2);
            _filterNameText.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_filterNameText, 1, 2);
            var exportXmlButton = new WinForms.Button { Text = "입력값 XML 저장...", Dock = WinForms.DockStyle.Fill };
            exportXmlButton.Click += (_, __) => ExportFilterXml();
            panel.Controls.Add(exportXmlButton, 2, 2);

            panel.Controls.Add(new WinForms.Label { Text = "문서 필터", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 3);
            panel.Controls.Add(new WinForms.Label { Text = "현재 열린 문서의 Parameter Filter를 읽어서 XML로 추출", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 1, 3);
            var exportExistingButton = new WinForms.Button { Text = "문서 필터 추출...", Dock = WinForms.DockStyle.Fill };
            exportExistingButton.Click += (_, __) => ExportExistingDocumentFilterXml();
            panel.Controls.Add(exportExistingButton, 2, 3);

            panel.Controls.Add(new WinForms.Label { Text = "카테고리", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 4);
            _filterCategoriesText.Dock = WinForms.DockStyle.Fill;
            _filterCategoriesText.Text = "OST_PipeFitting, OST_DuctFitting, OST_ConduitFitting, OST_CableTrayFitting";
            panel.Controls.Add(_filterCategoriesText, 1, 4);
            panel.Controls.Add(new WinForms.Label { Text = "쉼표 구분", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 4);

            panel.Controls.Add(new WinForms.Label { Text = "파라미터", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 5);
            _filterParameterText.Dock = WinForms.DockStyle.Fill;
            _filterParameterText.Text = "BIP:ALL_MODEL_INSTANCE_COMMENTS";
            panel.Controls.Add(_filterParameterText, 1, 5);
            panel.Controls.Add(new WinForms.Label { Text = "BIP:... 또는 이름/GUID", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 5);

            panel.Controls.Add(new WinForms.Label { Text = "연산", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 6);
            _filterOperatorCombo.Dock = WinForms.DockStyle.Left;
            _filterOperatorCombo.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
            _filterOperatorCombo.Items.AddRange(Enum.GetNames(typeof(FilterRuleOperator)));
            _filterOperatorCombo.SelectedIndex = 0;
            panel.Controls.Add(_filterOperatorCombo, 1, 6);
            panel.Controls.Add(new WinForms.Label { Text = string.Empty }, 2, 6);

            panel.Controls.Add(new WinForms.Label { Text = "값", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 7);
            _filterValueText.Dock = WinForms.DockStyle.Fill;
            panel.Controls.Add(_filterValueText, 1, 7);
            panel.Controls.Add(new WinForms.Label { Text = "단일 규칙 입력용", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 7);

            panel.Controls.Add(new WinForms.Label { Text = "구조 요약", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 8);
            _filterStructureText.Dock = WinForms.DockStyle.Fill;
            _filterStructureText.ReadOnly = true;
            panel.Controls.Add(_filterStructureText, 1, 8);
            panel.Controls.Add(new WinForms.Label { Text = "XML 불러오기/문서 추출 시 유지", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 2, 8);

            return group;
        }

        private WinForms.Control BuildLogGroup()
        {
            var group = new WinForms.GroupBox { Text = "로그", Dock = WinForms.DockStyle.Fill };
            _logText.Dock = WinForms.DockStyle.Fill;
            _logText.Multiline = true;
            _logText.ScrollBars = WinForms.ScrollBars.Both;
            _logText.WordWrap = false;
            _logText.ReadOnly = true;
            group.Controls.Add(_logText);
            return group;
        }

        private void InitializeEvents()
        {
            _prepareButton.Click += (_, __) => PrepareBatch();
            _purgeButton.Click += (_, __) => RunExperimentalPurgeBatch();
            _reviewButton.Click += (_, __) => RunVerification();
            _extractButton.Click += (_, __) => RunModelParameterExtraction();
            _saveButton.Click += (_, __) => SavePreparedBatch();
            _openFolderButton.Click += (_, __) => OpenOutputFolder();
            _useElementUpdateCheck.CheckedChanged += (_, __) =>
            {
                UpdateElementUpdateSummary();
                CacheCurrentSettingsSnapshot();
            };
            _useFilterCheck.CheckedChanged += (_, __) => CacheCurrentSettingsSnapshot();
            _applyFilterInitiallyCheck.CheckedChanged += (_, __) => CacheCurrentSettingsSnapshot();
            _autoEnableFilterIfEmptyCheck.CheckedChanged += (_, __) => CacheCurrentSettingsSnapshot();
            _outputFolderText.TextChanged += (_, __) =>
            {
                CacheCurrentSettingsSnapshot();
                UpdateActionButtons();
            };
        }

        private void AddFiles()
        {
            using (var dialog = new WinForms.OpenFileDialog())
            {
                dialog.Filter = "Revit Project (*.rvt)|*.rvt";
                dialog.Multiselect = true;
                dialog.Title = "정리할 RVT 파일 선택";

                if (dialog.ShowDialog(this) != WinForms.DialogResult.OK) return;

                var existing = _filesList.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (string file in dialog.FileNames)
                {
                    if (existing.Add(file))
                    {
                        _filesList.Items.Add(file);
                    }
                }

                UpdateActionButtons();
            }
        }

        private void RemoveSelectedFiles()
        {
            var selected = _filesList.SelectedItems.Cast<object>().ToList();
            foreach (object item in selected)
            {
                _filesList.Items.Remove(item);
            }

            UpdateActionButtons();
        }

        private void BrowseOutputFolder()
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "정리된 RVT 저장 폴더 선택";
                if (!string.IsNullOrWhiteSpace(_outputFolderText.Text) && Directory.Exists(_outputFolderText.Text))
                {
                    dialog.SelectedPath = _outputFolderText.Text;
                }

                if (dialog.ShowDialog(this) == WinForms.DialogResult.OK)
                {
                    _outputFolderText.Text = dialog.SelectedPath;
                    _openFolderButton.Enabled = true;
                    UpdateActionButtons();
                }
            }
        }

        private void ImportFilterXml()
        {
            using (var dialog = new WinForms.OpenFileDialog())
            {
                dialog.Filter = "XML (*.xml)|*.xml";
                dialog.Title = "뷰 필터 XML 불러오기";
                if (dialog.ShowDialog(this) != WinForms.DialogResult.OK) return;

                try
                {
                    ViewFilterProfile profile = RevitViewFilterProfileService.LoadFromXml(dialog.FileName);
                    BindFilterProfile(profile);
                    AppendLog("필터 XML 불러오기 완료: " + dialog.FileName);
                }
                catch (Exception ex)
                {
                    WinForms.MessageBox.Show(this, ex.Message, "불러오기 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                }
            }
        }

        private void ExportFilterXml()
        {
            try
            {
                ViewFilterProfile profile = ReadFilterProfile();
                using (var dialog = new WinForms.SaveFileDialog())
                {
                    dialog.Filter = "XML (*.xml)|*.xml";
                    dialog.Title = "뷰 필터 XML 저장";
                    dialog.FileName = string.IsNullOrWhiteSpace(profile.FilterName) ? "ViewFilterProfile.xml" : profile.FilterName + ".xml";

                    if (dialog.ShowDialog(this) != WinForms.DialogResult.OK) return;

                    RevitViewFilterProfileService.SaveToXml(profile, dialog.FileName);
                    AppendLog("입력값 기준 필터 XML 저장 완료: " + dialog.FileName);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(this, ex.Message, "저장 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void ExportExistingDocumentFilterXml()
        {
            Document doc = _uiapp.ActiveUIDocument?.Document;
            if (doc == null)
            {
                WinForms.MessageBox.Show(this, "현재 열린 문서가 없습니다.", "문서 없음", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            List<ParameterFilterElement> filters = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .OrderBy(x => x.Name)
                .ToList();

            if (filters.Count == 0)
            {
                WinForms.MessageBox.Show(this, "현재 문서에 Parameter Filter가 없습니다.", "필터 없음", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            ParameterFilterElement selectedFilter = PromptForExistingFilter(filters);
            if (selectedFilter == null) return;

            try
            {
                ViewFilterProfile profile = RevitViewFilterProfileService.ExtractProfileFromFilter(doc, selectedFilter.Id);
                BindFilterProfile(profile);

                using (var dialog = new WinForms.SaveFileDialog())
                {
                    dialog.Filter = "XML (*.xml)|*.xml";
                    dialog.Title = "현재 문서 필터 XML 추출";
                    dialog.FileName = string.IsNullOrWhiteSpace(profile.FilterName) ? "ViewFilterProfile.xml" : profile.FilterName + ".xml";
                    if (dialog.ShowDialog(this) != WinForms.DialogResult.OK) return;

                    RevitViewFilterProfileService.SaveToXml(profile, dialog.FileName);
                    AppendLog("현재 문서 필터 XML 추출 완료: " + dialog.FileName);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(this, ex.Message, "필터 추출 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private static ParameterFilterElement PromptForExistingFilter(IList<ParameterFilterElement> filters)
        {
            using (var dialog = new WinForms.Form())
            {
                dialog.Text = "추출할 Parameter Filter 선택";
                dialog.StartPosition = WinForms.FormStartPosition.CenterParent;
                dialog.Width = 520;
                dialog.Height = 420;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;

                var root = new WinForms.TableLayoutPanel { Dock = WinForms.DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new WinForms.Padding(12) };
                root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 30F));
                root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 100F));
                root.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 42F));
                dialog.Controls.Add(root);

                root.Controls.Add(new WinForms.Label { Text = "현재 문서에서 XML로 추출할 필터를 선택", Dock = WinForms.DockStyle.Fill, TextAlign = Drawing.ContentAlignment.MiddleLeft }, 0, 0);

                var list = new WinForms.ListBox { Dock = WinForms.DockStyle.Fill, DisplayMember = "Name" };
                foreach (ParameterFilterElement filter in filters)
                {
                    list.Items.Add(filter);
                }
                if (list.Items.Count > 0) list.SelectedIndex = 0;
                root.Controls.Add(list, 0, 1);

                var buttons = new WinForms.FlowLayoutPanel { Dock = WinForms.DockStyle.Fill, FlowDirection = WinForms.FlowDirection.RightToLeft, WrapContents = false };
                var okButton = new WinForms.Button { Text = "확인", Width = 90, Height = 30, DialogResult = WinForms.DialogResult.OK };
                var cancelButton = new WinForms.Button { Text = "취소", Width = 90, Height = 30, DialogResult = WinForms.DialogResult.Cancel };
                buttons.Controls.Add(okButton);
                buttons.Controls.Add(cancelButton);
                root.Controls.Add(buttons, 0, 2);

                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;

                return dialog.ShowDialog() == WinForms.DialogResult.OK ? list.SelectedItem as ParameterFilterElement : null;
            }
        }

        private void BindFilterProfile(ViewFilterProfile profile)
        {
            if (profile == null) return;

            _boundFilterProfile = profile.Clone();
            _filterNameText.Text = profile.FilterName ?? string.Empty;
            _filterCategoriesText.Text = profile.CategoriesCsv ?? string.Empty;
            _filterParameterText.Text = profile.ParameterToken ?? string.Empty;
            _filterOperatorCombo.SelectedItem = profile.Operator.ToString();
            _filterValueText.Text = profile.RuleValue ?? string.Empty;
            _filterStructureText.Text = profile.StructureSummary ?? string.Empty;
        }

        private ViewFilterProfile ReadFilterProfile()
        {
            var op = (FilterRuleOperator)Enum.Parse(typeof(FilterRuleOperator), _filterOperatorCombo.SelectedItem?.ToString() ?? nameof(FilterRuleOperator.Equals));
            var profile = new ViewFilterProfile
            {
                FilterName = _filterNameText.Text?.Trim(),
                CategoriesCsv = _filterCategoriesText.Text?.Trim(),
                ParameterToken = _filterParameterText.Text?.Trim(),
                Operator = op,
                RuleValue = _filterValueText.Text ?? string.Empty,
                StructureSummary = _filterStructureText.Text ?? string.Empty
            };

            if (_boundFilterProfile != null && ShouldPreserveSerializedDefinition(profile, _boundFilterProfile))
            {
                profile.FilterDefinitionXml = _boundFilterProfile.FilterDefinitionXml;
                if (string.IsNullOrWhiteSpace(profile.StructureSummary))
                {
                    profile.StructureSummary = _boundFilterProfile.StructureSummary;
                }
            }

            return profile;
        }

        private static bool ShouldPreserveSerializedDefinition(ViewFilterProfile current, ViewFilterProfile bound)
        {
            if (current == null || bound == null) return false;
            if (string.IsNullOrWhiteSpace(bound.FilterDefinitionXml)) return false;

            return string.Equals(current.ParameterToken ?? string.Empty, bound.ParameterToken ?? string.Empty, StringComparison.Ordinal)
                   && current.Operator == bound.Operator
                   && string.Equals(current.RuleValue ?? string.Empty, bound.RuleValue ?? string.Empty, StringComparison.Ordinal)
                   && string.Equals(current.StructureSummary ?? string.Empty, bound.StructureSummary ?? string.Empty, StringComparison.Ordinal);
        }

        private List<ViewParameterAssignment> ReadViewParameters()
        {
            var list = new List<ViewParameterAssignment>();
            foreach (WinForms.DataGridViewRow row in _viewParamGrid.Rows)
            {
                list.Add(new ViewParameterAssignment
                {
                    Enabled = ConvertToBool(row.Cells[0].Value),
                    ParameterName = Convert.ToString(row.Cells[1].Value),
                    ParameterValue = Convert.ToString(row.Cells[2].Value)
                });
            }
            return list;
        }

        private static bool ConvertToBool(object value)
        {
            if (value == null) return false;
            bool result;
            if (bool.TryParse(value.ToString(), out result)) return result;
            return false;
        }


        private ElementParameterUpdateSettings ReadElementUpdateSettings()
        {
            ElementParameterUpdateSettings copy = _elementUpdateSettings?.Clone() ?? new ElementParameterUpdateSettings();
            copy.Enabled = _useElementUpdateCheck.Checked;
            return copy;
        }

        private void ConfigureElementUpdate()
        {
            using (var dialog = new ElementParameterUpdateConfigForm(ReadElementUpdateSettings()))
            {
                if (dialog.ShowDialog(this) != WinForms.DialogResult.OK)
                {
                    return;
                }

                _elementUpdateSettings = dialog.ConfiguredSettings?.Clone() ?? new ElementParameterUpdateSettings();
                _elementUpdateSettings.Enabled = _useElementUpdateCheck.Checked;
                UpdateElementUpdateSummary();
            }
        }

        private void UpdateElementUpdateSummary()
        {
            _elementUpdateSummaryText.Text = ReadElementUpdateSettings().BuildSummary();
        }

        private BatchCleanSettings BuildSettings()
        {
            return new BatchCleanSettings
            {
                FilePaths = _filesList.Items.Cast<string>().ToList(),
                OutputFolder = _outputFolderText.Text?.Trim(),
                Target3DViewName = _viewNameText.Text?.Trim(),
                ViewParameters = ReadViewParameters(),
                ElementParameterUpdate = ReadElementUpdateSettings(),
                UseFilter = _useFilterCheck.Checked,
                ApplyFilterInitially = _applyFilterInitiallyCheck.Checked,
                AutoEnableFilterIfEmpty = _autoEnableFilterIfEmptyCheck.Checked,
                FilterProfile = ReadFilterProfile()
            };
        }

        private void ApplySettings(BatchCleanSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            _filesList.Items.Clear();
            if (settings.FilePaths != null)
            {
                foreach (string file in settings.FilePaths.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    _filesList.Items.Add(file);
                }
            }

            _outputFolderText.Text = settings.OutputFolder ?? string.Empty;
            _viewNameText.Text = string.IsNullOrWhiteSpace(settings.Target3DViewName) ? "KKY_CLEAN_3D" : settings.Target3DViewName;

            if (settings.ViewParameters != null)
            {
                for (int i = 0; i < _viewParamGrid.Rows.Count && i < settings.ViewParameters.Count; i++)
                {
                    ViewParameterAssignment assignment = settings.ViewParameters[i] ?? new ViewParameterAssignment();
                    _viewParamGrid.Rows[i].Cells[0].Value = assignment.Enabled;
                    _viewParamGrid.Rows[i].Cells[1].Value = assignment.ParameterName ?? string.Empty;
                    _viewParamGrid.Rows[i].Cells[2].Value = assignment.ParameterValue ?? string.Empty;
                }
            }

            _useFilterCheck.Checked = settings.UseFilter;
            _applyFilterInitiallyCheck.Checked = settings.ApplyFilterInitially;
            _autoEnableFilterIfEmptyCheck.Checked = settings.AutoEnableFilterIfEmpty;
            _boundFilterProfile = settings.FilterProfile != null ? settings.FilterProfile.Clone() : null;
            if (_boundFilterProfile != null)
            {
                BindFilterProfile(_boundFilterProfile);
            }
            else
            {
                _filterNameText.Text = string.Empty;
                _filterCategoriesText.Text = string.Empty;
                _filterParameterText.Text = string.Empty;
                _filterOperatorCombo.SelectedIndex = 0;
                _filterValueText.Text = string.Empty;
                _filterStructureText.Text = string.Empty;
            }

            _elementUpdateSettings = settings.ElementParameterUpdate != null ? settings.ElementParameterUpdate.Clone() : new ElementParameterUpdateSettings();
            _useElementUpdateCheck.Checked = _elementUpdateSettings.Enabled;
            UpdateElementUpdateSummary();
            _openFolderButton.Enabled = !string.IsNullOrWhiteSpace(_outputFolderText.Text) && Directory.Exists(_outputFolderText.Text);
            UpdateActionButtons();
        }

        private void CacheCurrentSettingsSnapshot()
        {
            try
            {
                App.SharedLastSettings = BuildSettings();
            }
            catch
            {
            }
        }

        private bool ValidateSettings(BatchCleanSettings settings)
        {
            if (settings.FilePaths.Count == 0)
            {
                WinForms.MessageBox.Show(this, "RVT 파일을 하나 이상 추가해야 합니다.", "입력 필요", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(settings.OutputFolder) || !Directory.Exists(settings.OutputFolder))
            {
                WinForms.MessageBox.Show(this, "정리 결과 저장 폴더를 지정해야 합니다.", "입력 필요", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return false;
            }
            if (settings.UseFilter && !settings.FilterProfile.IsConfigured())
            {
                WinForms.MessageBox.Show(this, "필터 사용이 켜져 있으면 필터 이름/카테고리/파라미터/값을 모두 지정해야 합니다.", "필터 설정 필요", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return false;
            }
            if (settings.ElementParameterUpdate != null && settings.ElementParameterUpdate.Enabled && !settings.ElementParameterUpdate.IsConfigured())
            {
                WinForms.MessageBox.Show(this, "객체 파라미터 일괄 입력을 사용하려면 조건 파라미터/입력 파라미터를 지정해야 합니다.", "객체 파라미터 설정 필요", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private bool EnsureNoPendingPreparedDocuments()
        {
            if (PurgeUiBatchService.IsRunning)
            {
                WinForms.MessageBox.Show(this, "현재 Purge 일괄처리가 실행 중입니다. 완료 후 다시 시도해 주세요.", "Purge 실행 중", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return false;
            }

            if (_currentSession == null)
            {
                return true;
            }

            if ((_currentSession.CleanedOutputPaths == null || _currentSession.CleanedOutputPaths.Count == 0) && (_currentSession.PreparedDocuments == null || _currentSession.PreparedDocuments.Count == 0))
            {
                return true;
            }

            WinForms.DialogResult dialogResult = WinForms.MessageBox.Show(this,
                "이전 정리 결과 목록이 있습니다. 새로 시작하면 기존 목록이 지워집니다. 계속할까요?",
                "이전 결과 있음",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Question);

            if (dialogResult != WinForms.DialogResult.Yes)
            {
                return false;
            }

            _currentSession = null;
            App.SharedPreparedSession = null;
            _saveButton.Enabled = false;
            UpdateActionButtons();
            return true;
        }

        private void PrepareBatch()
        {
            if (!EnsureNoPendingPreparedDocuments())
            {
                return;
            }

            BatchCleanSettings settings = BuildSettings();
            App.SharedLastSettings = settings;
            if (!ValidateSettings(settings))
            {
                return;
            }

            _prepareButton.Enabled = false;
            _saveButton.Enabled = false;
            _reviewButton.Enabled = false;
            _purgeButton.Enabled = false;
            AppendLog("정리 시작");

            try
            {
                _currentSession = BatchCleanService.CleanAndSave(_uiapp, settings, AppendLog);
                int successCount = _currentSession.Results.Count(x => x.Success);
                int failCount = _currentSession.Results.Count - successCount;
                int savedCount = _currentSession.CleanedOutputPaths != null ? _currentSession.CleanedOutputPaths.Count : 0;

                AppendLog($"정리 종료 / 성공 {successCount} / 실패 {failCount} / 저장완료 {savedCount}");
                App.SharedPreparedSession = _currentSession;
                _openFolderButton.Enabled = Directory.Exists(settings.OutputFolder);
                _saveButton.Enabled = false;
                UpdateActionButtons();

                WinForms.MessageBox.Show(this,
                    "정리 완료" + Environment.NewLine +
                    "성공: " + successCount + Environment.NewLine +
                    "실패: " + failCount + Environment.NewLine +
                    "저장완료: " + savedCount,
                    "정리 완료",
                    WinForms.MessageBoxButtons.OK,
                    failCount == 0 ? WinForms.MessageBoxIcon.Information : WinForms.MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                AppendLog("정리 중 오류: " + ex.Message);
                WinForms.MessageBox.Show(this, ex.Message, "정리 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
            finally
            {
                _prepareButton.Enabled = true;
                UpdateActionButtons();
            }
        }

        private void RunVerification()
        {
            List<string> targetPaths = ResolveEffectiveTargetPaths();
            if (targetPaths.Count == 0)
            {
                WinForms.MessageBox.Show(this, "검토할 파일이 없습니다. 목록에 RVT를 추가하거나 정리를 먼저 실행해 주세요.", "검토", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            BatchCleanSettings settings = BuildSettings();
            App.SharedLastSettings = settings;

            try
            {
                string csvPath = VerificationService.VerifyPaths(_uiapp, targetPaths, ResolveResultFolderOrFallback(targetPaths), settings, AppendLog);
                EnsureSessionForResolvedTargets(targetPaths).VerificationCsvPath = csvPath;
                App.SharedPreparedSession = _currentSession;
                WinForms.MessageBox.Show(this, "검토가 끝났습니다." + "\r\n" + "리포트(UTF-8 CSV): " + csvPath, "검토 완료", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("검토 중 오류: " + ex.Message);
                WinForms.MessageBox.Show(this, ex.Message, "검토 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void RunModelParameterExtraction()
        {
            List<string> targetPaths = ResolveEffectiveTargetPaths();
            if (targetPaths.Count == 0)
            {
                WinForms.MessageBox.Show(this, "추출할 파일이 없습니다. 목록에 RVT를 추가하거나 정리를 먼저 실행해 주세요.", "속성값 추출", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new ModelParameterExtractionForm(_lastExtractParameterNamesCsv))
            {
                if (dialog.ShowDialog(this) != WinForms.DialogResult.OK)
                {
                    return;
                }

                _lastExtractParameterNamesCsv = dialog.ParameterNamesCsv ?? string.Empty;
                try
                {
                    string csvPath = ModelParameterExtractionService.ExportModelParameters(_uiapp, targetPaths, ResolveResultFolderOrFallback(targetPaths), _lastExtractParameterNamesCsv, AppendLog);
                    WinForms.MessageBox.Show(this, "속성값 추출이 끝났습니다." + Environment.NewLine + "리포트(UTF-8 CSV): " + csvPath, "속성값 추출 완료", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    AppendLog("속성값 추출 중 오류: " + ex.Message);
                    WinForms.MessageBox.Show(this, ex.Message, "속성값 추출 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                }
            }
        }

        private void SavePreparedBatch()
        {
            WinForms.MessageBox.Show(this, "현재 프로세스에서는 정리 단계와 Purge 단계에서 각각 자동 저장됩니다. 별도 저장 단계는 사용하지 않습니다.", "저장 없음", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        protected override void OnFormClosing(WinForms.FormClosingEventArgs e)
        {
            CacheCurrentSettingsSnapshot();

            if (_closingForPurge)
            {
                base.OnFormClosing(e);
                return;
            }

            if (PurgeUiBatchService.IsRunning)
            {
                base.OnFormClosing(e);
                return;
            }

            if (_currentSession != null && ((_currentSession.CleanedOutputPaths != null && _currentSession.CleanedOutputPaths.Count > 0) || (_currentSession.PreparedDocuments != null && _currentSession.PreparedDocuments.Count > 0)))
            {
                WinForms.DialogResult dialogResult = WinForms.MessageBox.Show(this,
                    "현재 세션의 정리 결과 목록이 있습니다. 창을 닫으면 목록만 해제되고 저장된 파일은 그대로 남습니다. 닫을까요?",
                    "닫기 확인",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Question);

                if (dialogResult != WinForms.DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }

                _currentSession = null;
                App.SharedPreparedSession = null;
                PurgeProgressWindowHost.CloseWindow();
                UpdateActionButtons();
            }

            base.OnFormClosing(e);
        }

        private void RunExperimentalPurgeBatch()
        {
            List<string> targetPaths = ResolveEffectiveTargetPaths();
            if (targetPaths.Count == 0)
            {
                WinForms.MessageBox.Show(this, "Purge 대상 파일이 없습니다. 목록에 RVT를 추가하거나 정리를 먼저 실행해 주세요.", "Purge", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            if (PurgeUiBatchService.IsRunning)
            {
                WinForms.MessageBox.Show(this, "이미 Purge 일괄처리가 실행 중입니다.", "Purge", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            if (targetPaths.Any(x => string.IsNullOrWhiteSpace(x) || !File.Exists(x)))
            {
                WinForms.MessageBox.Show(this, "정리 결과 파일 중 누락된 파일이 있습니다. 정리를 다시 실행해 주세요.", "Purge", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            WinForms.DialogResult dialogResult = WinForms.MessageBox.Show(this,
                "Purge 일괄처리를 시작하면 저장된 정리 결과 파일을 하나씩 다시 열어 Purge를 반복 실행합니다.\r\n각 파일은 Purge 후 자동 저장되고 백업파일 삭제 후 닫힙니다.\r\n퍼지창이 실제로 나타날 수 있으니 작업 중에는 Revit을 건드리지 않는 편이 안전합니다. 퍼지 중에는 Revit을 항상 위로 유지하며 자동 진행을 시도합니다.\r\n\r\n계속할까요?",
                "Purge 일괄처리",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Question);

            if (dialogResult != WinForms.DialogResult.Yes)
            {
                return;
            }

            BatchPrepareSession effectiveSession = EnsureSessionForResolvedTargets(targetPaths);
            App.SharedPreparedSession = effectiveSession;
            CacheCurrentSettingsSnapshot();
            PurgeProgressWindowHost.ShowWindow();
            AppendLog("Purge 일괄처리 요청 - 저장된 정리 결과 파일을 하나씩 다시 열어 순차 Purge를 실행합니다.");

            bool started = false;
            try
            {
                started = PurgeUiBatchService.Start(_uiapp, effectiveSession, 5, AppendLog);
            }
            catch (Exception ex)
            {
                PurgeProgressWindowHost.CloseWindow();
                WinForms.MessageBox.Show(this, ex.Message, "Purge 시작 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                return;
            }

            if (!started)
            {
                PurgeProgressWindowHost.CloseWindow();
                WinForms.MessageBox.Show(this, "Purge 시작에 실패했습니다. 로그를 확인해 주세요.", "Purge 시작 실패", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                return;
            }

            _closingForPurge = true;
            Close();
        }

        private void RestoreSharedState()
        {
            BatchCleanSettings sharedSettings = App.SharedLastSettings;
            if (sharedSettings != null)
            {
                ApplySettings(sharedSettings);
            }

            IList<string> sharedLogs = App.GetSharedLogLinesSnapshot();
            if (sharedLogs != null && sharedLogs.Count > 0)
            {
                foreach (string line in sharedLogs)
                {
                    AppendLogCore(line);
                }
            }

            BatchPrepareSession sharedSession = App.SharedPreparedSession;
            if (sharedSession == null)
            {
                UpdateActionButtons();
                return;
            }

            sharedSession.PreparedDocuments = sharedSession.PreparedDocuments
                .Where(x => x != null && x.Document != null && x.Document.IsValidObject)
                .ToList();

            _currentSession = (sharedSession.CleanedOutputPaths != null && sharedSession.CleanedOutputPaths.Count > 0) ? sharedSession : null;
            _saveButton.Enabled = false;
            UpdateActionButtons();

            if (_currentSession != null)
            {
                AppendLog("이전 정리 세션을 복원했습니다. Purge 대상 파일: " + _currentSession.CleanedOutputPaths.Count);
                if (!string.IsNullOrWhiteSpace(_currentSession.OutputFolder))
                {
                    _outputFolderText.Text = _currentSession.OutputFolder;
                }
            }
        }

        private List<string> ResolveEffectiveTargetPaths()
        {
            List<string> listPaths = _filesList.Items.Cast<string>()
                .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (listPaths.Count > 0)
            {
                return listPaths;
            }

            return (_currentSession != null ? (_currentSession.CleanedOutputPaths ?? new List<string>()) : new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string ResolveResultFolderOrFallback(IList<string> targetPaths)
        {
            string path = _outputFolderText.Text != null ? _outputFolderText.Text.Trim() : string.Empty;
            if (!string.IsNullOrWhiteSpace(path))
            {
                Directory.CreateDirectory(path);
                return path;
            }

            if (targetPaths != null && targetPaths.Count > 0)
            {
                string fallback = Path.GetDirectoryName(targetPaths[0]);
                if (!string.IsNullOrWhiteSpace(fallback))
                {
                    Directory.CreateDirectory(fallback);
                    return fallback;
                }
            }

            return string.Empty;
        }

        private BatchPrepareSession EnsureSessionForResolvedTargets(IList<string> targetPaths)
        {
            List<string> normalizedTargets = targetPaths != null ? targetPaths.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList() : new List<string>();
            List<string> currentPaths = _currentSession != null ? (_currentSession.CleanedOutputPaths ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList() : new List<string>();
            bool sameTargets = normalizedTargets.Count == currentPaths.Count && !normalizedTargets.Except(currentPaths, StringComparer.OrdinalIgnoreCase).Any();

            if (_currentSession != null && sameTargets)
            {
                if (string.IsNullOrWhiteSpace(_currentSession.OutputFolder))
                {
                    _currentSession.OutputFolder = ResolveResultFolderOrFallback(normalizedTargets);
                }
                return _currentSession;
            }

            _currentSession = new BatchPrepareSession
            {
                OutputFolder = ResolveResultFolderOrFallback(normalizedTargets),
                CleanedOutputPaths = normalizedTargets
            };
            UpdateActionButtons();
            return _currentSession;
        }

        private void UpdateActionButtons()
        {
            bool hasSessionTargets = _currentSession != null && (_currentSession.CleanedOutputPaths ?? new List<string>()).Any(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x));
            bool hasListTargets = _filesList.Items.Cast<string>().Any(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x));
            bool hasTargets = hasSessionTargets || hasListTargets;
            _reviewButton.Enabled = hasTargets;
            _extractButton.Enabled = hasTargets;
            _purgeButton.Enabled = hasTargets && !PurgeUiBatchService.IsRunning;
            _openFolderButton.Enabled = !string.IsNullOrWhiteSpace(_outputFolderText.Text) && Directory.Exists(_outputFolderText.Text);
        }

        private void OpenOutputFolder()
        {
            string path = _outputFolderText.Text?.Trim();
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }

        private void AppendLog(string message)
        {
            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message;
            App.AppendSharedLog(line);
            AppendLogCore(line);
        }

        private void AppendLogCore(string line)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(AppendLogCore), line);
                return;
            }

            _logText.AppendText(line + Environment.NewLine);
            _logText.SelectionStart = _logText.TextLength;
            _logText.ScrollToCaret();
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
