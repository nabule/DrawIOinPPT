using System;
using System.Drawing;
using System.Windows.Forms;
using DrawioPpt.Core;
using DrawioPpt.Core.Models;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.PowerPointAddIn.UI
{
    public class SettingsForm : Form
    {
        private const string AppTitle = "DrawIO&PPT";
        private readonly ComboBox _editorModeComboBox;
        private readonly TextBox _desktopPathTextBox;
        private readonly TextBox _editorUrlTextBox;
        private readonly CheckBox _officeCompatibleLabelsCheckBox;
        private readonly CheckBox _autoOpenCheckBox;
        private readonly CheckBox _autoUpdateCheckBox;
        private readonly CheckBox _diagramInfoDialogCheckBox;
        private readonly CheckBox _keepSidecarCheckBox;
        private readonly TextBox _sidecarFolderTextBox;
        private readonly Label _versionLabel;
        private readonly Button _browseButton;
        private readonly Button _detectButton;
        private readonly Button _testUrlButton;
        private readonly Button _okButton;
        private readonly Button _cancelButton;
        private readonly DesktopEditorPathDetector _pathDetector;

        public SettingsForm(PluginSettings settings, DesktopEditorPathDetector pathDetector)
            : this(settings, pathDetector, true)
        {
        }

        public SettingsForm(
            PluginSettings settings,
            DesktopEditorPathDetector pathDetector,
            bool showAutoOpenOnSelection)
        {
            _pathDetector = pathDetector;
            this.Settings = Clone(settings);

            this.Text = AppTitle;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.ClientSize = new Size(700, 388);

            Label modeLabel = CreateLabel("编辑器模式", 16, 20);
            _editorModeComboBox = new ComboBox();
            _editorModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _editorModeComboBox.Items.Add(new EditorModeOption(EditorMode.Desktop, "桌面版"));
            _editorModeComboBox.Items.Add(new EditorModeOption(EditorMode.Url, "网页地址"));
            _editorModeComboBox.Location = new Point(150, 16);
            _editorModeComboBox.Width = 160;

            Label desktopLabel = CreateLabel("桌面版路径", 16, 58);
            _desktopPathTextBox = CreateTextBox(150, 54, 330);
            _browseButton = new Button();
            _browseButton.Text = "浏览...";
            _browseButton.AutoSize = true;
            _browseButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            _browseButton.Click += OnBrowseDesktopPath;

            _detectButton = new Button();
            _detectButton.Text = "自动检测";
            _detectButton.AutoSize = true;
            _detectButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            _detectButton.Click += OnDetectDesktopPath;

            FlowLayoutPanel desktopButtonsPanel = CreateButtonPanel(492, 50);
            desktopButtonsPanel.Controls.Add(_browseButton);
            desktopButtonsPanel.Controls.Add(_detectButton);

            Label urlLabel = CreateLabel("编辑器地址", 16, 98);
            _editorUrlTextBox = CreateTextBox(150, 94, 430);

            _testUrlButton = new Button();
            _testUrlButton.Text = "测试地址";
            _testUrlButton.AutoSize = true;
            _testUrlButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            _testUrlButton.Click += OnTestUrl;

            FlowLayoutPanel urlButtonsPanel = CreateButtonPanel(588, 90);
            urlButtonsPanel.Controls.Add(_testUrlButton);

            _officeCompatibleLabelsCheckBox = CreateCheckBox("使用兼容 Microsoft Office 的 SVG 文本标签", 150, 128);
            _autoOpenCheckBox = CreateCheckBox("选中图形时自动打开编辑器", 150, 156);
            _autoOpenCheckBox.Visible = showAutoOpenOnSelection;
            _autoUpdateCheckBox = CreateCheckBox("保存后自动刷新图形", 150, 184);
            _diagramInfoDialogCheckBox = CreateCheckBox("新建或编辑时显示图形信息弹窗", 150, 212);
            _keepSidecarCheckBox = CreateCheckBox("在 PPT 附近保留伴随工作文件", 150, 240);

            Label sidecarLabel = CreateLabel("伴随文件夹", 16, 276);
            _sidecarFolderTextBox = CreateTextBox(150, 272, 510);

            _okButton = new Button();
            _okButton.Text = "确定";
            _okButton.Size = new Size(80, 30);
            _okButton.DialogResult = DialogResult.OK;
            _okButton.Click += OnOk;

            _cancelButton = new Button();
            _cancelButton.Text = "取消";
            _cancelButton.Size = new Size(80, 30);
            _cancelButton.DialogResult = DialogResult.Cancel;

            FlowLayoutPanel bottomButtonsPanel = new FlowLayoutPanel();
            bottomButtonsPanel.FlowDirection = FlowDirection.RightToLeft;
            bottomButtonsPanel.WrapContents = false;
            bottomButtonsPanel.AutoSize = false;
            bottomButtonsPanel.Size = new Size(176, 34);
            bottomButtonsPanel.Location = new Point(484, 330);
            bottomButtonsPanel.Controls.Add(_cancelButton);
            bottomButtonsPanel.Controls.Add(_okButton);

            _versionLabel = CreateLabel("版本：" + BuildInfo.DisplayVersion, 16, 354);
            _versionLabel.Width = 440;

            this.Controls.Add(modeLabel);
            this.Controls.Add(_editorModeComboBox);
            this.Controls.Add(desktopLabel);
            this.Controls.Add(_desktopPathTextBox);
            this.Controls.Add(desktopButtonsPanel);
            this.Controls.Add(urlLabel);
            this.Controls.Add(_editorUrlTextBox);
            this.Controls.Add(urlButtonsPanel);
            this.Controls.Add(_officeCompatibleLabelsCheckBox);
            this.Controls.Add(_autoOpenCheckBox);
            this.Controls.Add(_autoUpdateCheckBox);
            this.Controls.Add(_diagramInfoDialogCheckBox);
            this.Controls.Add(_keepSidecarCheckBox);
            this.Controls.Add(sidecarLabel);
            this.Controls.Add(_sidecarFolderTextBox);
            this.Controls.Add(_versionLabel);
            this.Controls.Add(bottomButtonsPanel);

            this.AcceptButton = _okButton;
            this.CancelButton = _cancelButton;

            LoadSettings(this.Settings);
        }

        public PluginSettings Settings { get; private set; }

        private static CheckBox CreateCheckBox(string text, int left, int top)
        {
            CheckBox checkBox = new CheckBox();
            checkBox.Text = text;
            checkBox.Location = new Point(left, top);
            checkBox.AutoSize = true;
            return checkBox;
        }

        private static Label CreateLabel(string text, int left, int top)
        {
            Label label = new Label();
            label.Text = text;
            label.Location = new Point(left, top + 4);
            label.Width = 125;
            return label;
        }

        private static TextBox CreateTextBox(int left, int top, int width)
        {
            TextBox textBox = new TextBox();
            textBox.Location = new Point(left, top);
            textBox.Width = width;
            return textBox;
        }

        private static FlowLayoutPanel CreateButtonPanel(int left, int top)
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.Location = new Point(left, top);
            panel.AutoSize = true;
            panel.WrapContents = false;
            panel.FlowDirection = FlowDirection.LeftToRight;
            return panel;
        }

        private static PluginSettings Clone(PluginSettings settings)
        {
            PluginSettings source = settings ?? new PluginSettings();
            PluginSettings clone = new PluginSettings();
            clone.EditorMode = source.EditorMode;
            clone.DesktopEditorPath = source.DesktopEditorPath;
            clone.EditorUrl = source.EditorUrl;
            clone.UseOfficeCompatibleSvgLabels = source.UseOfficeCompatibleSvgLabels;
            clone.AutoOpenOnSelection = source.AutoOpenOnSelection;
            clone.AutoUpdateOnSave = source.AutoUpdateOnSave;
            clone.ShowDiagramInfoDialog = source.ShowDiagramInfoDialog;
            clone.KeepSidecarFile = source.KeepSidecarFile;
            clone.SidecarFolderName = source.SidecarFolderName;
            return clone;
        }

        private void LoadSettings(PluginSettings settings)
        {
            foreach (object item in _editorModeComboBox.Items)
            {
                EditorModeOption option = item as EditorModeOption;
                if (option != null && option.Value == settings.EditorMode)
                {
                    _editorModeComboBox.SelectedItem = option;
                    break;
                }
            }

            _desktopPathTextBox.Text = settings.DesktopEditorPath ?? string.Empty;
            _editorUrlTextBox.Text = settings.EditorUrl ?? string.Empty;
            _officeCompatibleLabelsCheckBox.Checked = settings.UseOfficeCompatibleSvgLabels;
            _autoOpenCheckBox.Checked = settings.AutoOpenOnSelection;
            _autoUpdateCheckBox.Checked = settings.AutoUpdateOnSave;
            _diagramInfoDialogCheckBox.Checked = settings.ShowDiagramInfoDialog;
            _keepSidecarCheckBox.Checked = settings.KeepSidecarFile;
            _sidecarFolderTextBox.Text = settings.SidecarFolderName ?? string.Empty;
        }

        private void OnBrowseDesktopPath(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Title = "选择 draw.io 或 diagrams.net 的可执行文件";

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _desktopPathTextBox.Text = dialog.FileName;
                }
            }
        }

        private void OnDetectDesktopPath(object sender, EventArgs e)
        {
            if (_pathDetector == null)
            {
                return;
            }

            string detectedPath = _pathDetector.Detect();
            if (string.IsNullOrWhiteSpace(detectedPath))
            {
                MessageBox.Show(this, "未检测到已安装的 draw.io/diagrams.net 桌面版。", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _desktopPathTextBox.Text = detectedPath;
        }

        private void OnOk(object sender, EventArgs e)
        {
            object selectedItem = _editorModeComboBox.SelectedItem;
            if (selectedItem == null)
            {
                MessageBox.Show(this, "请选择编辑器模式。", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }

            this.Settings = BuildSettingsFromInputs();
        }

        private void OnTestUrl(object sender, EventArgs e)
        {
            PluginSettings draftSettings = BuildSettingsFromInputs();
            if (string.IsNullOrWhiteSpace(draftSettings.EditorUrl))
            {
                MessageBox.Show(this, "请先填写可测试的编辑器地址。", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PluginTraceLog traceLog = new PluginTraceLog();
            DiagramEnvelope envelope = CreateUrlTestEnvelope(draftSettings.EditorUrl);
            bool saved = false;
            int svgLength = 0;

            using (UrlDiagramEditorForm form = new UrlDiagramEditorForm(draftSettings.EditorUrl, envelope.DiagramName, envelope.DrawioXml, draftSettings.UseOfficeCompatibleSvgLabels, traceLog))
            {
                form.DiagramSaved += delegate(object testSender, UrlDiagramSavedEventArgs args)
                {
                    saved = true;
                    svgLength = args.SvgMarkup == null ? 0 : args.SvgMarkup.Length;
                };

                DialogResult result = form.ShowDialog(this);
                if (saved)
                {
                    MessageBox.Show(
                        this,
                        "URL 模式测试成功。" +
                        Environment.NewLine + "SVG 长度: " + svgLength +
                        Environment.NewLine + "日志: " + traceLog.LogPath,
                        AppTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                string detail = result == DialogResult.Cancel
                    ? "编辑器已关闭，但没有检测到保存回调。"
                    : "没有检测到有效的保存或导出回调。";

                MessageBox.Show(
                    this,
                    detail + Environment.NewLine + "请检查编辑器地址和日志文件：" + Environment.NewLine + traceLog.LogPath,
                    AppTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private PluginSettings BuildSettingsFromInputs()
        {
            PluginSettings updated = new PluginSettings();
            EditorModeOption selectedOption = _editorModeComboBox.SelectedItem as EditorModeOption;
            updated.EditorMode = selectedOption == null ? EditorMode.Desktop : selectedOption.Value;
            updated.DesktopEditorPath = _desktopPathTextBox.Text;
            updated.EditorUrl = _editorUrlTextBox.Text;
            updated.UseOfficeCompatibleSvgLabels = _officeCompatibleLabelsCheckBox.Checked;
            updated.AutoOpenOnSelection = _autoOpenCheckBox.Checked;
            updated.AutoUpdateOnSave = _autoUpdateCheckBox.Checked;
            updated.ShowDiagramInfoDialog = _diagramInfoDialogCheckBox.Checked;
            updated.KeepSidecarFile = _keepSidecarCheckBox.Checked;
            updated.SidecarFolderName = _sidecarFolderTextBox.Text;
            return updated;
        }

        private static DiagramEnvelope CreateUrlTestEnvelope(string editorUrl)
        {
            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "DrawIO&PPT 地址测试";
            envelope.EditorMode = EditorMode.Url;
            envelope.EditorTarget = editorUrl ?? string.Empty;
            envelope.UpdatedUtc = DateTime.UtcNow;
            envelope.DrawioXml =
                "<mxfile host=\"" + AppTitle + "\">" +
                "<diagram id=\"" + envelope.DiagramId + "\" name=\"DrawIO&PPT 地址测试\">" +
                "<mxGraphModel dx=\"1000\" dy=\"1000\" grid=\"1\" gridSize=\"10\" guides=\"1\" tooltips=\"1\" connect=\"1\" arrows=\"1\" fold=\"1\" page=\"1\" pageScale=\"1\" pageWidth=\"850\" pageHeight=\"1100\" math=\"0\" shadow=\"0\">" +
                "<root>" +
                "<mxCell id=\"0\"/>" +
                "<mxCell id=\"1\" parent=\"0\"/>" +
                "<mxCell id=\"2\" value=\"地址测试\" style=\"rounded=1;whiteSpace=wrap;html=1;\" vertex=\"1\" parent=\"1\">" +
                "<mxGeometry x=\"120\" y=\"120\" width=\"140\" height=\"60\" as=\"geometry\"/>" +
                "</mxCell>" +
                "</root>" +
                "</mxGraphModel>" +
                "</diagram>" +
                "</mxfile>";
            return envelope;
        }

        private class EditorModeOption
        {
            public EditorModeOption(EditorMode value, string text)
            {
                this.Value = value;
                this.Text = text;
            }

            public EditorMode Value { get; private set; }

            public string Text { get; private set; }

            public override string ToString()
            {
                return this.Text;
            }
        }
    }
}
