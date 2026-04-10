using System;
using System.Drawing;
using System.Windows.Forms;
using DrawioPpt.Core.Models;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.PowerPointAddIn.UI
{
    public class SettingsForm : Form
    {
        private readonly ComboBox _editorModeComboBox;
        private readonly TextBox _desktopPathTextBox;
        private readonly TextBox _editorUrlTextBox;
        private readonly CheckBox _autoOpenCheckBox;
        private readonly CheckBox _autoUpdateCheckBox;
        private readonly CheckBox _keepSidecarCheckBox;
        private readonly TextBox _sidecarFolderTextBox;
        private readonly Button _browseButton;
        private readonly Button _detectButton;
        private readonly Button _okButton;
        private readonly Button _cancelButton;
        private readonly DesktopEditorPathDetector _pathDetector;

        public SettingsForm(PluginSettings settings, DesktopEditorPathDetector pathDetector)
        {
            _pathDetector = pathDetector;
            this.Settings = Clone(settings);

            this.Text = "DrawioPpt Settings";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(560, 280);

            Label modeLabel = CreateLabel("Editor Mode", 16, 20);
            _editorModeComboBox = new ComboBox();
            _editorModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _editorModeComboBox.Items.Add(EditorMode.Desktop);
            _editorModeComboBox.Items.Add(EditorMode.Url);
            _editorModeComboBox.Location = new Point(150, 16);
            _editorModeComboBox.Width = 140;

            Label desktopLabel = CreateLabel("Desktop Path", 16, 58);
            _desktopPathTextBox = CreateTextBox(150, 54, 320);
            _browseButton = new Button();
            _browseButton.Text = "Browse...";
            _browseButton.Location = new Point(480, 52);
            _browseButton.Width = 64;
            _browseButton.Click += OnBrowseDesktopPath;

            _detectButton = new Button();
            _detectButton.Text = "Detect";
            _detectButton.Location = new Point(480, 80);
            _detectButton.Width = 64;
            _detectButton.Click += OnDetectDesktopPath;

            Label urlLabel = CreateLabel("Editor URL", 16, 96);
            _editorUrlTextBox = CreateTextBox(150, 92, 394);

            _autoOpenCheckBox = CreateCheckBox("Auto open on selection", 150, 128);
            _autoUpdateCheckBox = CreateCheckBox("Auto update on save", 150, 154);
            _keepSidecarCheckBox = CreateCheckBox("Keep sidecar file next to PPT", 150, 180);

            Label sidecarLabel = CreateLabel("Sidecar Folder", 16, 216);
            _sidecarFolderTextBox = CreateTextBox(150, 212, 394);

            _okButton = new Button();
            _okButton.Text = "OK";
            _okButton.Location = new Point(388, 244);
            _okButton.DialogResult = DialogResult.OK;
            _okButton.Click += OnOk;

            _cancelButton = new Button();
            _cancelButton.Text = "Cancel";
            _cancelButton.Location = new Point(472, 244);
            _cancelButton.DialogResult = DialogResult.Cancel;

            this.Controls.Add(modeLabel);
            this.Controls.Add(_editorModeComboBox);
            this.Controls.Add(desktopLabel);
            this.Controls.Add(_desktopPathTextBox);
            this.Controls.Add(_browseButton);
            this.Controls.Add(_detectButton);
            this.Controls.Add(urlLabel);
            this.Controls.Add(_editorUrlTextBox);
            this.Controls.Add(_autoOpenCheckBox);
            this.Controls.Add(_autoUpdateCheckBox);
            this.Controls.Add(_keepSidecarCheckBox);
            this.Controls.Add(sidecarLabel);
            this.Controls.Add(_sidecarFolderTextBox);
            this.Controls.Add(_okButton);
            this.Controls.Add(_cancelButton);

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
            label.Width = 120;
            return label;
        }

        private static TextBox CreateTextBox(int left, int top, int width)
        {
            TextBox textBox = new TextBox();
            textBox.Location = new Point(left, top);
            textBox.Width = width;
            return textBox;
        }

        private static PluginSettings Clone(PluginSettings settings)
        {
            PluginSettings source = settings ?? new PluginSettings();
            PluginSettings clone = new PluginSettings();
            clone.EditorMode = source.EditorMode;
            clone.DesktopEditorPath = source.DesktopEditorPath;
            clone.EditorUrl = source.EditorUrl;
            clone.AutoOpenOnSelection = source.AutoOpenOnSelection;
            clone.AutoUpdateOnSave = source.AutoUpdateOnSave;
            clone.KeepSidecarFile = source.KeepSidecarFile;
            clone.SidecarFolderName = source.SidecarFolderName;
            return clone;
        }

        private void LoadSettings(PluginSettings settings)
        {
            _editorModeComboBox.SelectedItem = settings.EditorMode;
            _desktopPathTextBox.Text = settings.DesktopEditorPath ?? string.Empty;
            _editorUrlTextBox.Text = settings.EditorUrl ?? string.Empty;
            _autoOpenCheckBox.Checked = settings.AutoOpenOnSelection;
            _autoUpdateCheckBox.Checked = settings.AutoUpdateOnSave;
            _keepSidecarCheckBox.Checked = settings.KeepSidecarFile;
            _sidecarFolderTextBox.Text = settings.SidecarFolderName ?? string.Empty;
        }

        private void OnBrowseDesktopPath(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Title = "Select draw.io or diagrams.net executable";

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
                MessageBox.Show(this, "未检测到已安装的 draw.io/diagrams.net 桌面版。", "DrawioPpt", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _desktopPathTextBox.Text = detectedPath;
        }

        private void OnOk(object sender, EventArgs e)
        {
            object selectedItem = _editorModeComboBox.SelectedItem;
            if (selectedItem == null)
            {
                MessageBox.Show(this, "Please choose an editor mode.", "DrawioPpt", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }

            PluginSettings updated = new PluginSettings();
            updated.EditorMode = (EditorMode)selectedItem;
            updated.DesktopEditorPath = _desktopPathTextBox.Text;
            updated.EditorUrl = _editorUrlTextBox.Text;
            updated.AutoOpenOnSelection = _autoOpenCheckBox.Checked;
            updated.AutoUpdateOnSave = _autoUpdateCheckBox.Checked;
            updated.KeepSidecarFile = _keepSidecarCheckBox.Checked;
            updated.SidecarFolderName = _sidecarFolderTextBox.Text;
            this.Settings = updated;
        }
    }
}
