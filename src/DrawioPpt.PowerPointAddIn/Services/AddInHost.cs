using System;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.PowerPoint;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class AddInHost : IDisposable
    {
        private readonly IDiagramEnvelopeSerializer _envelopeSerializer;
        private readonly IPluginSettingsStore _settingsStore;
        private readonly SelectionMonitor _selectionMonitor;
        private readonly UserNotifier _userNotifier;
        private PluginSettings _settings;
        private SelectionContext _currentSelection;

        public AddInHost(PptInterop.Application application)
        {
            _envelopeSerializer = new DiagramEnvelopeSerializer();
            _settingsStore = new FilePluginSettingsStore();
            _userNotifier = new UserNotifier();
            _selectionMonitor = new SelectionMonitor(application, new PowerPointShapeSelectionReader(_envelopeSerializer));
            _selectionMonitor.SelectionChanged += OnSelectionChanged;
            _currentSelection = new SelectionContext();
        }

        public event EventHandler SelectionStateChanged;

        public bool HasEditableSelection
        {
            get { return _currentSelection != null && _currentSelection.IsManagedShape; }
        }

        public void Start()
        {
            _settings = _settingsStore.Load();
            _selectionMonitor.Start();
        }

        public void Dispose()
        {
            _selectionMonitor.Dispose();
        }

        public void CreateNewDiagram()
        {
            string message =
                "初始版本已完成宿主骨架。" +
                Environment.NewLine + Environment.NewLine +
                "下一阶段将把这里接到真正的 Draw.io 新建流程，包括：" +
                Environment.NewLine +
                "- 启动本地或 URL 编辑器" +
                Environment.NewLine +
                "- 生成 SVG" +
                Environment.NewLine +
                "- 写入 diagram 元数据";

            _userNotifier.ShowInfo(message, "DrawioPpt");
        }

        public void EditSelectedDiagram()
        {
            if (_currentSelection == null || !_currentSelection.IsManagedShape)
            {
                _userNotifier.ShowInfo("当前未选中插件管理的 Draw.io 图形。", "DrawioPpt");
                return;
            }

            string detail = "已识别到可编辑图形。";
            detail += Environment.NewLine + "Shape ID: " + _currentSelection.ShapeId;
            detail += Environment.NewLine + "Shape Name: " + _currentSelection.ShapeName;
            detail += Environment.NewLine + "Diagram ID: " + _currentSelection.DiagramId;

            if (_envelopeSerializer.CanDeserialize(_currentSelection.AlternativeText))
            {
                DiagramEnvelope envelope = _envelopeSerializer.Deserialize(_currentSelection.AlternativeText);
                detail += Environment.NewLine + "Editor Mode: " + envelope.EditorMode;
                detail += Environment.NewLine + "Updated UTC: " + envelope.UpdatedUtc.ToString("u");
            }

            detail += Environment.NewLine + Environment.NewLine + "下一阶段会在这里接外部编辑器并回写 SVG。";
            _userNotifier.ShowInfo(detail, "DrawioPpt");
        }

        public void OpenSettings()
        {
            string message =
                "当前设置摘要：" +
                Environment.NewLine + "Mode: " + _settings.EditorMode +
                Environment.NewLine + "Desktop Path: " + _settings.DesktopEditorPath +
                Environment.NewLine + "Editor URL: " + _settings.EditorUrl +
                Environment.NewLine + "Keep Sidecar: " + _settings.KeepSidecarFile +
                Environment.NewLine +
                Environment.NewLine +
                "下一阶段将补完整设置对话框。";

            _userNotifier.ShowInfo(message, "DrawioPpt");
        }

        public string GetSelectionSummary()
        {
            if (_currentSelection == null || !_currentSelection.HasSelection)
            {
                return "Draw.io: 未选中图形";
            }

            if (!_currentSelection.HasSingleShape)
            {
                return "Draw.io: 请选择单个图形";
            }

            if (_currentSelection.IsManagedShape)
            {
                return "Draw.io: 已识别 " + _currentSelection.ShapeName;
            }

            return "Draw.io: 当前图形未绑定";
        }

        private void OnSelectionChanged(object sender, SelectionContextChangedEventArgs e)
        {
            _currentSelection = e.Context;

            EventHandler handler = this.SelectionStateChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
