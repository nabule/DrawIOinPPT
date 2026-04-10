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
        private readonly PptInterop.Application _application;
        private readonly IDiagramEnvelopeSerializer _envelopeSerializer;
        private readonly IPluginSettingsStore _settingsStore;
        private readonly SelectionMonitor _selectionMonitor;
        private readonly SelectedShapeAccessor _selectedShapeAccessor;
        private readonly ShapeMetadataService _shapeMetadataService;
        private readonly UserNotifier _userNotifier;
        private PluginSettings _settings;
        private SelectionContext _currentSelection;

        public AddInHost(PptInterop.Application application)
        {
            _application = application;
            _envelopeSerializer = new DiagramEnvelopeSerializer();
            _settingsStore = new FilePluginSettingsStore();
            _selectedShapeAccessor = new SelectedShapeAccessor();
            _shapeMetadataService = new ShapeMetadataService(_envelopeSerializer, new PresentationSidecarPathBuilder());
            _userNotifier = new UserNotifier();
            _selectionMonitor = new SelectionMonitor(application, new PowerPointShapeSelectionReader(_envelopeSerializer));
            _selectionMonitor.SelectionChanged += OnSelectionChanged;
            _currentSelection = new SelectionContext();
            _settings = new PluginSettings();
        }

        public event EventHandler SelectionStateChanged;

        public bool HasEditableSelection
        {
            get { return _currentSelection != null && _currentSelection.IsManagedShape; }
        }

        public bool HasSingleShapeSelection
        {
            get { return _currentSelection != null && _currentSelection.HasSingleShape; }
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

        public void BindSelectedShape()
        {
            PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
            if (shape == null)
            {
                _userNotifier.ShowInfo("请先选中一个图形，再执行绑定。", "DrawioPpt");
                return;
            }

            string presentationPath = _selectedShapeAccessor.GetPresentationFullName(_application);
            DiagramEnvelope envelope = _shapeMetadataService.CreateOrUpdateEnvelope(shape, _settings, presentationPath);
            _shapeMetadataService.Save(shape, envelope);
            _currentSelection = _shapeMetadataService.BuildSelectionContext(shape);
            RaiseSelectionStateChanged();

            string message =
                "已为当前图形写入 Draw.io 元数据。" +
                Environment.NewLine + "Shape Name: " + shape.Name +
                Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                Environment.NewLine + "Sidecar: " + envelope.SidecarPath;

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

        public void ClearSelectedShapeBinding()
        {
            PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
            if (shape == null)
            {
                _userNotifier.ShowInfo("请先选中一个图形，再执行清除绑定。", "DrawioPpt");
                return;
            }

            _shapeMetadataService.Clear(shape);
            _currentSelection = new SelectionContext();
            _currentSelection.HasSelection = true;
            _currentSelection.HasSingleShape = true;
            _currentSelection.ShapeId = shape.Id;
            _currentSelection.ShapeName = shape.Name;
            RaiseSelectionStateChanged();

            _userNotifier.ShowInfo("已清除当前图形上的 Draw.io 绑定信息。", "DrawioPpt");
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
            RaiseSelectionStateChanged();
        }

        private void RaiseSelectionStateChanged()
        {
            EventHandler handler = this.SelectionStateChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
