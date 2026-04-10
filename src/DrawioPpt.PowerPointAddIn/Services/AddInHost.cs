using System;
using System.IO;
using System.Windows.Forms;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.PowerPoint;
using DrawioPpt.PowerPointAddIn.UI;
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
        private readonly DesktopEditorLauncher _desktopEditorLauncher;
        private readonly SvgFileProvider _svgFileProvider;
        private readonly PowerPointSvgShapeService _svgShapeService;
        private readonly DesktopDiagramMonitor _desktopDiagramMonitor;
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
            _desktopEditorLauncher = new DesktopEditorLauncher();
            _svgFileProvider = new SvgFileProvider();
            _svgShapeService = new PowerPointSvgShapeService(_selectedShapeAccessor);
            _desktopDiagramMonitor = new DesktopDiagramMonitor(RefreshDiagramFromMonitoredFile);
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
            _desktopDiagramMonitor.Dispose();
        }

        public void CreateNewDiagram()
        {
            if (_application == null || _application.ActivePresentation == null)
            {
                _userNotifier.ShowInfo("请先打开或创建一个 PowerPoint 演示文稿。", "DrawioPpt");
                return;
            }

            string presentationPath = _selectedShapeAccessor.GetPresentationFullName(_application);
            string diagramName = BuildNextDiagramName();
            DiagramEnvelope envelope = _shapeMetadataService.CreateNewEnvelope(_settings, presentationPath, diagramName);
            string workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, presentationPath);
            envelope.SidecarPath = workingFile;

            string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, workingFile);
            PptInterop.Shape shape = _svgShapeService.InsertOnActiveSlide(_application, svgPath, envelope.DiagramName);
            _shapeMetadataService.Save(shape, envelope);
            _currentSelection = _shapeMetadataService.BuildSelectionContext(shape);
            RaiseSelectionStateChanged();

            bool launched = TryLaunchDesktopEditor(envelope, workingFile);
            string message =
                "已创建新的 Draw.io 图形。" +
                Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                Environment.NewLine + "Working file: " + workingFile;

            if (!launched)
            {
                message += Environment.NewLine + "当前未启动外部编辑器，图形先以 SVG 预览插入。";
            }

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

            PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
            if (shape == null)
            {
                _userNotifier.ShowInfo("请先选中一个已绑定的图形。", "DrawioPpt");
                return;
            }

            DiagramEnvelope envelope;
            if (!_shapeMetadataService.TryRead(shape, out envelope))
            {
                _userNotifier.ShowInfo("当前图形缺少可读取的 Draw.io 元数据。", "DrawioPpt");
                return;
            }

            if (_settings.EditorMode == EditorMode.Desktop)
            {
                string workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, _selectedShapeAccessor.GetPresentationFullName(_application));
                envelope.SidecarPath = workingFile;
                envelope.UpdatedUtc = DateTime.UtcNow;
                _shapeMetadataService.Save(shape, envelope);
                if (!TryLaunchDesktopEditor(envelope, workingFile))
                {
                    _userNotifier.ShowInfo("请先在设置中配置本地 draw.io/diagrams.net 路径。", "DrawioPpt");
                    return;
                }

                _userNotifier.ShowInfo("已启动外部编辑器。" + Environment.NewLine + workingFile, "DrawioPpt");
                return;
            }

            string detail = "URL 模式编辑器尚未接入。";
            detail += Environment.NewLine + "Editor URL: " + _settings.EditorUrl;
            detail += Environment.NewLine + "Shape Name: " + _currentSelection.ShapeName;
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

        public void RefreshSelectedDiagram()
        {
            PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
            if (shape == null)
            {
                _userNotifier.ShowInfo("请先选中一个图形。", "DrawioPpt");
                return;
            }

            DiagramEnvelope envelope;
            if (!_shapeMetadataService.TryRead(shape, out envelope))
            {
                _userNotifier.ShowInfo("当前图形没有可刷新的 Draw.io 元数据。", "DrawioPpt");
                return;
            }

            if (string.IsNullOrWhiteSpace(envelope.SidecarPath) || !File.Exists(envelope.SidecarPath))
            {
                _userNotifier.ShowInfo("当前图形还没有可读取的 .drawio 工作文件。", "DrawioPpt");
                return;
            }

            RefreshDiagramFromFile(shape, envelope, envelope.SidecarPath, true);
        }

        public void OpenSettings()
        {
            using (SettingsForm form = new SettingsForm(_settings))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                _settings = form.Settings;
                _settingsStore.Save(_settings);
                _userNotifier.ShowInfo("设置已保存。", "DrawioPpt");
            }
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

        private string BuildNextDiagramName()
        {
            return "Drawio Diagram " + DateTime.Now.ToString("yyyyMMdd-HHmmss");
        }

        private void RefreshDiagramFromMonitoredFile(string diagramId, string filePath)
        {
            if (!_settings.AutoUpdateOnSave)
            {
                return;
            }

            PptInterop.Shape shape = _selectedShapeAccessor.FindShapeByDiagramId(_application, diagramId);
            if (shape == null)
            {
                return;
            }

            DiagramEnvelope envelope;
            if (!_shapeMetadataService.TryRead(shape, out envelope))
            {
                return;
            }

            RefreshDiagramFromFile(shape, envelope, filePath, false);
        }

        private void RefreshDiagramFromFile(PptInterop.Shape shape, DiagramEnvelope envelope, string filePath, bool notifyUser)
        {
            if (shape == null || envelope == null || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            string xmlContent;
            if (!TryReadWorkingFile(filePath, out xmlContent))
            {
                if (notifyUser)
                {
                    _userNotifier.ShowInfo("当前无法读取 .drawio 文件，可能仍在保存中。请稍后重试。", "DrawioPpt");
                }

                return;
            }

            envelope.DrawioXml = xmlContent;
            envelope.SidecarPath = filePath;
            envelope.UpdatedUtc = DateTime.UtcNow;

            string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, filePath);
            PptInterop.Shape newShape = _svgShapeService.Replace(shape, svgPath);
            _shapeMetadataService.Save(newShape, envelope);
            _currentSelection = _shapeMetadataService.BuildSelectionContext(newShape);
            RaiseSelectionStateChanged();

            if (notifyUser)
            {
                _userNotifier.ShowInfo("已刷新当前图形。", "DrawioPpt");
            }
        }

        private static bool TryReadWorkingFile(string filePath, out string xmlContent)
        {
            xmlContent = string.Empty;

            int attempt;
            for (attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    xmlContent = File.ReadAllText(filePath);
                    return true;
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }

                System.Threading.Thread.Sleep(250);
            }

            return false;
        }

        private bool TryLaunchDesktopEditor(DiagramEnvelope envelope, string workingFile)
        {
            if (_settings.EditorMode != EditorMode.Desktop)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(_settings.DesktopEditorPath) || !File.Exists(_settings.DesktopEditorPath))
            {
                return false;
            }

            _desktopEditorLauncher.Launch(_settings.DesktopEditorPath, workingFile);
            _desktopDiagramMonitor.Track(envelope.DiagramId, workingFile);
            return true;
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
