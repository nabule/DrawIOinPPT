using System;
using System.Collections.Generic;
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
        private readonly DesktopEditorPathDetector _desktopEditorPathDetector;
        private readonly DesktopEditorLauncher _desktopEditorLauncher;
        private readonly PresentationDiagramStore _presentationDiagramStore;
        private readonly SvgFileProvider _svgFileProvider;
        private readonly SvgMarkupFileStore _svgMarkupFileStore;
        private readonly PowerPointSvgShapeService _svgShapeService;
        private readonly DesktopDiagramMonitor _desktopDiagramMonitor;
        private readonly PluginTraceLog _traceLog;
        private readonly UserNotifier _userNotifier;
        private PluginSettings _settings;
        private SelectionContext _currentSelection;
        private string _lastAutoOpenedDiagramId;
        private DateTime _lastAutoOpenUtc;
        private DateTime _suppressAutoOpenUntilUtc;
        private DateTime _lastCleanupUtc;
        private string _lastCleanupPresentationPath;

        public AddInHost(PptInterop.Application application)
        {
            _application = application;
            _envelopeSerializer = new DiagramEnvelopeSerializer();
            _settingsStore = new FilePluginSettingsStore();
            _selectedShapeAccessor = new SelectedShapeAccessor();
            _shapeMetadataService = new ShapeMetadataService(_envelopeSerializer, new PresentationSidecarPathBuilder());
            _desktopEditorPathDetector = new DesktopEditorPathDetector();
            _desktopEditorLauncher = new DesktopEditorLauncher();
            _presentationDiagramStore = new PresentationDiagramStore(_envelopeSerializer);
            _svgFileProvider = new SvgFileProvider();
            _svgMarkupFileStore = new SvgMarkupFileStore();
            _svgShapeService = new PowerPointSvgShapeService(_selectedShapeAccessor);
            _desktopDiagramMonitor = new DesktopDiagramMonitor(RefreshDiagramFromMonitoredFile);
            _traceLog = new PluginTraceLog();
            _userNotifier = new UserNotifier();
            _selectionMonitor = new SelectionMonitor(application, new PowerPointShapeSelectionReader(_envelopeSerializer));
            _selectionMonitor.SelectionChanged += OnSelectionChanged;
            _selectionMonitor.SelectionDoubleClicked += OnSelectionDoubleClicked;
            _currentSelection = new SelectionContext();
            _settings = new PluginSettings();
            _lastAutoOpenedDiagramId = string.Empty;
            _lastAutoOpenUtc = DateTime.MinValue;
            _suppressAutoOpenUntilUtc = DateTime.MinValue;
            _lastCleanupUtc = DateTime.MinValue;
            _lastCleanupPresentationPath = string.Empty;
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
            ExecuteGuarded("Start", "初始化插件时出现异常。", false, delegate
            {
                _settings = _settingsStore.Load();
                if (string.IsNullOrWhiteSpace(_settings.DesktopEditorPath))
                {
                    _settings.DesktopEditorPath = _desktopEditorPathDetector.Detect();
                }

                _traceLog.Info("AddInHost", "Starting add-in host. EditorMode=" + _settings.EditorMode + ", EditorUrl=" + (_settings.EditorUrl ?? string.Empty) + ", DesktopPath=" + (_settings.DesktopEditorPath ?? string.Empty));
                _selectionMonitor.Start();
                MaybeCleanupActivePresentation(true);
            });
        }

        public void Dispose()
        {
            _selectionMonitor.Dispose();
            _desktopDiagramMonitor.Dispose();
        }

        public void CreateNewDiagram()
        {
            ExecuteGuarded("CreateNewDiagram", "创建 Draw.io 图形时出现异常。", true, delegate
            {
                PptInterop.Presentation presentation = _selectedShapeAccessor.GetActivePresentation(_application);
                if (_application == null || presentation == null)
                {
                    _userNotifier.ShowInfo("请先打开或创建一个 PowerPoint 演示文稿。", "DrawioPpt");
                    return;
                }

                string presentationPath = _selectedShapeAccessor.GetPresentationFullName(_application);
                string diagramName = BuildNextDiagramName();
                DiagramEnvelope envelope = _shapeMetadataService.CreateNewEnvelope(_settings, presentationPath, diagramName);
                _traceLog.Info("AddInHost", "Creating new diagram " + envelope.DiagramId + " in mode " + _settings.EditorMode + ".");
                string workingFile = envelope.SidecarPath;

                if (_settings.EditorMode == EditorMode.Desktop)
                {
                    workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, presentationPath);
                    envelope.SidecarPath = workingFile;
                }
                else
                {
                    WriteSidecarFile(envelope);
                }

                string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, workingFile);
                SuppressAutoOpenForSeconds(2);
                PptInterop.Shape shape = _svgShapeService.InsertOnActiveSlide(_application, svgPath, envelope.DiagramName);
                SaveManagedEnvelope(shape, envelope);
                _currentSelection = _shapeMetadataService.BuildSelectionContext(shape);
                RaiseSelectionStateChanged();

                bool launched = TryOpenShapeForEditing(shape, envelope, false);
                if (_settings.EditorMode == EditorMode.Url)
                {
                    return;
                }

                string message =
                    "已创建新的 Draw.io 图形。" +
                    Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                    Environment.NewLine + "Working file: " + workingFile;

                if (!launched)
                {
                    message += Environment.NewLine + "当前未启动外部编辑器，图形先以 SVG 预览插入。";
                }

                _userNotifier.ShowInfo(message, "DrawioPpt");
            });
        }

        public void BindSelectedShape()
        {
            ExecuteGuarded("BindSelectedShape", "绑定当前图形时出现异常。", true, delegate
            {
                PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
                if (shape == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图形，再执行绑定。", "DrawioPpt");
                    return;
                }

                string presentationPath = _selectedShapeAccessor.GetPresentationFullName(_application);
                DiagramEnvelope envelope = _shapeMetadataService.CreateOrUpdateEnvelope(shape, _settings, presentationPath);
                SaveManagedEnvelope(shape, envelope);
                _traceLog.Info("AddInHost", "Bound existing shape '" + shape.Name + "' to diagram " + envelope.DiagramId + ".");
                _currentSelection = _shapeMetadataService.BuildSelectionContext(shape);
                RaiseSelectionStateChanged();

                string message =
                    "已为当前图形写入 Draw.io 元数据。" +
                    Environment.NewLine + "Shape Name: " + shape.Name +
                    Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                    Environment.NewLine + "Sidecar: " + envelope.SidecarPath;

                _userNotifier.ShowInfo(message, "DrawioPpt");
            });
        }

        public void EditSelectedDiagram()
        {
            ExecuteGuarded("EditSelectedDiagram", "打开图形编辑器时出现异常。", true, delegate
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
                if (!TryReadManagedEnvelope(shape, out envelope))
                {
                    _userNotifier.ShowInfo("当前图形缺少可读取的 Draw.io 元数据。", "DrawioPpt");
                    return;
                }

                _traceLog.Info("AddInHost", "Editing diagram " + envelope.DiagramId + " from shape '" + shape.Name + "'.");
                if (!TryOpenShapeForEditing(shape, envelope, true))
                {
                    if (_settings.EditorMode == EditorMode.Desktop)
                    {
                        _userNotifier.ShowInfo("请先在设置中配置本地 draw.io/diagrams.net 路径。", "DrawioPpt");
                    }
                }
            });
        }

        public void ClearSelectedShapeBinding()
        {
            ExecuteGuarded("ClearSelectedShapeBinding", "清除图形绑定时出现异常。", true, delegate
            {
                PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
                if (shape == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图形，再执行清除绑定。", "DrawioPpt");
                    return;
                }

                DeleteManagedEnvelope(shape);
                _shapeMetadataService.Clear(shape);
                MaybeCleanupActivePresentation(true);
                _traceLog.Info("AddInHost", "Cleared binding for shape '" + shape.Name + "'.");
                _currentSelection = new SelectionContext();
                _currentSelection.HasSelection = true;
                _currentSelection.HasSingleShape = true;
                _currentSelection.ShapeId = shape.Id;
                _currentSelection.ShapeName = shape.Name;
                RaiseSelectionStateChanged();

                _userNotifier.ShowInfo("已清除当前图形上的 Draw.io 绑定信息。", "DrawioPpt");
            });
        }

        public void RefreshSelectedDiagram()
        {
            ExecuteGuarded("RefreshSelectedDiagram", "刷新当前图形时出现异常。", true, delegate
            {
                PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
                if (shape == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图形。", "DrawioPpt");
                    return;
                }

                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(shape, out envelope))
                {
                    _userNotifier.ShowInfo("当前图形没有可刷新的 Draw.io 元数据。", "DrawioPpt");
                    return;
                }

                if (string.IsNullOrWhiteSpace(envelope.SidecarPath) || !File.Exists(envelope.SidecarPath))
                {
                    _userNotifier.ShowInfo("当前图形还没有可读取的 .drawio 工作文件。", "DrawioPpt");
                    return;
                }

                _traceLog.Info("AddInHost", "Refreshing diagram " + envelope.DiagramId + " from file " + envelope.SidecarPath + ".");
                RefreshDiagramFromFile(shape, envelope, envelope.SidecarPath, true);
            });
        }

        public void OpenSettings()
        {
            ExecuteGuarded("OpenSettings", "打开或保存设置时出现异常。", true, delegate
            {
                using (SettingsForm form = new SettingsForm(_settings, _desktopEditorPathDetector))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    _settings = form.Settings;
                    _settingsStore.Save(_settings);
                    _traceLog.Info("AddInHost", "Settings saved. EditorMode=" + _settings.EditorMode + ", EditorUrl=" + (_settings.EditorUrl ?? string.Empty) + ", DesktopPath=" + (_settings.DesktopEditorPath ?? string.Empty));
                    _userNotifier.ShowInfo("设置已保存。", "DrawioPpt");
                }
            });
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
            ExecuteGuarded("OnSelectionChanged", "处理图形选中状态时出现异常。", false, delegate
            {
                _currentSelection = e.Context;
                RaiseSelectionStateChanged();
                MaybeCleanupActivePresentation(false);
                MaybeAutoOpenSelection();
            });
        }

        private void OnSelectionDoubleClicked(object sender, SelectionDoubleClickEventArgs e)
        {
            ExecuteGuarded("OnSelectionDoubleClicked", "处理双击编辑时出现异常。", false, delegate
            {
                if (e == null || e.Selection == null)
                {
                    return;
                }

                if (e.Selection.Type != PptInterop.PpSelectionType.ppSelectionShapes ||
                    e.Selection.ShapeRange == null ||
                    e.Selection.ShapeRange.Count != 1)
                {
                    return;
                }

                PptInterop.Shape shape = e.Selection.ShapeRange[1];
                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(shape, out envelope))
                {
                    return;
                }

                if (TryOpenShapeForEditing(shape, envelope, true))
                {
                    e.Cancel = true;
                }
            });
        }

        private string BuildNextDiagramName()
        {
            return "Drawio Diagram " + DateTime.Now.ToString("yyyyMMdd-HHmmss");
        }

        private void RefreshDiagramFromMonitoredFile(string diagramId, string filePath)
        {
            ExecuteGuarded("RefreshDiagramFromMonitoredFile", "监控到文件变化后刷新图形时出现异常。", false, delegate
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
                if (!TryReadManagedEnvelope(shape, out envelope))
                {
                    return;
                }

                RefreshDiagramFromFile(shape, envelope, filePath, false);
            });
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
            _traceLog.Info("AddInHost", "Applying refreshed diagram " + envelope.DiagramId + " from monitored file " + filePath + ".");

            string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, filePath);
            SuppressAutoOpenForSeconds(2);
            PptInterop.Shape newShape = _svgShapeService.Replace(shape, svgPath);
            SaveManagedEnvelope(newShape, envelope);
            _currentSelection = _shapeMetadataService.BuildSelectionContext(newShape);
            RaiseSelectionStateChanged();

            if (notifyUser)
            {
                _userNotifier.ShowInfo("已刷新当前图形。", "DrawioPpt");
            }
        }

        private bool TryOpenShapeForEditing(PptInterop.Shape shape, DiagramEnvelope envelope, bool explicitUserAction)
        {
            if (shape == null || envelope == null)
            {
                return false;
            }

            if (_settings.EditorMode == EditorMode.Desktop)
            {
                string presentationPath = _selectedShapeAccessor.GetPresentationFullName(_application);
                string workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, presentationPath);
                envelope.SidecarPath = workingFile;
                envelope.UpdatedUtc = DateTime.UtcNow;
                SaveManagedEnvelope(shape, envelope);
                _traceLog.Info("AddInHost", "Prepared desktop editor working file for diagram " + envelope.DiagramId + ": " + workingFile);

                if (!LaunchDesktopEditor(envelope, workingFile, explicitUserAction))
                {
                    return false;
                }

                if (explicitUserAction)
                {
                    _userNotifier.ShowInfo("已启动外部编辑器。" + Environment.NewLine + workingFile, "DrawioPpt");
                }

                return true;
            }

            if (_settings.EditorMode == EditorMode.Url)
            {
                if (string.IsNullOrWhiteSpace(_settings.EditorUrl))
                {
                    if (explicitUserAction)
                    {
                        _userNotifier.ShowInfo("请先在设置中配置 URL 模式编辑器地址。", "DrawioPpt");
                    }

                    return false;
                }

                WriteSidecarFile(envelope);
                _traceLog.Info("AddInHost", "Opening URL editor for diagram " + envelope.DiagramId + " with URL " + _settings.EditorUrl + ".");
                using (UrlDiagramEditorForm form = new UrlDiagramEditorForm(_settings.EditorUrl, envelope.DiagramName, envelope.DrawioXml, _traceLog))
                {
                    form.DiagramSaved += delegate(object sender, UrlDiagramSavedEventArgs args)
                    {
                        ApplyUrlEditorSave(shape, envelope, args);
                    };

                    form.ShowDialog();
                }

                return true;
            }

            if (explicitUserAction)
            {
                _userNotifier.ShowInfo("当前编辑模式不受支持。", "DrawioPpt");
            }

            return false;
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

        private bool LaunchDesktopEditor(DiagramEnvelope envelope, string workingFile, bool explicitUserAction)
        {
            if (_settings.EditorMode != EditorMode.Desktop)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(_settings.DesktopEditorPath) || !File.Exists(_settings.DesktopEditorPath))
            {
                if (explicitUserAction)
                {
                    _userNotifier.ShowInfo("请先在设置中配置本地 draw.io/diagrams.net 路径。", "DrawioPpt");
                }

                _traceLog.Warn("AddInHost", "Desktop editor path is missing or invalid.");
                return false;
            }

            _traceLog.Info("AddInHost", "Launching desktop editor for diagram " + envelope.DiagramId + " using " + _settings.DesktopEditorPath + ".");
            _desktopEditorLauncher.Launch(_settings.DesktopEditorPath, workingFile);
            _desktopDiagramMonitor.Track(envelope.DiagramId, workingFile);
            return true;
        }

        private void ApplyUrlEditorSave(PptInterop.Shape originalShape, DiagramEnvelope envelope, UrlDiagramSavedEventArgs args)
        {
            if (envelope == null || args == null || string.IsNullOrWhiteSpace(args.Xml) || string.IsNullOrWhiteSpace(args.SvgMarkup))
            {
                return;
            }

            try
            {
                envelope.DrawioXml = args.Xml;
                envelope.UpdatedUtc = DateTime.UtcNow;
                WriteSidecarFile(envelope);
                _traceLog.Info("AddInHost", "Received URL editor save for diagram " + envelope.DiagramId + ". ExitRequested=" + args.ExitRequested);

                string svgPath = _svgMarkupFileStore.Write(envelope, args.SvgMarkup);
                PptInterop.Shape shape = _selectedShapeAccessor.FindShapeByDiagramId(_application, envelope.DiagramId) ?? originalShape;
                if (shape == null)
                {
                    return;
                }

                SuppressAutoOpenForSeconds(2);
                PptInterop.Shape newShape = _svgShapeService.Replace(shape, svgPath);
                SaveManagedEnvelope(newShape, envelope);
                _currentSelection = _shapeMetadataService.BuildSelectionContext(newShape);
                RaiseSelectionStateChanged();
            }
            catch (Exception ex)
            {
                _traceLog.Error("AddInHost", "Failed to apply URL editor save for diagram " + envelope.DiagramId + ".", ex);
                _userNotifier.ShowInfo("URL 模式保存后刷新失败。" + Environment.NewLine + ex.Message, "DrawioPpt");
            }
        }

        private static void WriteSidecarFile(DiagramEnvelope envelope)
        {
            if (envelope == null || string.IsNullOrWhiteSpace(envelope.SidecarPath))
            {
                return;
            }

            string directory = Path.GetDirectoryName(envelope.SidecarPath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(envelope.SidecarPath, envelope.DrawioXml ?? string.Empty);
        }

        private void MaybeAutoOpenSelection()
        {
            if (!_settings.AutoOpenOnSelection || _currentSelection == null || !_currentSelection.IsManagedShape)
            {
                return;
            }

            if (DateTime.UtcNow < _suppressAutoOpenUntilUtc)
            {
                return;
            }

            string diagramId = _currentSelection.DiagramId ?? string.Empty;
            if (string.IsNullOrWhiteSpace(diagramId))
            {
                return;
            }

            if (string.Equals(_lastAutoOpenedDiagramId, diagramId, StringComparison.OrdinalIgnoreCase) &&
                DateTime.UtcNow.Subtract(_lastAutoOpenUtc).TotalSeconds < 2)
            {
                return;
            }

            PptInterop.Shape shape = _selectedShapeAccessor.GetSingleSelectedShape(_application);
            if (shape == null)
            {
                return;
            }

            DiagramEnvelope envelope;
            if (!TryReadManagedEnvelope(shape, out envelope))
            {
                return;
            }

            if (TryOpenShapeForEditing(shape, envelope, false))
            {
                _lastAutoOpenedDiagramId = diagramId;
                _lastAutoOpenUtc = DateTime.UtcNow;
            }
        }

        private bool TryReadManagedEnvelope(PptInterop.Shape shape, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (shape == null)
            {
                return false;
            }

            DiagramEnvelope fallbackEnvelope = null;
            bool hasFallbackEnvelope = _shapeMetadataService.TryRead(shape, out fallbackEnvelope);
            PptInterop.Presentation presentation = _selectedShapeAccessor.GetPresentation(shape) ?? _selectedShapeAccessor.GetActivePresentation(_application);
            string diagramId = _currentSelection != null && _currentSelection.ShapeId == shape.Id
                ? _currentSelection.DiagramId
                : string.Empty;

            if (string.IsNullOrWhiteSpace(diagramId))
            {
                SelectionContext context = _shapeMetadataService.BuildSelectionContext(shape);
                diagramId = context.DiagramId;
            }

            string partId = _shapeMetadataService.ReadDiagramPartId(shape);
            if (hasFallbackEnvelope &&
                string.IsNullOrWhiteSpace(partId) &&
                !string.IsNullOrWhiteSpace(fallbackEnvelope.DrawioXml))
            {
                SaveManagedEnvelope(shape, fallbackEnvelope);
                partId = _shapeMetadataService.ReadDiagramPartId(shape);
                _traceLog.Info("AddInHost", "Rehydrated legacy AlternativeText metadata for diagram " + fallbackEnvelope.DiagramId + ".");
            }

            string resolvedPartId;
            DiagramEnvelope storedEnvelope;
            if (_presentationDiagramStore.TryRead(presentation, diagramId, partId, out storedEnvelope, out resolvedPartId))
            {
                envelope = storedEnvelope;

                if (!string.Equals(partId, resolvedPartId, StringComparison.OrdinalIgnoreCase) || !hasFallbackEnvelope)
                {
                    _shapeMetadataService.Save(shape, storedEnvelope, resolvedPartId);
                }

                return true;
            }

            if (hasFallbackEnvelope)
            {
                envelope = fallbackEnvelope;
                return true;
            }

            return false;
        }

        private void SaveManagedEnvelope(PptInterop.Shape shape, DiagramEnvelope envelope)
        {
            if (shape == null || envelope == null)
            {
                return;
            }

            PptInterop.Presentation presentation = _selectedShapeAccessor.GetPresentation(shape) ?? _selectedShapeAccessor.GetActivePresentation(_application);
            if (presentation == null)
            {
                _traceLog.Info("AddInHost", "Skipped saving managed envelope because no presentation was available.");
                return;
            }

            string existingPartId = _shapeMetadataService.ReadDiagramPartId(shape);
            string storedPartId = _presentationDiagramStore.Upsert(presentation, envelope, existingPartId);

            if (string.IsNullOrWhiteSpace(storedPartId))
            {
                storedPartId = existingPartId;
            }

            _shapeMetadataService.Save(shape, envelope, storedPartId);
        }

        private void DeleteManagedEnvelope(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            PptInterop.Presentation presentation = _selectedShapeAccessor.GetPresentation(shape) ?? _selectedShapeAccessor.GetActivePresentation(_application);
            if (presentation == null)
            {
                return;
            }

            string partId = _shapeMetadataService.ReadDiagramPartId(shape);
            SelectionContext context = _shapeMetadataService.BuildSelectionContext(shape);
            _presentationDiagramStore.Delete(presentation, context.DiagramId, partId);
        }

        private void MaybeCleanupActivePresentation(bool force)
        {
            PptInterop.Presentation presentation = _selectedShapeAccessor.GetActivePresentation(_application);
            if (presentation == null)
            {
                return;
            }

            string presentationPath = presentation.FullName ?? string.Empty;
            bool presentationChanged = !string.Equals(_lastCleanupPresentationPath, presentationPath, StringComparison.OrdinalIgnoreCase);
            if (!force &&
                !presentationChanged &&
                DateTime.UtcNow.Subtract(_lastCleanupUtc).TotalSeconds < 30)
            {
                return;
            }

            try
            {
                HashSet<string> liveDiagramIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> livePartIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _selectedShapeAccessor.CollectManagedDiagramReferences(presentation, liveDiagramIds, livePartIds);
                int deletedCount = _presentationDiagramStore.DeleteOrphans(presentation, liveDiagramIds, livePartIds);
                _lastCleanupUtc = DateTime.UtcNow;
                _lastCleanupPresentationPath = presentationPath;
                if (deletedCount > 0)
                {
                    _traceLog.Info("AddInHost", "Removed " + deletedCount + " orphaned CustomXMLParts from active presentation.");
                }
            }
            catch (Exception ex)
            {
                _traceLog.Error("AddInHost", "Failed while cleaning orphaned CustomXMLParts.", ex);
            }
        }

        private void SuppressAutoOpenForSeconds(int seconds)
        {
            _suppressAutoOpenUntilUtc = DateTime.UtcNow.AddSeconds(seconds);
        }

        private void ExecuteGuarded(string operationName, string userMessage, bool notifyUser, Action action)
        {
            if (action == null)
            {
                return;
            }

            try
            {
                action();
            }
            catch (Exception ex)
            {
                _traceLog.Error("AddInHost", "Operation failed: " + (operationName ?? "Unknown"), ex);
                if (notifyUser)
                {
                    string message =
                        (string.IsNullOrWhiteSpace(userMessage) ? "操作执行失败。" : userMessage) +
                        Environment.NewLine + ex.Message +
                        Environment.NewLine + "日志: " + _traceLog.LogPath;
                    _userNotifier.ShowError(message, "DrawioPpt");
                }
            }
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
