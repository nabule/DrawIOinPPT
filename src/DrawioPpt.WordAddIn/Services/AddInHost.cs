using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.Services;
using DrawioPpt.PowerPointAddIn.UI;
using DrawioPpt.WordAddIn.Word;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Services
{
    public class AddInHost : IDisposable
    {
        private readonly WordInterop.Application _application;
        private readonly IDiagramEnvelopeSerializer _envelopeSerializer;
        private readonly IPluginSettingsStore _settingsStore;
        private readonly SelectedPictureAccessor _selectedPictureAccessor;
        private readonly WordPictureMetadataService _pictureMetadataService;
        private readonly DesktopEditorPathDetector _desktopEditorPathDetector;
        private readonly DesktopEditorLauncher _desktopEditorLauncher;
        private readonly DocumentDiagramStore _documentDiagramStore;
        private readonly SvgFileProvider _svgFileProvider;
        private readonly SvgMarkupFileStore _svgMarkupFileStore;
        private readonly WordSvgPictureService _svgPictureService;
        private readonly DesktopDiagramMonitor _desktopDiagramMonitor;
        private readonly PluginTraceLog _traceLog;
        private readonly UserNotifier _userNotifier;
        private readonly SelectionMonitor _selectionMonitor;
        private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);
        private PluginSettings _settings;
        private SelectionContext _currentSelection;
        private DateTime _lastCleanupUtc;
        private string _lastCleanupDocumentPath;

        public AddInHost(WordInterop.Application application)
        {
            _application = application;
            _envelopeSerializer = new DiagramEnvelopeSerializer();
            _settingsStore = new FilePluginSettingsStore();
            _selectedPictureAccessor = new SelectedPictureAccessor(_envelopeSerializer);
            _pictureMetadataService = new WordPictureMetadataService(_envelopeSerializer, new PresentationSidecarPathBuilder());
            _desktopEditorPathDetector = new DesktopEditorPathDetector();
            _desktopEditorLauncher = new DesktopEditorLauncher();
            _documentDiagramStore = new DocumentDiagramStore(_envelopeSerializer);
            _svgFileProvider = new SvgFileProvider();
            _svgMarkupFileStore = new SvgMarkupFileStore();
            _svgPictureService = new WordSvgPictureService(_selectedPictureAccessor);
            _desktopDiagramMonitor = new DesktopDiagramMonitor(RefreshDiagramFromMonitoredFile);
            _traceLog = new PluginTraceLog();
            _userNotifier = new UserNotifier();
            _selectionMonitor = new SelectionMonitor(application);
            _selectionMonitor.SelectionDoubleClicked += OnSelectionDoubleClicked;
            _currentSelection = new SelectionContext();
            _settings = new PluginSettings();
            _lastCleanupUtc = DateTime.MinValue;
            _lastCleanupDocumentPath = string.Empty;
        }

        public event EventHandler SelectionStateChanged;

        public void Start()
        {
            ExecuteGuarded("Start", "初始化 Word 插件时出现异常。", false, delegate
            {
                _settings = _settingsStore.Load();
                if (string.IsNullOrWhiteSpace(_settings.DesktopEditorPath))
                {
                    _settings.DesktopEditorPath = _desktopEditorPathDetector.Detect();
                }

                _traceLog.Info("WordAddInHost", "Starting Word add-in host. EditorMode=" + _settings.EditorMode + ", EditorUrl=" + (_settings.EditorUrl ?? string.Empty) + ", DesktopPath=" + (_settings.DesktopEditorPath ?? string.Empty));
                _selectionMonitor.Start();
                _currentSelection = new SelectionContext();
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
                WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
                if (_application == null || document == null)
                {
                    _userNotifier.ShowInfo("请先打开或创建一个 Word 文档。", "DrawioWord");
                    return;
                }

                string documentPath = _selectedPictureAccessor.GetDocumentFullName(_application);
                string diagramName = BuildNextDiagramName();
                DiagramEnvelope envelope = _pictureMetadataService.CreateNewEnvelope(_settings, documentPath, diagramName);
                _traceLog.Info("WordAddInHost", "Creating new diagram " + envelope.DiagramId + " in mode " + _settings.EditorMode + ".");
                string workingFile = envelope.SidecarPath;

                if (_settings.EditorMode == EditorMode.Desktop)
                {
                    workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, documentPath);
                    envelope.SidecarPath = workingFile;
                }
                else
                {
                    WriteSidecarFile(envelope);
                }

                string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, workingFile);
                WordPictureReference picture = _svgPictureService.InsertAtSelection(_application, svgPath, envelope.DiagramName);
                SaveManagedEnvelope(picture, envelope);
                ApplySelectionContext(_pictureMetadataService.BuildSelectionContext(picture));

                bool launched = TryOpenPictureForEditing(picture, envelope, false);
                ShowDiagramInfoDialog("已创建新的 Draw.io 图形。", envelope, picture, workingFile, launched);
            });
        }

        public void BindSelectedPicture()
        {
            ExecuteGuarded("BindSelectedPicture", "绑定当前图片时出现异常。", true, delegate
            {
                WordPictureReference picture;
                SelectionContext selection = ReadLiveSelectionContext(out picture);
                ApplySelectionContext(selection);
                if (selection == null || !selection.HasSinglePicture || selection.IsManagedPicture)
                {
                    _userNotifier.ShowInfo(
                        selection != null && selection.IsManagedPicture
                            ? "当前图片已经绑定，无需重复绑定。"
                            : "请先只选中一个普通图片，再执行绑定。",
                        "DrawioWord");
                    return;
                }

                if (picture == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图片，再执行绑定。", "DrawioWord");
                    return;
                }

                string documentPath = _selectedPictureAccessor.GetDocumentFullName(_application);
                DiagramEnvelope envelope = _pictureMetadataService.CreateOrUpdateEnvelope(picture, _settings, documentPath);
                SaveManagedEnvelope(picture, envelope);
                _traceLog.Info("WordAddInHost", "Bound selected picture '" + picture.Name + "' to diagram " + envelope.DiagramId + ".");
                ApplySelectionContext(_pictureMetadataService.BuildSelectionContext(picture));

                string message =
                    "已为当前图片写入 Draw.io 元数据。" +
                    Environment.NewLine + "Picture: " + picture.Name +
                    Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                    Environment.NewLine + "Sidecar: " + envelope.SidecarPath;

                _userNotifier.ShowInfo(message, "DrawioWord");
            });
        }

        public void EditSelectedDiagram()
        {
            ExecuteGuarded("EditSelectedDiagram", "打开图形编辑器时出现异常。", true, delegate
            {
                WordPictureReference picture;
                SelectionContext selection = ReadLiveSelectionContext(out picture);
                ApplySelectionContext(selection);

                if (selection == null || !selection.IsManagedPicture)
                {
                    _userNotifier.ShowInfo("当前未选中插件管理的 Draw.io 图形。", "DrawioWord");
                    return;
                }

                if (picture == null)
                {
                    _userNotifier.ShowInfo("请先选中一个已绑定的图形。", "DrawioWord");
                    return;
                }

                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(picture, out envelope))
                {
                    _userNotifier.ShowInfo("当前图形缺少可读取的 Draw.io 元数据。", "DrawioWord");
                    return;
                }

                _traceLog.Info("WordAddInHost", "Editing diagram " + envelope.DiagramId + " from picture '" + picture.Name + "'.");
                if (!TryOpenPictureForEditing(picture, envelope, true))
                {
                    if (_settings.EditorMode == EditorMode.Desktop)
                    {
                        _userNotifier.ShowInfo("请先在设置中配置本地 draw.io/diagrams.net 路径。", "DrawioWord");
                    }

                    return;
                }

                ShowDiagramInfoDialog("已打开 Draw.io 图形编辑器。", envelope, picture, envelope.SidecarPath, true);
            });
        }

        public void ClearSelectedPictureBinding()
        {
            ExecuteGuarded("ClearSelectedPictureBinding", "清除图形绑定时出现异常。", true, delegate
            {
                WordPictureReference picture;
                SelectionContext selection = ReadLiveSelectionContext(out picture);
                ApplySelectionContext(selection);
                if (selection == null || !selection.IsManagedPicture)
                {
                    _userNotifier.ShowInfo("请先选中一个已绑定的图形，再执行清除绑定。", "DrawioWord");
                    return;
                }

                if (picture == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图片，再执行清除绑定。", "DrawioWord");
                    return;
                }

                DeleteManagedEnvelope(picture);
                _pictureMetadataService.Clear(picture);
                MaybeCleanupActiveDocument(true);
                _traceLog.Info("WordAddInHost", "Cleared binding for picture '" + picture.Name + "'.");
                ApplySelectionContext(_pictureMetadataService.BuildSelectionContext(picture));

                _userNotifier.ShowInfo("已清除当前图形上的 Draw.io 绑定信息。", "DrawioWord");
            });
        }

        public void RefreshSelectedDiagram()
        {
            ExecuteGuarded("RefreshSelectedDiagram", "刷新当前图形时出现异常。", true, delegate
            {
                WordPictureReference picture;
                SelectionContext selection = ReadLiveSelectionContext(out picture);
                ApplySelectionContext(selection);
                if (picture == null)
                {
                    _userNotifier.ShowInfo("请先选中一个图形。", "DrawioWord");
                    return;
                }

                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(picture, out envelope))
                {
                    _userNotifier.ShowInfo("当前图形没有可刷新的 Draw.io 元数据。", "DrawioWord");
                    return;
                }

                NormalizeSidecarPath(picture, envelope);
                if (string.IsNullOrWhiteSpace(envelope.SidecarPath) || !File.Exists(envelope.SidecarPath))
                {
                    _userNotifier.ShowInfo("当前图形还没有可读取的 .drawio 工作文件。", "DrawioWord");
                    return;
                }

                _traceLog.Info("WordAddInHost", "Refreshing diagram " + envelope.DiagramId + " from file " + envelope.SidecarPath + ".");
                RefreshDiagramFromFile(picture, envelope, envelope.SidecarPath, true);
            });
        }

        public void OpenSettings()
        {
            ExecuteGuarded("OpenSettings", "打开或保存设置时出现异常。", true, delegate
            {
                using (SettingsForm form = new SettingsForm(_settings, _desktopEditorPathDetector, false))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    _settings = form.Settings;
                    _settingsStore.Save(_settings);
                    _traceLog.Info("WordAddInHost", "Settings saved. EditorMode=" + _settings.EditorMode + ", EditorUrl=" + (_settings.EditorUrl ?? string.Empty) + ", DesktopPath=" + (_settings.DesktopEditorPath ?? string.Empty));
                    _userNotifier.ShowInfo("设置已保存。", "DrawioWord");
                }
            });
        }

        public string GetSelectionSummary()
        {
            return "对象：选中后点击操作";
        }

        public string GetSelectionDetailSummary()
        {
            return "状态：点击按钮时识别";
        }

        public string GetEditorModeSummary()
        {
            PluginSettings settings = _settings ?? new PluginSettings();
            string modeText = settings.EditorMode == EditorMode.Url ? "URL 模式" : "桌面模式";
            return "模式：" + modeText;
        }

        public string GetEditButtonLabel()
        {
            return "重新编辑";
        }

        private void OnSelectionDoubleClicked(object sender, SelectionDoubleClickEventArgs e)
        {
            ExecuteGuarded("OnSelectionDoubleClicked", "处理双击编辑时出现异常。", false, delegate
            {
                if (e == null || e.Selection == null)
                {
                    return;
                }

                WordPictureReference picture = _selectedPictureAccessor.GetSingleSelectedPicture(e.Selection);
                if (picture == null)
                {
                    return;
                }

                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(picture, out envelope))
                {
                    return;
                }

                if (TryOpenPictureForEditing(picture, envelope, true))
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

                WordPictureReference picture = _selectedPictureAccessor.FindPictureByDiagramId(_application, diagramId);
                if (picture == null)
                {
                    return;
                }

                DiagramEnvelope envelope;
                if (!TryReadManagedEnvelope(picture, out envelope))
                {
                    return;
                }

                RefreshDiagramFromFile(picture, envelope, filePath, false);
            });
        }

        private void RefreshDiagramFromFile(WordPictureReference picture, DiagramEnvelope envelope, string filePath, bool notifyUser)
        {
            if (picture == null || envelope == null || string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            string xmlContent;
            if (!TryReadWorkingFile(filePath, out xmlContent))
            {
                if (notifyUser)
                {
                    _userNotifier.ShowInfo("无法读取 .drawio 文件，请确认文件已保存且没有被编辑器占用。", "DrawioWord");
                }

                return;
            }

            envelope.DrawioXml = xmlContent;
            envelope.UpdatedUtc = DateTime.UtcNow;
            envelope.SidecarPath = filePath;

            string svgPath = _svgFileProvider.GetSvgPath(envelope, _settings, filePath);
            WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
            if (document == null)
            {
                return;
            }

            WordPictureReference newPicture = _svgPictureService.Replace(document, picture, svgPath);
            SaveManagedEnvelope(newPicture, envelope);
            ApplySelectionContext(_pictureMetadataService.BuildSelectionContext(newPicture));

            if (notifyUser)
            {
                _userNotifier.ShowInfo("已从 .drawio 文件刷新当前图形。", "DrawioWord");
            }
        }

        private bool TryOpenPictureForEditing(WordPictureReference picture, DiagramEnvelope envelope, bool explicitUserAction)
        {
            if (picture == null || envelope == null)
            {
                return false;
            }

            NormalizeSidecarPath(picture, envelope);
            if (_settings.EditorMode == EditorMode.Desktop)
            {
                string documentPath = _selectedPictureAccessor.GetDocumentFullName(_application);
                string workingFile = _desktopEditorLauncher.PrepareWorkingFile(envelope, _settings, documentPath);
                envelope.SidecarPath = workingFile;
                envelope.UpdatedUtc = DateTime.UtcNow;
                SaveManagedEnvelope(picture, envelope);
                _traceLog.Info("WordAddInHost", "Prepared desktop editor working file for diagram " + envelope.DiagramId + ": " + workingFile);

                if (!LaunchDesktopEditor(envelope, workingFile, explicitUserAction))
                {
                    return false;
                }

                return true;
            }

            if (_settings.EditorMode == EditorMode.Url)
            {
                if (string.IsNullOrWhiteSpace(_settings.EditorUrl))
                {
                    if (explicitUserAction)
                    {
                        _userNotifier.ShowInfo("请先在设置中配置 URL 模式编辑器地址。", "DrawioWord");
                    }

                    return false;
                }

                WriteSidecarFile(envelope);
                _traceLog.Info("WordAddInHost", "Opening URL editor for diagram " + envelope.DiagramId + " with URL " + _settings.EditorUrl + ".");
                using (UrlDiagramEditorForm form = new UrlDiagramEditorForm(_settings.EditorUrl, envelope.DiagramName, envelope.DrawioXml, _settings.UseOfficeCompatibleSvgLabels, _traceLog))
                {
                    form.DiagramSaved += delegate(object sender, UrlDiagramSavedEventArgs args)
                    {
                        ApplyUrlEditorSave(picture, envelope, args);
                    };

                    form.ShowDialog();
                }

                return true;
            }

            if (explicitUserAction)
            {
                _userNotifier.ShowInfo("当前编辑模式不受支持。", "DrawioWord");
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
                    _userNotifier.ShowInfo("请先在设置中配置本地 draw.io/diagrams.net 路径。", "DrawioWord");
                }

                _traceLog.Warn("WordAddInHost", "Desktop editor path is missing or invalid.");
                return false;
            }

            _traceLog.Info("WordAddInHost", "Launching desktop editor for diagram " + envelope.DiagramId + " using " + _settings.DesktopEditorPath + ".");
            _desktopEditorLauncher.Launch(_settings.DesktopEditorPath, workingFile);
            _desktopDiagramMonitor.Track(envelope.DiagramId, workingFile);
            return true;
        }

        private void ApplyUrlEditorSave(WordPictureReference originalPicture, DiagramEnvelope envelope, UrlDiagramSavedEventArgs args)
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
                _traceLog.Info("WordAddInHost", "Received URL editor save for diagram " + envelope.DiagramId + ". ExitRequested=" + args.ExitRequested);

                string svgPath = _svgMarkupFileStore.Write(envelope, args.SvgMarkup);
                WordPictureReference picture = _selectedPictureAccessor.FindPictureByDiagramId(_application, envelope.DiagramId) ?? originalPicture;
                if (picture == null)
                {
                    return;
                }

                WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
                if (document == null)
                {
                    return;
                }

                WordPictureReference newPicture = _svgPictureService.Replace(document, picture, svgPath);
                SaveManagedEnvelope(newPicture, envelope);
                ApplySelectionContext(new SelectionContext());
            }
            catch (Exception ex)
            {
                _traceLog.Error("WordAddInHost", "Failed to apply URL editor save for diagram " + envelope.DiagramId + ".", ex);
                _userNotifier.ShowInfo("URL 模式保存后刷新失败。" + Environment.NewLine + ex.Message, "DrawioWord");
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

            File.WriteAllText(envelope.SidecarPath, envelope.DrawioXml ?? string.Empty, Utf8WithoutBom);
        }

        private SelectionContext ReadLiveSelectionContext(out WordPictureReference picture)
        {
            picture = null;
            SelectionContext emptyContext = new SelectionContext();
            if (_application == null)
            {
                return emptyContext;
            }

            try
            {
                picture = _selectedPictureAccessor.GetSingleSelectedPicture(_application);
                if (picture == null)
                {
                    WordInterop.Selection selection = _selectedPictureAccessor.GetSelection(_application);
                    if (selection != null)
                    {
                        emptyContext.HasSelection = true;
                    }

                    return emptyContext;
                }

                return _pictureMetadataService.BuildSelectionContext(picture);
            }
            catch (COMException)
            {
                return emptyContext;
            }
            catch (InvalidComObjectException)
            {
                return emptyContext;
            }
        }

        private void ApplySelectionContext(SelectionContext context)
        {
            SelectionContext nextContext = context ?? new SelectionContext();
            bool changed = !AreSelectionContextsEquivalent(_currentSelection, nextContext);
            _currentSelection = nextContext;

            if (changed)
            {
                RaiseSelectionStateChanged();
                MaybeCleanupActiveDocument(false);
            }
        }

        private static bool AreSelectionContextsEquivalent(SelectionContext left, SelectionContext right)
        {
            SelectionContext normalizedLeft = left ?? new SelectionContext();
            SelectionContext normalizedRight = right ?? new SelectionContext();
            return
                normalizedLeft.HasSelection == normalizedRight.HasSelection &&
                normalizedLeft.HasSinglePicture == normalizedRight.HasSinglePicture &&
                normalizedLeft.IsManagedPicture == normalizedRight.IsManagedPicture &&
                string.Equals(normalizedLeft.ReferenceKey ?? string.Empty, normalizedRight.ReferenceKey ?? string.Empty, StringComparison.Ordinal) &&
                string.Equals(normalizedLeft.PictureName ?? string.Empty, normalizedRight.PictureName ?? string.Empty, StringComparison.Ordinal) &&
                string.Equals(normalizedLeft.DiagramId ?? string.Empty, normalizedRight.DiagramId ?? string.Empty, StringComparison.Ordinal) &&
                string.Equals(normalizedLeft.AlternativeText ?? string.Empty, normalizedRight.AlternativeText ?? string.Empty, StringComparison.Ordinal);
        }

        private bool TryReadManagedEnvelope(WordPictureReference picture, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (picture == null)
            {
                return false;
            }

            DiagramEnvelope fallbackEnvelope = null;
            bool hasFallbackEnvelope = _pictureMetadataService.TryRead(picture, out fallbackEnvelope);
            WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
            string diagramId = hasFallbackEnvelope ? fallbackEnvelope.DiagramId : string.Empty;

            DiagramEnvelope storedEnvelope;
            if (_documentDiagramStore.TryRead(document, diagramId, out storedEnvelope))
            {
                envelope = storedEnvelope;
                _pictureMetadataService.Save(picture, storedEnvelope);
                return true;
            }

            if (hasFallbackEnvelope)
            {
                envelope = fallbackEnvelope;
                SaveManagedEnvelope(picture, fallbackEnvelope);
                return true;
            }

            return false;
        }

        private void SaveManagedEnvelope(WordPictureReference picture, DiagramEnvelope envelope)
        {
            if (picture == null || envelope == null)
            {
                return;
            }

            WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
            if (document == null)
            {
                _traceLog.Info("WordAddInHost", "Skipped saving managed envelope because no document was available.");
                return;
            }

            string customXmlPartId = _documentDiagramStore.Upsert(document, envelope);
            if (string.IsNullOrWhiteSpace(customXmlPartId))
            {
                _traceLog.Info(
                    "WordAddInHost",
                    "Document metadata storage was unavailable; retained the full picture fallback for diagram " +
                    (envelope.DiagramId ?? string.Empty) + ".");
                _pictureMetadataService.SaveFallback(picture, envelope);
                return;
            }

            _pictureMetadataService.Save(picture, envelope);
        }

        private void DeleteManagedEnvelope(WordPictureReference picture)
        {
            if (picture == null)
            {
                return;
            }

            WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
            if (document == null)
            {
                return;
            }

            DiagramEnvelope envelope;
            if (_pictureMetadataService.TryRead(picture, out envelope))
            {
                _documentDiagramStore.Delete(document, envelope.DiagramId);
            }
        }

        private void NormalizeSidecarPath(WordPictureReference picture, DiagramEnvelope envelope)
        {
            if (picture == null || envelope == null || _settings == null)
            {
                return;
            }

            string documentPath = _selectedPictureAccessor.GetDocumentFullName(_application);
            string preferredPath = _desktopEditorLauncher.ResolvePreferredSidecarPath(envelope, _settings, documentPath);
            if (!string.IsNullOrWhiteSpace(preferredPath) &&
                !string.Equals(preferredPath, envelope.SidecarPath, StringComparison.OrdinalIgnoreCase))
            {
                envelope.SidecarPath = preferredPath;
                SaveManagedEnvelope(picture, envelope);
            }
        }

        private void MaybeCleanupActiveDocument(bool force)
        {
            try
            {
                WordInterop.Document document = _selectedPictureAccessor.GetActiveDocument(_application);
                if (document == null)
                {
                    return;
                }

                string documentPath = document.FullName ?? string.Empty;
                bool documentChanged = !string.Equals(_lastCleanupDocumentPath, documentPath, StringComparison.OrdinalIgnoreCase);
                if (!force &&
                    !documentChanged &&
                    DateTime.UtcNow.Subtract(_lastCleanupUtc).TotalSeconds < 30)
                {
                    return;
                }

                HashSet<string> liveDiagramIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _selectedPictureAccessor.CollectManagedDiagramReferences(document, liveDiagramIds);
                int deletedCount = _documentDiagramStore.DeleteOrphans(document, liveDiagramIds);
                _lastCleanupUtc = DateTime.UtcNow;
                _lastCleanupDocumentPath = documentPath;
                if (deletedCount > 0)
                {
                    _traceLog.Info("WordAddInHost", "Removed " + deletedCount + " orphaned CustomXMLParts from active document.");
                }
            }
            catch (COMException)
            {
            }
            catch (InvalidComObjectException)
            {
            }
            catch (Exception ex)
            {
                _traceLog.Error("WordAddInHost", "Failed while cleaning orphaned CustomXMLParts.", ex);
            }
        }

        private void ShowDiagramInfoDialog(string title, DiagramEnvelope envelope, WordPictureReference picture, string workingFile, bool editorLaunched)
        {
            if (_settings == null || !_settings.ShowDiagramInfoDialog || envelope == null)
            {
                return;
            }

            string message =
                title +
                Environment.NewLine + "Diagram ID: " + envelope.DiagramId +
                Environment.NewLine + "名称: " + envelope.DiagramName +
                Environment.NewLine + "模式: " + envelope.EditorMode +
                Environment.NewLine + "图片: " + (picture == null ? string.Empty : picture.Name) +
                Environment.NewLine + ".drawio: " + (workingFile ?? string.Empty);

            if (!editorLaunched)
            {
                message += Environment.NewLine + "当前未启动外部编辑器，图形先以 SVG 预览插入。";
            }

            _userNotifier.ShowInfo(message, "DrawioWord");
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
                _traceLog.Error("WordAddInHost", "Operation failed: " + (operationName ?? "Unknown"), ex);
                if (notifyUser)
                {
                    string message =
                        (string.IsNullOrWhiteSpace(userMessage) ? "操作执行失败。" : userMessage) +
                        Environment.NewLine + ex.Message +
                        Environment.NewLine + "日志: " + _traceLog.LogPath;
                    _userNotifier.ShowError(message, "DrawioWord");
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

        private static string SummarizePictureName(string pictureName)
        {
            string normalized = string.IsNullOrWhiteSpace(pictureName) ? "未命名图片" : pictureName.Trim();
            if (normalized.Length <= 18)
            {
                return normalized;
            }

            return normalized.Substring(0, 15) + "...";
        }
    }
}
