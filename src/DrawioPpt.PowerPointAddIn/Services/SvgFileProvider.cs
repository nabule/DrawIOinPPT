using System;
using System.IO;
using DrawioPpt.Core.Models;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgFileProvider
    {
        private readonly DesktopSvgExporter _desktopSvgExporter;
        private readonly DiagramVisualSvgService _visualSvgService;

        public SvgFileProvider()
            : this(new DesktopSvgExporter(), new DiagramVisualSvgService())
        {
        }

        public SvgFileProvider(DesktopSvgExporter desktopSvgExporter, DiagramVisualSvgService visualSvgService)
        {
            _desktopSvgExporter = desktopSvgExporter;
            _visualSvgService = visualSvgService;
        }

        public string GetSvgPath(DiagramEnvelope envelope, PluginSettings settings, string drawioFilePath)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            if (settings != null &&
                settings.EditorMode == EditorMode.Desktop &&
                !string.IsNullOrWhiteSpace(settings.DesktopEditorPath) &&
                !string.IsNullOrWhiteSpace(drawioFilePath))
            {
                string outputSvgPath = BuildDesktopExportPath(envelope.DiagramId);
                if (_desktopSvgExporter.TryExport(settings.DesktopEditorPath, drawioFilePath, outputSvgPath))
                {
                    return outputSvgPath;
                }
            }

            return _visualSvgService.WritePreviewSvg(envelope);
        }

        private static string BuildDesktopExportPath(string diagramId)
        {
            string targetDirectory = Path.Combine(Path.GetTempPath(), "DrawioPpt", "export");
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            string safeId = string.IsNullOrWhiteSpace(diagramId) ? Guid.NewGuid().ToString("N") : diagramId;
            return Path.Combine(targetDirectory, safeId + ".svg");
        }
    }
}

