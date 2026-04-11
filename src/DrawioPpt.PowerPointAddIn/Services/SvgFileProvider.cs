using System;
using System.IO;
using DrawioPpt.Core.Models;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgFileProvider
    {
        private readonly DesktopSvgExporter _desktopSvgExporter;
        private readonly DiagramVisualSvgService _visualSvgService;
        private readonly SvgDiagramDataEmbedder _embedder;

        public SvgFileProvider()
            : this(new DesktopSvgExporter(), new DiagramVisualSvgService(), new SvgDiagramDataEmbedder())
        {
        }

        public SvgFileProvider(DesktopSvgExporter desktopSvgExporter, DiagramVisualSvgService visualSvgService, SvgDiagramDataEmbedder embedder)
        {
            _desktopSvgExporter = desktopSvgExporter;
            _visualSvgService = visualSvgService;
            _embedder = embedder;
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
                    if (_embedder != null)
                    {
                        _embedder.EnsureEmbeddedFile(outputSvgPath, envelope.DrawioXml);
                    }

                    return outputSvgPath;
                }
            }

            string previewSvgPath = _visualSvgService.WritePreviewSvg(envelope);
            if (_embedder != null)
            {
                _embedder.EnsureEmbeddedFile(previewSvgPath, envelope.DrawioXml);
            }

            return previewSvgPath;
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
