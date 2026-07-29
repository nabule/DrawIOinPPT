using System;
using System.IO;
using System.Text;
using DrawioPpt.Core.Models;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.WordAddIn.Services
{
    public class WordPreviewImageProvider
    {
        private const int PreviewWidth = 620;
        private readonly DesktopSvgExporter exporter;

        public WordPreviewImageProvider()
            : this(new DesktopSvgExporter())
        {
        }

        public WordPreviewImageProvider(DesktopSvgExporter exporter)
        {
            if (exporter == null)
            {
                throw new ArgumentNullException("exporter");
            }

            this.exporter = exporter;
        }

        public string GetPreviewImagePath(
            DiagramEnvelope envelope,
            string editorExecutablePath,
            string fallbackSvgPath)
        {
            if (envelope == null ||
                string.IsNullOrWhiteSpace(envelope.DrawioXml))
            {
                return fallbackSvgPath;
            }

            string previewDirectory = Path.Combine(
                Path.GetTempPath(),
                "DrawioPpt",
                "word-preview",
                SanitizeDiagramId(envelope.DiagramId) + "-" +
                    Guid.NewGuid().ToString("N"));
            string drawioFilePath = Path.Combine(
                previewDirectory,
                "source.drawio");
            string outputPngPath = Path.Combine(
                previewDirectory,
                "preview.png");
            bool exported = false;

            try
            {
                Directory.CreateDirectory(previewDirectory);
                File.WriteAllText(
                    drawioFilePath,
                    envelope.DrawioXml,
                    new UTF8Encoding(false));
                exported = this.exporter.TryExportPng(
                    editorExecutablePath,
                    drawioFilePath,
                    outputPngPath,
                    PreviewWidth);
            }
            catch
            {
                exported = false;
            }
            finally
            {
                TryDeleteFile(drawioFilePath);
                if (File.Exists(drawioFilePath))
                {
                    exported = false;
                }

                if (!exported)
                {
                    TryDeleteFile(outputPngPath);
                    TryDeleteEmptyDirectory(previewDirectory);
                }
            }

            return exported ? outputPngPath : fallbackSvgPath;
        }

        private static string SanitizeDiagramId(string diagramId)
        {
            if (string.IsNullOrWhiteSpace(diagramId))
            {
                return "diagram";
            }

            StringBuilder builder = new StringBuilder();
            int index;
            for (index = 0;
                 index < diagramId.Length && builder.Length < 48;
                 index++)
            {
                char current = diagramId[index];
                if ((current >= 'a' && current <= 'z') ||
                    (current >= 'A' && current <= 'Z') ||
                    (current >= '0' && current <= '9') ||
                    current == '-' ||
                    current == '_')
                {
                    builder.Append(current);
                }
                else
                {
                    builder.Append('_');
                }
            }

            string sanitized = builder.ToString().Trim('_');
            return sanitized.Length == 0 ? "diagram" : sanitized;
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
            }
        }

        private static void TryDeleteEmptyDirectory(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath))
                {
                    Directory.Delete(directoryPath, false);
                }
            }
            catch
            {
            }
        }
    }
}
