using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Threading;
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
            QueueStalePreviewCleanup();
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
                bool sourceRemoved = TryDeleteFileWithRetries(
                    drawioFilePath,
                    4,
                    50);
                if (!exported)
                {
                    exported = TryCreateFallbackPreview(
                        fallbackSvgPath,
                        outputPngPath);
                }

                if (!exported)
                {
                    TryCleanupPreviewDirectory(previewDirectory);
                    if (Directory.Exists(previewDirectory))
                    {
                        QueueDirectoryCleanup(previewDirectory);
                    }
                }
                else if (!sourceRemoved)
                {
                    QueueSourceFileCleanup(drawioFilePath);
                }
            }

            return exported ? outputPngPath : fallbackSvgPath;
        }

        public void CleanupPreviewImage(string previewImagePath)
        {
            if (string.IsNullOrWhiteSpace(previewImagePath) ||
                !string.Equals(
                    Path.GetExtension(previewImagePath),
                    ".png",
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string previewRoot = Path.GetFullPath(
                Path.Combine(
                    Path.GetTempPath(),
                    "DrawioPpt",
                    "word-preview"));
            string previewRootPrefix = previewRoot.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar) +
                Path.DirectorySeparatorChar;
            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(previewImagePath);
            }
            catch
            {
                return;
            }

            if (!fullPath.StartsWith(
                    previewRootPrefix,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    Path.GetFileName(fullPath),
                    "preview.png",
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string previewDirectory = Path.GetDirectoryName(fullPath);
            TryCleanupPreviewDirectory(previewDirectory);
            if (Directory.Exists(previewDirectory))
            {
                QueueDirectoryCleanup(previewDirectory);
            }
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

        private static bool TryCreateFallbackPreview(
            string fallbackSvgPath,
            string outputPngPath)
        {
            try
            {
                float sourceWidth;
                float sourceHeight;
                SvgPowerPointSupportService svgSupportService =
                    new SvgPowerPointSupportService();
                bool hasSourceSize =
                    svgSupportService.TryReadSize(
                        fallbackSvgPath,
                        out sourceWidth,
                        out sourceHeight) &&
                    sourceWidth > 0f &&
                    sourceHeight > 0f;
                float aspectRatio = hasSourceSize
                    ? sourceWidth / sourceHeight
                    : 1.6f;
                int pixelWidth = PreviewWidth;
                int pixelHeight = Math.Max(
                    1,
                    (int)Math.Round(
                        pixelWidth / aspectRatio));
                if (pixelHeight > 1200)
                {
                    pixelHeight = 1200;
                    pixelWidth = Math.Max(
                        1,
                        (int)Math.Round(
                            pixelHeight * aspectRatio));
                }

                if (pixelHeight < 48)
                {
                    pixelHeight = 48;
                    pixelWidth = Math.Max(
                        1,
                        Math.Min(
                            1200,
                            (int)Math.Round(
                                pixelHeight * aspectRatio)));
                }

                using (Bitmap bitmap =
                    new Bitmap(pixelWidth, pixelHeight))
                using (Graphics graphics =
                    Graphics.FromImage(bitmap))
                using (Pen borderPen =
                    new Pen(
                        Color.FromArgb(91, 155, 213),
                        2f))
                using (Brush textBrush =
                    new SolidBrush(
                        Color.FromArgb(44, 62, 80)))
                using (StringFormat textFormat =
                    new StringFormat())
                {
                    graphics.SmoothingMode =
                        SmoothingMode.HighQuality;
                    graphics.Clear(
                        Color.FromArgb(236, 244, 251));
                    graphics.DrawRectangle(
                        borderPen,
                        1,
                        1,
                        Math.Max(1, pixelWidth - 3),
                        Math.Max(1, pixelHeight - 3));
                    textFormat.Alignment =
                        StringAlignment.Center;
                    textFormat.LineAlignment =
                        StringAlignment.Center;
                    graphics.DrawString(
                        "Draw.io diagram\r\nClick Edit to open",
                        SystemFonts.MessageBoxFont,
                        textBrush,
                        new RectangleF(
                            8,
                            8,
                            Math.Max(1, pixelWidth - 16),
                            Math.Max(1, pixelHeight - 16)),
                        textFormat);
                    bitmap.Save(
                        outputPngPath,
                        System.Drawing.Imaging.ImageFormat.Png);
                }

                return File.Exists(outputPngPath) &&
                    new FileInfo(outputPngPath).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryDeleteFileWithRetries(
            string filePath,
            int attemptCount,
            int delayMilliseconds)
        {
            int attempt;
            for (attempt = 0; attempt < attemptCount; attempt++)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        return true;
                    }

                    File.Delete(filePath);
                    if (!File.Exists(filePath))
                    {
                        return true;
                    }
                }
                catch
                {
                }

                if (attempt + 1 < attemptCount &&
                    delayMilliseconds > 0)
                {
                    Thread.Sleep(delayMilliseconds);
                }
            }

            return !File.Exists(filePath);
        }

        private static void TryCleanupPreviewDirectory(
            string directoryPath)
        {
            if (!IsOwnedPreviewDirectory(directoryPath))
            {
                return;
            }

            TryDeleteFileWithRetries(
                Path.Combine(directoryPath, "source.drawio"),
                4,
                50);
            TryDeleteFileWithRetries(
                Path.Combine(directoryPath, "preview.png"),
                4,
                50);
            TryDeleteEmptyDirectory(directoryPath);
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

        private static bool IsOwnedPreviewDirectory(
            string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                return false;
            }

            try
            {
                string previewRoot = Path.GetFullPath(
                    Path.Combine(
                        Path.GetTempPath(),
                        "DrawioPpt",
                        "word-preview"));
                string fullDirectory = Path.GetFullPath(directoryPath);
                string parent = Path.GetDirectoryName(fullDirectory);
                string name = Path.GetFileName(fullDirectory);
                if (!string.Equals(
                        parent,
                        previewRoot,
                        StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(name) ||
                    name.Length < 33)
                {
                    return false;
                }

                string suffix = name.Substring(name.Length - 32);
                Guid parsed;
                return Guid.TryParseExact(suffix, "N", out parsed);
            }
            catch
            {
                return false;
            }
        }

        private static void QueueDirectoryCleanup(
            string directoryPath)
        {
            if (!IsOwnedPreviewDirectory(directoryPath))
            {
                return;
            }

            ThreadPool.QueueUserWorkItem(
                delegate
                {
                    int attempt;
                    for (attempt = 0; attempt < 20; attempt++)
                    {
                        TryCleanupPreviewDirectory(directoryPath);
                        if (!Directory.Exists(directoryPath))
                        {
                            return;
                        }

                        Thread.Sleep(250);
                    }
                });
        }

        private static void QueueSourceFileCleanup(
            string sourceFilePath)
        {
            if (string.IsNullOrWhiteSpace(sourceFilePath) ||
                !string.Equals(
                    Path.GetFileName(sourceFilePath),
                    "source.drawio",
                    StringComparison.OrdinalIgnoreCase) ||
                !IsOwnedPreviewDirectory(
                    Path.GetDirectoryName(sourceFilePath)))
            {
                return;
            }

            ThreadPool.QueueUserWorkItem(
                delegate
                {
                    int attempt;
                    for (attempt = 0; attempt < 20; attempt++)
                    {
                        if (TryDeleteFileWithRetries(
                                sourceFilePath,
                                1,
                                0))
                        {
                            return;
                        }

                        Thread.Sleep(250);
                    }
                });
        }

        private static void QueueStalePreviewCleanup()
        {
            ThreadPool.QueueUserWorkItem(
                delegate
                {
                    string previewRoot = Path.Combine(
                        Path.GetTempPath(),
                        "DrawioPpt",
                        "word-preview");
                    try
                    {
                        if (!Directory.Exists(previewRoot))
                        {
                            return;
                        }

                        foreach (string directoryPath in
                            Directory.GetDirectories(previewRoot))
                        {
                            try
                            {
                                if (Directory.GetLastWriteTimeUtc(
                                        directoryPath) <=
                                    DateTime.UtcNow.AddHours(-1))
                                {
                                    TryCleanupPreviewDirectory(
                                        directoryPath);
                                }
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                });
        }
    }
}
