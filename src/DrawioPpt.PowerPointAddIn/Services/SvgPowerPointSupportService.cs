using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgPowerPointSupportService
    {
        private static readonly Regex DrawioWarningRegex = new Regex(
            "<switch>\\s*<g\\s+requiredFeatures=\"http://www\\.w3\\.org/TR/SVG11/feature#Extensibility\"\\s*/>\\s*<a[^>]*svg-export-text-problems[^>]*>.*?</a>\\s*</switch>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex DarkColorSchemeRegex = new Regex(
            "color-scheme\\s*:\\s*light\\s+dark\\s*;?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public void SanitizeFile(string svgPath)
        {
            if (string.IsNullOrWhiteSpace(svgPath) || !File.Exists(svgPath))
            {
                return;
            }

            string markup = File.ReadAllText(svgPath);
            string updatedMarkup = SanitizeMarkup(markup);
            if (!string.Equals(markup, updatedMarkup, StringComparison.Ordinal))
            {
                File.WriteAllText(svgPath, updatedMarkup, Encoding.UTF8);
            }
        }

        public string SanitizeMarkup(string svgMarkup)
        {
            if (string.IsNullOrWhiteSpace(svgMarkup))
            {
                return svgMarkup ?? string.Empty;
            }

            string updated = DarkColorSchemeRegex.Replace(svgMarkup, "color-scheme: light;");
            updated = DrawioWarningRegex.Replace(updated, string.Empty);
            return updated;
        }

        public bool TryReadSize(string svgPath, out float width, out float height)
        {
            width = 0f;
            height = 0f;

            if (string.IsNullOrWhiteSpace(svgPath) || !File.Exists(svgPath))
            {
                return false;
            }

            try
            {
                XmlDocument document = new XmlDocument();
                document.XmlResolver = null;

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Ignore;
                settings.XmlResolver = null;

                using (XmlReader reader = XmlReader.Create(svgPath, settings))
                {
                    document.Load(reader);
                }

                XmlElement root = document.DocumentElement;
                if (root == null || !string.Equals(root.LocalName, "svg", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (TryParseLength(root.GetAttribute("width"), out width) &&
                    TryParseLength(root.GetAttribute("height"), out height) &&
                    width > 0f &&
                    height > 0f)
                {
                    return true;
                }

                string viewBox = root.GetAttribute("viewBox");
                if (TryParseViewBox(viewBox, out width, out height))
                {
                    return true;
                }
            }
            catch
            {
            }

            width = 0f;
            height = 0f;
            return false;
        }

        public string PrepareForPresentation(string svgPath, float targetWidth, float targetHeight)
        {
            if (string.IsNullOrWhiteSpace(svgPath) ||
                !File.Exists(svgPath) ||
                targetWidth <= 0f ||
                targetHeight <= 0f)
            {
                return svgPath;
            }

            XmlDocument document;
            XmlElement root;
            if (!TryLoadRoot(svgPath, out document, out root))
            {
                return svgPath;
            }

            float sourceMinX;
            float sourceMinY;
            float sourceWidth;
            float sourceHeight;
            if (!TryReadCanvas(root, out sourceMinX, out sourceMinY, out sourceWidth, out sourceHeight) ||
                sourceWidth <= 0f ||
                sourceHeight <= 0f)
            {
                return svgPath;
            }

            float targetRatio = targetWidth / targetHeight;
            float sourceRatio = sourceWidth / sourceHeight;
            if (targetRatio <= 0f ||
                sourceRatio <= 0f ||
                Math.Abs(targetRatio - sourceRatio) < 0.01f)
            {
                return svgPath;
            }

            float newMinX = sourceMinX;
            float newMinY = sourceMinY;
            float newWidth = sourceWidth;
            float newHeight = sourceHeight;

            if (targetRatio > sourceRatio)
            {
                newWidth = sourceHeight * targetRatio;
                newMinX = sourceMinX - ((newWidth - sourceWidth) / 2f);
            }
            else
            {
                newHeight = sourceWidth / targetRatio;
                newMinY = sourceMinY - ((newHeight - sourceHeight) / 2f);
            }

            root.SetAttribute("viewBox", FormatFloat(newMinX) + " " + FormatFloat(newMinY) + " " + FormatFloat(newWidth) + " " + FormatFloat(newHeight));
            root.SetAttribute("width", FormatFloat(newWidth) + "px");
            root.SetAttribute("height", FormatFloat(newHeight) + "px");

            string targetDirectory = Path.Combine(Path.GetTempPath(), "DrawioPpt", "presentation-svg");
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            string outputPath = Path.Combine(
                targetDirectory,
                Path.GetFileNameWithoutExtension(svgPath) + "-" + Math.Round(targetWidth).ToString(CultureInfo.InvariantCulture) + "x" + Math.Round(targetHeight).ToString(CultureInfo.InvariantCulture) + ".svg");

            using (XmlWriter writer = XmlWriter.Create(outputPath, new XmlWriterSettings { Encoding = Encoding.UTF8, Indent = false }))
            {
                document.Save(writer);
            }

            return outputPath;
        }

        private static bool TryLoadRoot(string svgPath, out XmlDocument document, out XmlElement root)
        {
            document = null;
            root = null;

            try
            {
                document = new XmlDocument();
                document.XmlResolver = null;

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Ignore;
                settings.XmlResolver = null;

                using (XmlReader reader = XmlReader.Create(svgPath, settings))
                {
                    document.Load(reader);
                }

                root = document.DocumentElement;
                return root != null && string.Equals(root.LocalName, "svg", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                document = null;
                root = null;
                return false;
            }
        }

        private static bool TryReadCanvas(XmlElement root, out float minX, out float minY, out float width, out float height)
        {
            minX = 0f;
            minY = 0f;
            width = 0f;
            height = 0f;

            if (root == null)
            {
                return false;
            }

            string viewBox = root.GetAttribute("viewBox");
            if (!string.IsNullOrWhiteSpace(viewBox))
            {
                string[] parts = viewBox.Split(new char[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 4 &&
                    float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out minX) &&
                    float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out minY) &&
                    float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out width) &&
                    float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out height) &&
                    width > 0f &&
                    height > 0f)
                {
                    return true;
                }
            }

            return TryParseLength(root.GetAttribute("width"), out width) &&
                   TryParseLength(root.GetAttribute("height"), out height) &&
                   width > 0f &&
                   height > 0f;
        }

        private static string FormatFloat(float value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static bool TryParseLength(string value, out float result)
        {
            result = 0f;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            int length = 0;
            while (length < trimmed.Length)
            {
                char current = trimmed[length];
                if ((current >= '0' && current <= '9') ||
                    current == '.' ||
                    current == '+' ||
                    current == '-' ||
                    current == 'e' ||
                    current == 'E')
                {
                    length++;
                    continue;
                }

                break;
            }

            if (length <= 0)
            {
                return false;
            }

            return float.TryParse(
                trimmed.Substring(0, length),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }

        private static bool TryParseViewBox(string value, out float width, out float height)
        {
            width = 0f;
            height = 0f;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string[] parts = value.Split(new char[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
            {
                return false;
            }

            float parsedWidth;
            float parsedHeight;
            if (!float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out parsedWidth) ||
                !float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out parsedHeight))
            {
                return false;
            }

            if (parsedWidth <= 0f || parsedHeight <= 0f)
            {
                return false;
            }

            width = parsedWidth;
            height = parsedHeight;
            return true;
        }
    }
}
