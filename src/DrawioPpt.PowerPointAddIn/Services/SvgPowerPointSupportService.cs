using System;
using System.Collections.Generic;
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
            updated = ReplaceHtmlLabelFallbacks(updated);
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

        private static string ReplaceHtmlLabelFallbacks(string svgMarkup)
        {
            XmlDocument document;
            XmlElement root;
            if (!TryLoadMarkupRoot(svgMarkup, out document, out root))
            {
                return svgMarkup;
            }

            XmlNodeList switchNodes = root.GetElementsByTagName("switch");
            if (switchNodes == null || switchNodes.Count == 0)
            {
                return svgMarkup;
            }

            List<XmlElement> targets = new List<XmlElement>();
            int index;
            for (index = 0; index < switchNodes.Count; index++)
            {
                XmlElement switchElement = switchNodes[index] as XmlElement;
                if (switchElement != null)
                {
                    targets.Add(switchElement);
                }
            }

            bool changed = false;
            for (index = targets.Count - 1; index >= 0; index--)
            {
                XmlElement replacement;
                if (!TryCreateVectorTextFallback(targets[index], out replacement))
                {
                    continue;
                }

                XmlNode parentNode = targets[index].ParentNode;
                if (parentNode == null)
                {
                    continue;
                }

                parentNode.ReplaceChild(replacement, targets[index]);
                changed = true;
            }

            return changed ? SaveMarkup(document) : svgMarkup;
        }

        private static bool TryLoadMarkupRoot(string svgMarkup, out XmlDocument document, out XmlElement root)
        {
            document = null;
            root = null;

            if (string.IsNullOrWhiteSpace(svgMarkup))
            {
                return false;
            }

            try
            {
                document = new XmlDocument();
                document.XmlResolver = null;

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Ignore;
                settings.XmlResolver = null;

                using (StringReader stringReader = new StringReader(svgMarkup))
                using (XmlReader reader = XmlReader.Create(stringReader, settings))
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

        private static bool TryCreateVectorTextFallback(XmlElement switchElement, out XmlElement textElement)
        {
            textElement = null;

            if (switchElement == null)
            {
                return false;
            }

            XmlElement foreignObject = FindFirstDescendant(switchElement, "foreignObject");
            XmlElement image = FindFirstDescendant(switchElement, "image");
            if (foreignObject == null || image == null)
            {
                return false;
            }

            List<string> lines = ExtractTextLines(foreignObject);
            if (lines.Count == 0)
            {
                return false;
            }

            float imageX;
            float imageY;
            float imageWidth;
            float imageHeight;
            if (!TryParseLength(image.GetAttribute("x"), out imageX) ||
                !TryParseLength(image.GetAttribute("y"), out imageY) ||
                !TryParseLength(image.GetAttribute("width"), out imageWidth) ||
                !TryParseLength(image.GetAttribute("height"), out imageHeight) ||
                imageWidth <= 0f ||
                imageHeight <= 0f)
            {
                return false;
            }

            XmlElement textAnchor = FindTextAnchorElement(foreignObject);
            LabelStyle style = ReadLabelStyle(textAnchor ?? foreignObject);
            if (style.FontSize <= 0f)
            {
                style.FontSize = 12f;
            }

            if (style.LineAdvance <= 0f)
            {
                style.LineAdvance = style.FontSize * 1.2f;
            }

            if (string.IsNullOrWhiteSpace(style.Color))
            {
                style.Color = "#000000";
            }

            if (string.IsNullOrWhiteSpace(style.FontFamily))
            {
                style.FontFamily = "Helvetica";
            }

            XmlDocument document = switchElement.OwnerDocument;
            if (document == null)
            {
                return false;
            }

            float totalHeight = style.FontSize + ((lines.Count - 1) * style.LineAdvance);
            float startY = imageY + ((imageHeight - totalHeight) / 2f) + (style.FontSize * 0.85f);
            float anchorX = imageX;
            string textAnchorValue = "start";

            if (string.Equals(style.HorizontalAlignment, "middle", StringComparison.Ordinal))
            {
                anchorX = imageX + (imageWidth / 2f);
                textAnchorValue = "middle";
            }
            else if (string.Equals(style.HorizontalAlignment, "end", StringComparison.Ordinal))
            {
                anchorX = imageX + imageWidth;
                textAnchorValue = "end";
            }

            textElement = document.CreateElement("text", switchElement.NamespaceURI);
            textElement.SetAttribute("fill", NormalizeColor(style.Color));
            textElement.SetAttribute("font-size", FormatFloat(style.FontSize) + "px");
            textElement.SetAttribute("font-family", style.FontFamily);
            textElement.SetAttribute("text-anchor", textAnchorValue);
            textElement.SetAttribute("pointer-events", "none");

            if (style.Bold)
            {
                textElement.SetAttribute("font-weight", "bold");
            }

            if (style.Italic)
            {
                textElement.SetAttribute("font-style", "italic");
            }

            int index;
            for (index = 0; index < lines.Count; index++)
            {
                XmlElement tspan = document.CreateElement("tspan", switchElement.NamespaceURI);
                tspan.SetAttribute("x", FormatFloat(anchorX));

                if (index == 0)
                {
                    tspan.SetAttribute("y", FormatFloat(startY));
                }
                else
                {
                    tspan.SetAttribute("dy", FormatFloat(style.LineAdvance));
                }

                tspan.InnerText = lines[index];
                textElement.AppendChild(tspan);
            }

            return textElement.ChildNodes.Count > 0;
        }

        private static XmlElement FindFirstDescendant(XmlElement parent, string localName)
        {
            if (parent == null || string.IsNullOrWhiteSpace(localName))
            {
                return null;
            }

            XmlNodeList nodes = parent.GetElementsByTagName(localName);
            return nodes != null && nodes.Count > 0 ? nodes[0] as XmlElement : null;
        }

        private static XmlElement FindTextAnchorElement(XmlNode node)
        {
            if (node == null)
            {
                return null;
            }

            if (node.NodeType == XmlNodeType.Text || node.NodeType == XmlNodeType.CDATA)
            {
                return string.IsNullOrWhiteSpace(node.Value) ? null : node.ParentNode as XmlElement;
            }

            XmlNode childNode = node.FirstChild;
            while (childNode != null)
            {
                XmlElement candidate = FindTextAnchorElement(childNode);
                if (candidate != null)
                {
                    return candidate;
                }

                childNode = childNode.NextSibling;
            }

            return null;
        }

        private static List<string> ExtractTextLines(XmlNode foreignObject)
        {
            List<string> lines = new List<string>();
            StringBuilder currentLine = new StringBuilder();
            AppendTextLines(foreignObject, lines, currentLine);
            FinalizeCurrentLine(lines, currentLine);

            List<string> trimmedLines = new List<string>();
            int index;
            for (index = 0; index < lines.Count; index++)
            {
                string value = lines[index] == null ? string.Empty : lines[index].Trim();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    trimmedLines.Add(value);
                }
            }

            return trimmedLines;
        }

        private static void AppendTextLines(XmlNode node, List<string> lines, StringBuilder currentLine)
        {
            if (node == null)
            {
                return;
            }

            if (node.NodeType == XmlNodeType.Text || node.NodeType == XmlNodeType.CDATA)
            {
                if (!string.IsNullOrWhiteSpace(node.Value))
                {
                    currentLine.Append(node.Value);
                }

                return;
            }

            XmlElement element = node as XmlElement;
            if (element != null && string.Equals(element.LocalName, "br", StringComparison.OrdinalIgnoreCase))
            {
                FinalizeCurrentLine(lines, currentLine);
                return;
            }

            XmlNode childNode = node.FirstChild;
            while (childNode != null)
            {
                AppendTextLines(childNode, lines, currentLine);
                childNode = childNode.NextSibling;
            }

            if (element != null &&
                (string.Equals(element.LocalName, "p", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(element.LocalName, "li", StringComparison.OrdinalIgnoreCase)))
            {
                FinalizeCurrentLine(lines, currentLine);
            }
        }

        private static void FinalizeCurrentLine(List<string> lines, StringBuilder currentLine)
        {
            if (lines == null || currentLine == null)
            {
                return;
            }

            string value = currentLine.ToString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                lines.Add(value);
            }

            currentLine.Length = 0;
        }

        private static LabelStyle ReadLabelStyle(XmlElement textElement)
        {
            LabelStyle style = new LabelStyle();
            style.FontFamily = FindStyleValue(textElement, "font-family");
            style.Color = FindStyleValue(textElement, "color");
            style.HorizontalAlignment = NormalizeAlignment(
                FindStyleValue(textElement, "text-align"),
                FindStyleValue(textElement, "justify-content"));

            float fontSize;
            if (TryParseLength(FindStyleValue(textElement, "font-size"), out fontSize) && fontSize > 0f)
            {
                style.FontSize = fontSize;
            }

            style.LineAdvance = ParseLineAdvance(FindStyleValue(textElement, "line-height"), style.FontSize);
            style.Bold = HasFontWeight(textElement);
            style.Italic = HasFontStyle(textElement);
            return style;
        }

        private static string FindStyleValue(XmlElement element, string propertyName)
        {
            XmlNode currentNode = element;
            while (currentNode != null)
            {
                XmlElement currentElement = currentNode as XmlElement;
                if (currentElement != null)
                {
                    string style = currentElement.GetAttribute("style");
                    string value;
                    if (TryGetStyleValue(style, propertyName, out value))
                    {
                        return value;
                    }
                }

                currentNode = currentNode.ParentNode;
            }

            return string.Empty;
        }

        private static bool TryGetStyleValue(string style, string propertyName, out string value)
        {
            value = string.Empty;
            if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            string[] parts = style.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            int index;
            for (index = 0; index < parts.Length; index++)
            {
                string part = parts[index];
                int separatorIndex = part.IndexOf(':');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                string name = part.Substring(0, separatorIndex).Trim();
                if (!string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                value = part.Substring(separatorIndex + 1).Trim();
                return !string.IsNullOrWhiteSpace(value);
            }

            return false;
        }

        private static string NormalizeAlignment(string textAlign, string justifyContent)
        {
            string alignment = string.IsNullOrWhiteSpace(textAlign) ? justifyContent : textAlign;
            if (string.IsNullOrWhiteSpace(alignment))
            {
                return "start";
            }

            string normalized = alignment.Trim().ToLowerInvariant();
            if (normalized.IndexOf("center", StringComparison.Ordinal) >= 0)
            {
                return "middle";
            }

            if (normalized.IndexOf("right", StringComparison.Ordinal) >= 0 ||
                normalized.IndexOf("end", StringComparison.Ordinal) >= 0)
            {
                return "end";
            }

            return "start";
        }

        private static bool HasFontWeight(XmlElement element)
        {
            XmlNode currentNode = element;
            while (currentNode != null)
            {
                XmlElement currentElement = currentNode as XmlElement;
                if (currentElement != null)
                {
                    if (string.Equals(currentElement.LocalName, "b", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(currentElement.LocalName, "strong", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    string fontWeight = FindStyleValue(currentElement, "font-weight");
                    if (string.Equals(fontWeight, "bold", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    float parsedWeight;
                    if (float.TryParse(fontWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out parsedWeight) &&
                        parsedWeight >= 600f)
                    {
                        return true;
                    }
                }

                currentNode = currentNode.ParentNode;
            }

            return false;
        }

        private static bool HasFontStyle(XmlElement element)
        {
            XmlNode currentNode = element;
            while (currentNode != null)
            {
                XmlElement currentElement = currentNode as XmlElement;
                if (currentElement != null)
                {
                    if (string.Equals(currentElement.LocalName, "i", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(currentElement.LocalName, "em", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    string fontStyle = FindStyleValue(currentElement, "font-style");
                    if (string.Equals(fontStyle, "italic", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fontStyle, "oblique", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                currentNode = currentNode.ParentNode;
            }

            return false;
        }

        private static float ParseLineAdvance(string lineHeightValue, float fontSize)
        {
            if (fontSize <= 0f)
            {
                fontSize = 12f;
            }

            if (string.IsNullOrWhiteSpace(lineHeightValue))
            {
                return fontSize * 1.2f;
            }

            string trimmed = lineHeightValue.Trim();
            float numericValue;
            if (!TryParseLength(trimmed, out numericValue) || numericValue <= 0f)
            {
                return fontSize * 1.2f;
            }

            if (trimmed.EndsWith("%", StringComparison.Ordinal))
            {
                return fontSize * (numericValue / 100f);
            }

            if (trimmed.EndsWith("px", StringComparison.OrdinalIgnoreCase) ||
                trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                return numericValue;
            }

            return numericValue * fontSize;
        }

        private static string NormalizeColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color))
            {
                return "#000000";
            }

            string trimmed = color.Trim();
            if (trimmed.IndexOf("light-dark(", StringComparison.OrdinalIgnoreCase) == 0)
            {
                int startIndex = trimmed.IndexOf('(');
                int commaIndex = trimmed.IndexOf(',', startIndex + 1);
                if (startIndex >= 0 && commaIndex > startIndex)
                {
                    return trimmed.Substring(startIndex + 1, commaIndex - startIndex - 1).Trim();
                }
            }

            return trimmed;
        }

        private static string SaveMarkup(XmlDocument document)
        {
            if (document == null)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.Indent = false;
            settings.OmitXmlDeclaration = true;

            using (XmlWriter writer = XmlWriter.Create(builder, settings))
            {
                document.Save(writer);
            }

            return builder.ToString();
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

        private sealed class LabelStyle
        {
            public bool Bold;
            public string Color;
            public string FontFamily;
            public float FontSize;
            public string HorizontalAlignment;
            public bool Italic;
            public float LineAdvance;
        }
    }
}
