using System;
using System.IO;
using System.Security;
using System.Text.RegularExpressions;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgDiagramDataEmbedder
    {
        private static readonly Regex ContentAttributeRegex = new Regex("\\scontent\\s*=\\s*(\"[^\"]*\"|'[^']*')", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline);

        public string EnsureEmbedded(string svgMarkup, string drawioXml)
        {
            if (string.IsNullOrWhiteSpace(svgMarkup) || string.IsNullOrWhiteSpace(drawioXml))
            {
                return svgMarkup ?? string.Empty;
            }

            int svgIndex = svgMarkup.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
            if (svgIndex < 0)
            {
                return svgMarkup;
            }

            int tagEndIndex = FindTagEnd(svgMarkup, svgIndex);
            if (tagEndIndex < 0)
            {
                return svgMarkup;
            }

            string rootStartTag = svgMarkup.Substring(svgIndex, tagEndIndex - svgIndex + 1);
            string escapedXml = SecurityElement.Escape(drawioXml) ?? string.Empty;
            string replacementAttribute = " content=\"" + escapedXml + "\"";
            string updatedRootStartTag;

            if (ContentAttributeRegex.IsMatch(rootStartTag))
            {
                updatedRootStartTag = ContentAttributeRegex.Replace(rootStartTag, replacementAttribute, 1);
            }
            else
            {
                int insertIndex = rootStartTag.Length - 1;
                if (insertIndex > 0 && rootStartTag[insertIndex - 1] == '/')
                {
                    insertIndex--;
                }

                updatedRootStartTag = rootStartTag.Insert(insertIndex, replacementAttribute);
            }

            if (string.Equals(rootStartTag, updatedRootStartTag, StringComparison.Ordinal))
            {
                return svgMarkup;
            }

            return svgMarkup.Substring(0, svgIndex) + updatedRootStartTag + svgMarkup.Substring(tagEndIndex + 1);
        }

        public void EnsureEmbeddedFile(string svgPath, string drawioXml)
        {
            if (string.IsNullOrWhiteSpace(svgPath) || string.IsNullOrWhiteSpace(drawioXml) || !File.Exists(svgPath))
            {
                return;
            }

            string svgMarkup = File.ReadAllText(svgPath);
            string updatedMarkup = EnsureEmbedded(svgMarkup, drawioXml);
            if (!string.Equals(svgMarkup, updatedMarkup, StringComparison.Ordinal))
            {
                File.WriteAllText(svgPath, updatedMarkup);
            }
        }

        private static int FindTagEnd(string text, int startIndex)
        {
            bool insideDoubleQuotes = false;
            bool insideSingleQuotes = false;

            int index;
            for (index = startIndex; index < text.Length; index++)
            {
                char current = text[index];
                if (current == '"' && !insideSingleQuotes)
                {
                    insideDoubleQuotes = !insideDoubleQuotes;
                    continue;
                }

                if (current == '\'' && !insideDoubleQuotes)
                {
                    insideSingleQuotes = !insideSingleQuotes;
                    continue;
                }

                if (current == '>' && !insideDoubleQuotes && !insideSingleQuotes)
                {
                    return index;
                }
            }

            return -1;
        }
    }
}
