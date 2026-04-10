using System;
using System.Globalization;
using System.Xml.Linq;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;

namespace DrawioPpt.Core.Services
{
    public class DiagramEnvelopeSerializer : IDiagramEnvelopeSerializer
    {
        public bool CanDeserialize(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return false;
            }

            try
            {
                XDocument document = XDocument.Parse(content);
                return document.Root != null && document.Root.Name.LocalName == KnownMetadata.EnvelopeRootName;
            }
            catch
            {
                return false;
            }
        }

        public DiagramEnvelope Deserialize(string content)
        {
            if (!this.CanDeserialize(content))
            {
                throw new InvalidOperationException("Content is not a DrawioPpt envelope.");
            }

            XDocument document = XDocument.Parse(content);
            XElement root = document.Root;
            DiagramEnvelope envelope = new DiagramEnvelope();

            envelope.FormatVersion = ReadElementValue(root, "formatVersion", KnownMetadata.EnvelopeVersion);
            envelope.DiagramId = ReadElementValue(root, "diagramId", string.Empty);
            envelope.DiagramName = ReadElementValue(root, "diagramName", string.Empty);
            envelope.EditorTarget = ReadElementValue(root, "editorTarget", string.Empty);
            envelope.SidecarPath = ReadElementValue(root, "sidecarPath", string.Empty);

            string editorModeText = ReadElementValue(root, "editorMode", EditorMode.Desktop.ToString());
            EditorMode editorMode;
            if (!Enum.TryParse(editorModeText, true, out editorMode))
            {
                editorMode = EditorMode.Desktop;
            }

            envelope.EditorMode = editorMode;

            string updatedUtcText = ReadElementValue(root, "updatedUtc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            DateTime updatedUtc;
            if (!DateTime.TryParse(updatedUtcText, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out updatedUtc))
            {
                updatedUtc = DateTime.UtcNow;
            }

            envelope.UpdatedUtc = updatedUtc;

            XElement xmlElement = root.Element("drawioXml");
            if (xmlElement != null)
            {
                envelope.DrawioXml = CompressionUtility.DecompressFromBase64(xmlElement.Value);
            }
            else
            {
                envelope.DrawioXml = string.Empty;
            }

            return envelope;
        }

        public string Serialize(DiagramEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            XElement root = new XElement(
                KnownMetadata.EnvelopeRootName,
                new XElement("formatVersion", envelope.FormatVersion ?? KnownMetadata.EnvelopeVersion),
                new XElement("diagramId", envelope.DiagramId ?? string.Empty),
                new XElement("diagramName", envelope.DiagramName ?? string.Empty),
                new XElement("editorMode", envelope.EditorMode.ToString()),
                new XElement("editorTarget", envelope.EditorTarget ?? string.Empty),
                new XElement("sidecarPath", envelope.SidecarPath ?? string.Empty),
                new XElement("updatedUtc", envelope.UpdatedUtc.ToString("o", CultureInfo.InvariantCulture)),
                new XElement(
                    "drawioXml",
                    new XAttribute("encoding", "gzip-base64"),
                    CompressionUtility.CompressToBase64(envelope.DrawioXml ?? string.Empty))
            );

            return root.ToString(SaveOptions.DisableFormatting);
        }

        private static string ReadElementValue(XElement root, string elementName, string defaultValue)
        {
            XElement element = root.Element(elementName);
            if (element == null)
            {
                return defaultValue;
            }

            return element.Value ?? defaultValue;
        }
    }
}

