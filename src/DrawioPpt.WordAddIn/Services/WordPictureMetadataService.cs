using System;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Word;

namespace DrawioPpt.WordAddIn.Services
{
    public class WordPictureMetadataService
    {
        private readonly IDiagramEnvelopeSerializer _serializer;
        private readonly PresentationSidecarPathBuilder _pathBuilder;

        public WordPictureMetadataService(IDiagramEnvelopeSerializer serializer, PresentationSidecarPathBuilder pathBuilder)
        {
            _serializer = serializer;
            _pathBuilder = pathBuilder;
        }

        public SelectionContext BuildSelectionContext(WordPictureReference picture)
        {
            SelectionContext context = new SelectionContext();
            if (picture == null)
            {
                return context;
            }

            context.HasSelection = true;
            context.HasSinglePicture = true;
            context.PictureName = picture.Name;
            context.ReferenceKey = picture.ReferenceKey;
            context.AlternativeText = picture.AlternativeText ?? string.Empty;

            DiagramEnvelope envelope;
            if (TryRead(picture, out envelope))
            {
                context.DiagramId = envelope.DiagramId ?? string.Empty;
                context.IsManagedPicture = !string.IsNullOrWhiteSpace(context.DiagramId);
            }

            return context;
        }

        public void Clear(WordPictureReference picture)
        {
            if (picture == null)
            {
                return;
            }

            if (_serializer.CanDeserialize(picture.AlternativeText))
            {
                picture.AlternativeText = string.Empty;
            }

            picture.Title = string.Empty;
        }

        public DiagramEnvelope CreateOrUpdateEnvelope(WordPictureReference picture, PluginSettings settings, string documentPath)
        {
            if (picture == null)
            {
                throw new ArgumentNullException("picture");
            }

            if (settings == null)
            {
                settings = new PluginSettings();
            }

            DiagramEnvelope envelope;
            if (!TryRead(picture, out envelope))
            {
                envelope = new DiagramEnvelope();
            }

            if (string.IsNullOrWhiteSpace(envelope.DiagramId))
            {
                envelope.DiagramId = Guid.NewGuid().ToString("N");
            }

            envelope.DiagramName = string.IsNullOrWhiteSpace(picture.Name) ? "Drawio Diagram" : picture.Name;
            envelope.EditorMode = settings.EditorMode;
            envelope.EditorTarget = settings.EditorMode == EditorMode.Desktop ? settings.DesktopEditorPath : settings.EditorUrl;
            envelope.UpdatedUtc = DateTime.UtcNow;

            if (string.IsNullOrWhiteSpace(envelope.DrawioXml))
            {
                envelope.DrawioXml = BuildPlaceholderXml(envelope.DiagramId, envelope.DiagramName);
            }

            if (settings.KeepSidecarFile && _pathBuilder.CanBuildPath(documentPath))
            {
                envelope.SidecarPath = _pathBuilder.BuildPath(
                    documentPath,
                    envelope.DiagramId,
                    envelope.DiagramName,
                    settings.SidecarFolderName);
            }
            else
            {
                envelope.SidecarPath = string.Empty;
            }

            return envelope;
        }

        public DiagramEnvelope CreateNewEnvelope(PluginSettings settings, string documentPath, string diagramName)
        {
            if (settings == null)
            {
                settings = new PluginSettings();
            }

            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = string.IsNullOrWhiteSpace(diagramName) ? "Drawio Diagram" : diagramName;
            envelope.EditorMode = settings.EditorMode;
            envelope.EditorTarget = settings.EditorMode == EditorMode.Desktop ? settings.DesktopEditorPath : settings.EditorUrl;
            envelope.UpdatedUtc = DateTime.UtcNow;
            envelope.DrawioXml = BuildPlaceholderXml(envelope.DiagramId, envelope.DiagramName);

            if (settings.KeepSidecarFile && _pathBuilder.CanBuildPath(documentPath))
            {
                envelope.SidecarPath = _pathBuilder.BuildPath(
                    documentPath,
                    envelope.DiagramId,
                    envelope.DiagramName,
                    settings.SidecarFolderName);
            }

            return envelope;
        }

        public void Save(WordPictureReference picture, DiagramEnvelope envelope)
        {
            ValidateSaveArguments(picture, envelope);
            SaveEnvelope(picture, CreatePictureReferenceEnvelope(envelope));
        }

        private static void ValidateSaveArguments(WordPictureReference picture, DiagramEnvelope envelope)
        {
            if (picture == null)
            {
                throw new ArgumentNullException("picture");
            }

            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }
        }

        private static DiagramEnvelope CreatePictureReferenceEnvelope(DiagramEnvelope envelope)
        {
            DiagramEnvelope referenceEnvelope = new DiagramEnvelope();
            referenceEnvelope.FormatVersion = envelope.FormatVersion;
            referenceEnvelope.DiagramId = envelope.DiagramId;
            referenceEnvelope.DiagramName = envelope.DiagramName;
            referenceEnvelope.EditorMode = envelope.EditorMode;
            referenceEnvelope.EditorTarget = envelope.EditorTarget;
            referenceEnvelope.SidecarPath = envelope.SidecarPath;
            referenceEnvelope.UpdatedUtc = envelope.UpdatedUtc;
            referenceEnvelope.DrawioXml = string.Empty;
            return referenceEnvelope;
        }

        private void SaveEnvelope(WordPictureReference picture, DiagramEnvelope envelope)
        {
            picture.Title = envelope.DiagramName ?? string.Empty;
            picture.AlternativeText = _serializer.Serialize(envelope);
        }

        public bool TryRead(WordPictureReference picture, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (picture == null)
            {
                return false;
            }

            string alternativeText = picture.AlternativeText;
            if (!_serializer.CanDeserialize(alternativeText))
            {
                return false;
            }

            envelope = _serializer.Deserialize(alternativeText);
            return envelope != null && !string.IsNullOrWhiteSpace(envelope.DiagramId);
        }

        private static string BuildPlaceholderXml(string diagramId, string diagramName)
        {
            string safeDiagramId = EscapeXml(diagramId);
            string safeDiagramName = EscapeXml(diagramName);
            return
                "<mxfile host=\"DrawioWord\">" +
                "<diagram id=\"" + safeDiagramId + "\" name=\"" + safeDiagramName + "\">" +
                "<mxGraphModel dx=\"1000\" dy=\"1000\" grid=\"1\" gridSize=\"10\" guides=\"1\" tooltips=\"1\" connect=\"1\" arrows=\"1\" fold=\"1\" page=\"1\" pageScale=\"1\" pageWidth=\"850\" pageHeight=\"1100\" math=\"0\" shadow=\"0\">" +
                "<root>" +
                "<mxCell id=\"0\"/>" +
                "<mxCell id=\"1\" parent=\"0\"/>" +
                "<mxCell id=\"2\" value=\"Double-click to edit\" style=\"rounded=1;whiteSpace=wrap;html=1;\" vertex=\"1\" parent=\"1\">" +
                "<mxGeometry x=\"120\" y=\"120\" width=\"180\" height=\"60\" as=\"geometry\"/>" +
                "</mxCell>" +
                "</root>" +
                "</mxGraphModel>" +
                "</diagram>" +
                "</mxfile>";
        }

        private static string EscapeXml(string value)
        {
            return System.Security.SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }
}
