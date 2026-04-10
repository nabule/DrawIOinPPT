using System;
using DrawioPpt.Core;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.PowerPointAddIn.PowerPoint;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class ShapeMetadataService
    {
        private readonly IDiagramEnvelopeSerializer _serializer;
        private readonly PresentationSidecarPathBuilder _pathBuilder;

        public ShapeMetadataService(IDiagramEnvelopeSerializer serializer, PresentationSidecarPathBuilder pathBuilder)
        {
            _serializer = serializer;
            _pathBuilder = pathBuilder;
        }

        public SelectionContext BuildSelectionContext(PptInterop.Shape shape)
        {
            SelectionContext context = new SelectionContext();
            if (shape == null)
            {
                return context;
            }

            context.HasSelection = true;
            context.HasSingleShape = true;
            context.ShapeId = shape.Id;
            context.ShapeName = shape.Name;
            context.AlternativeText = shape.AlternativeText ?? string.Empty;
            context.DiagramId = ReadDiagramId(shape);
            context.IsManagedShape =
                !string.IsNullOrEmpty(context.DiagramId) ||
                !string.IsNullOrEmpty(ReadDiagramPartId(shape)) ||
                _serializer.CanDeserialize(context.AlternativeText);
            return context;
        }

        public void Clear(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            TryDeleteTag(shape, KnownMetadata.DiagramIdTagName);
            TryDeleteTag(shape, KnownMetadata.DiagramPartIdTagName);
            if (_serializer.CanDeserialize(shape.AlternativeText))
            {
                shape.AlternativeText = string.Empty;
            }
        }

        public DiagramEnvelope CreateOrUpdateEnvelope(PptInterop.Shape shape, PluginSettings settings, string presentationPath)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (settings == null)
            {
                settings = new PluginSettings();
            }

            DiagramEnvelope envelope;
            if (!TryRead(shape, out envelope))
            {
                envelope = new DiagramEnvelope();
            }

            if (string.IsNullOrWhiteSpace(envelope.DiagramId))
            {
                envelope.DiagramId = Guid.NewGuid().ToString("N");
            }

            envelope.DiagramName = shape.Name;
            envelope.EditorMode = settings.EditorMode;
            envelope.EditorTarget = settings.EditorMode == EditorMode.Desktop ? settings.DesktopEditorPath : settings.EditorUrl;
            envelope.UpdatedUtc = DateTime.UtcNow;

            if (string.IsNullOrWhiteSpace(envelope.DrawioXml))
            {
                envelope.DrawioXml = BuildPlaceholderXml(envelope.DiagramId, envelope.DiagramName);
            }

            if (settings.KeepSidecarFile && !string.IsNullOrWhiteSpace(presentationPath))
            {
                envelope.SidecarPath = _pathBuilder.BuildPath(
                    presentationPath,
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

        public DiagramEnvelope CreateNewEnvelope(PluginSettings settings, string presentationPath, string diagramName)
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

            if (settings.KeepSidecarFile && !string.IsNullOrWhiteSpace(presentationPath))
            {
                envelope.SidecarPath = _pathBuilder.BuildPath(
                    presentationPath,
                    envelope.DiagramId,
                    envelope.DiagramName,
                    settings.SidecarFolderName);
            }

            return envelope;
        }

        public void Save(PptInterop.Shape shape, DiagramEnvelope envelope)
        {
            Save(shape, envelope, string.Empty);
        }

        public void Save(PptInterop.Shape shape, DiagramEnvelope envelope, string customXmlPartId)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            TryDeleteTag(shape, KnownMetadata.DiagramIdTagName);
            TryDeleteTag(shape, KnownMetadata.DiagramPartIdTagName);
            shape.Tags.Add(KnownMetadata.DiagramIdTagName, envelope.DiagramId ?? string.Empty);
            if (!string.IsNullOrWhiteSpace(customXmlPartId))
            {
                shape.Tags.Add(KnownMetadata.DiagramPartIdTagName, customXmlPartId);
            }
            shape.AlternativeText = _serializer.Serialize(envelope);
        }

        public bool TryRead(PptInterop.Shape shape, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (shape == null)
            {
                return false;
            }

            string alternativeText = shape.AlternativeText;
            if (!_serializer.CanDeserialize(alternativeText))
            {
                return false;
            }

            envelope = _serializer.Deserialize(alternativeText);
            if (string.IsNullOrWhiteSpace(envelope.DiagramId))
            {
                envelope.DiagramId = ReadDiagramId(shape);
            }

            return true;
        }

        private static string BuildPlaceholderXml(string diagramId, string diagramName)
        {
            return
                "<mxfile host=\"DrawioPpt\">" +
                "<diagram id=\"" + diagramId + "\" name=\"" + diagramName + "\">" +
                "<mxGraphModel />" +
                "</diagram>" +
                "</mxfile>";
        }

        private static string ReadDiagramId(PptInterop.Shape shape)
        {
            if (shape == null || shape.Tags == null)
            {
                return string.Empty;
            }

            string value = shape.Tags[KnownMetadata.DiagramIdTagName];
            return value ?? string.Empty;
        }

        public string ReadDiagramPartId(PptInterop.Shape shape)
        {
            if (shape == null || shape.Tags == null)
            {
                return string.Empty;
            }

            string value = shape.Tags[KnownMetadata.DiagramPartIdTagName];
            return value ?? string.Empty;
        }

        private static void TryDeleteTag(PptInterop.Shape shape, string tagName)
        {
            if (shape == null || shape.Tags == null)
            {
                return;
            }

            string existingValue = shape.Tags[tagName];
            if (!string.IsNullOrEmpty(existingValue))
            {
                shape.Tags.Delete(tagName);
            }
        }
    }
}
