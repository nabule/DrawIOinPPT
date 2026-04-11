using System;
using System.Collections.Generic;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using Microsoft.Office.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class PresentationDiagramStore
    {
        private readonly IDiagramEnvelopeSerializer _serializer;

        public PresentationDiagramStore(IDiagramEnvelopeSerializer serializer)
        {
            if (serializer == null)
            {
                throw new ArgumentNullException("serializer");
            }

            _serializer = serializer;
        }

        public string Upsert(PptInterop.Presentation presentation, DiagramEnvelope envelope, string existingPartId)
        {
            if (presentation == null || envelope == null)
            {
                return string.Empty;
            }

            string xml = _serializer.Serialize(envelope);
            CustomXMLPart addedPart = null;

            try
            {
                addedPart = presentation.CustomXMLParts.Add(xml, Type.Missing);
                if (addedPart == null)
                {
                    return string.Empty;
                }

                string addedPartId = addedPart.Id ?? string.Empty;
                DeleteDuplicates(presentation, envelope.DiagramId, addedPartId, existingPartId);
                return addedPartId;
            }
            catch
            {
                return string.Empty;
            }
        }

        public bool TryRead(PptInterop.Presentation presentation, string diagramId, string partId, out DiagramEnvelope envelope, out string resolvedPartId)
        {
            envelope = null;
            resolvedPartId = string.Empty;

            if (presentation == null)
            {
                return false;
            }

            CustomXMLPart part;
            if (TryGetPartById(presentation, partId, out part) && TryDeserializePart(part, out envelope))
            {
                resolvedPartId = part.Id ?? string.Empty;
                return true;
            }

            if (string.IsNullOrWhiteSpace(diagramId))
            {
                return false;
            }

            int index;
            for (index = 1; index <= presentation.CustomXMLParts.Count; index++)
            {
                CustomXMLPart candidate = presentation.CustomXMLParts[index];
                DiagramEnvelope candidateEnvelope;
                if (!TryDeserializePart(candidate, out candidateEnvelope))
                {
                    continue;
                }

                if (string.Equals(candidateEnvelope.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase))
                {
                    envelope = candidateEnvelope;
                    resolvedPartId = candidate.Id ?? string.Empty;
                    return true;
                }
            }

            return false;
        }

        public void Delete(PptInterop.Presentation presentation, string diagramId, string partId)
        {
            if (presentation == null)
            {
                return;
            }

            DeleteDuplicates(presentation, diagramId, string.Empty, partId);
        }

        public int DeleteOrphans(PptInterop.Presentation presentation, ISet<string> liveDiagramIds, ISet<string> livePartIds)
        {
            if (presentation == null)
            {
                return 0;
            }

            int deletedCount = 0;
            int index;
            for (index = presentation.CustomXMLParts.Count; index >= 1; index--)
            {
                CustomXMLPart part = presentation.CustomXMLParts[index];
                DiagramEnvelope envelope;
                if (!TryDeserializePart(part, out envelope))
                {
                    continue;
                }

                string currentPartId = part.Id ?? string.Empty;
                bool hasLivePartReference =
                    livePartIds != null &&
                    !string.IsNullOrWhiteSpace(currentPartId) &&
                    livePartIds.Contains(currentPartId);
                bool hasLiveDiagramReference =
                    liveDiagramIds != null &&
                    !string.IsNullOrWhiteSpace(envelope.DiagramId) &&
                    liveDiagramIds.Contains(envelope.DiagramId);

                if (hasLivePartReference || hasLiveDiagramReference)
                {
                    continue;
                }

                try
                {
                    part.Delete();
                    deletedCount++;
                }
                catch
                {
                }
            }

            return deletedCount;
        }

        private void DeleteDuplicates(PptInterop.Presentation presentation, string diagramId, string keepPartId, string preferredDeletePartId)
        {
            if (presentation == null)
            {
                return;
            }

            int index;
            for (index = presentation.CustomXMLParts.Count; index >= 1; index--)
            {
                CustomXMLPart part = presentation.CustomXMLParts[index];
                string currentPartId = part.Id ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(keepPartId) &&
                    string.Equals(currentPartId, keepPartId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                bool shouldDelete = false;
                if (!string.IsNullOrWhiteSpace(preferredDeletePartId) &&
                    string.Equals(currentPartId, preferredDeletePartId, StringComparison.OrdinalIgnoreCase))
                {
                    shouldDelete = true;
                }
                else if (!string.IsNullOrWhiteSpace(diagramId))
                {
                    DiagramEnvelope envelope;
                    if (TryDeserializePart(part, out envelope) &&
                        string.Equals(envelope.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase))
                    {
                        shouldDelete = true;
                    }
                }

                if (!shouldDelete)
                {
                    continue;
                }

                try
                {
                    part.Delete();
                }
                catch
                {
                }
            }
        }

        private bool TryDeserializePart(CustomXMLPart part, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (part == null)
            {
                return false;
            }

            string xml;
            try
            {
                xml = part.XML;
            }
            catch
            {
                return false;
            }

            if (!_serializer.CanDeserialize(xml))
            {
                return false;
            }

            envelope = _serializer.Deserialize(xml);
            return envelope != null;
        }

        private static bool TryGetPartById(PptInterop.Presentation presentation, string partId, out CustomXMLPart part)
        {
            part = null;
            if (presentation == null || string.IsNullOrWhiteSpace(partId))
            {
                return false;
            }

            try
            {
                part = presentation.CustomXMLParts.SelectByID(partId);
                return part != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
