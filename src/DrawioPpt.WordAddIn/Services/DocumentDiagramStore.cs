using System;
using System.Collections.Generic;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Services
{
    public class DocumentDiagramStore
    {
        private readonly IDiagramEnvelopeSerializer _serializer;

        public DocumentDiagramStore(IDiagramEnvelopeSerializer serializer)
        {
            if (serializer == null)
            {
                throw new ArgumentNullException("serializer");
            }

            _serializer = serializer;
        }

        public string Upsert(WordInterop.Document document, DiagramEnvelope envelope)
        {
            if (document == null || envelope == null)
            {
                return string.Empty;
            }

            string xml = _serializer.Serialize(envelope);
            CustomXMLPart addedPart = null;

            try
            {
                addedPart = document.CustomXMLParts.Add(xml, Type.Missing);
                if (addedPart == null)
                {
                    return string.Empty;
                }

                string addedPartId = addedPart.Id ?? string.Empty;
                DeleteDuplicates(document, envelope.DiagramId, addedPartId);
                return addedPartId;
            }
            catch
            {
                return string.Empty;
            }
        }

        public bool TryRead(WordInterop.Document document, string diagramId, out DiagramEnvelope envelope)
        {
            envelope = null;

            if (document == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return false;
            }

            try
            {
                int index;
                for (index = 1; index <= document.CustomXMLParts.Count; index++)
                {
                    CustomXMLPart candidate = document.CustomXMLParts[index];
                    DiagramEnvelope candidateEnvelope;
                    if (!TryDeserializePart(candidate, out candidateEnvelope))
                    {
                        continue;
                    }

                    if (string.Equals(candidateEnvelope.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase))
                    {
                        envelope = candidateEnvelope;
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public void Delete(WordInterop.Document document, string diagramId)
        {
            if (document == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return;
            }

            DeleteDuplicates(document, diagramId, string.Empty);
        }

        public int DeleteOrphans(WordInterop.Document document, ISet<string> liveDiagramIds)
        {
            if (document == null)
            {
                return 0;
            }

            int deletedCount = 0;

            try
            {
                int index;
                for (index = document.CustomXMLParts.Count; index >= 1; index--)
                {
                    CustomXMLPart part = document.CustomXMLParts[index];
                    DiagramEnvelope envelope;
                    if (!TryDeserializePart(part, out envelope))
                    {
                        continue;
                    }

                    bool hasLiveDiagramReference =
                        liveDiagramIds != null &&
                        !string.IsNullOrWhiteSpace(envelope.DiagramId) &&
                        liveDiagramIds.Contains(envelope.DiagramId);

                    if (hasLiveDiagramReference)
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
            }
            catch
            {
            }

            return deletedCount;
        }

        private void DeleteDuplicates(WordInterop.Document document, string diagramId, string keepPartId)
        {
            if (document == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return;
            }

            try
            {
                int index;
                for (index = document.CustomXMLParts.Count; index >= 1; index--)
                {
                    CustomXMLPart part = document.CustomXMLParts[index];
                    string currentPartId = part.Id ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(keepPartId) &&
                        string.Equals(currentPartId, keepPartId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    DiagramEnvelope envelope;
                    if (!TryDeserializePart(part, out envelope) ||
                        !string.Equals(envelope.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase))
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
            catch
            {
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
    }
}
