using System;
using System.Collections.Generic;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class SelectedPictureAccessor
    {
        private readonly IDiagramEnvelopeSerializer _serializer;

        public SelectedPictureAccessor()
            : this(new DiagramEnvelopeSerializer())
        {
        }

        public SelectedPictureAccessor(IDiagramEnvelopeSerializer serializer)
        {
            _serializer = serializer;
        }

        public WordInterop.Document GetActiveDocument(WordInterop.Application application)
        {
            if (application == null)
            {
                return null;
            }

            try
            {
                return application.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        public string GetDocumentFullName(WordInterop.Application application)
        {
            WordInterop.Document document = GetActiveDocument(application);
            if (document == null)
            {
                return string.Empty;
            }

            try
            {
                return document.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public WordInterop.Selection GetSelection(WordInterop.Application application)
        {
            if (application == null)
            {
                return null;
            }

            try
            {
                return application.Selection;
            }
            catch
            {
                return null;
            }
        }

        public WordPictureReference GetSingleSelectedPicture(WordInterop.Application application)
        {
            return GetSingleSelectedPicture(GetSelection(application));
        }

        public WordPictureReference GetSingleSelectedPicture(WordInterop.Selection selection)
        {
            if (selection == null)
            {
                return null;
            }

            try
            {
                if (selection.InlineShapes != null && selection.InlineShapes.Count == 1)
                {
                    return WordPictureReference.FromInlineShape(selection.InlineShapes[1]);
                }
            }
            catch
            {
            }

            try
            {
                if (selection.ShapeRange != null && selection.ShapeRange.Count == 1)
                {
                    return WordPictureReference.FromShape(selection.ShapeRange[1]);
                }
            }
            catch
            {
            }

            return null;
        }

        public WordPictureReference FindPictureByDiagramId(WordInterop.Application application, string diagramId)
        {
            return FindPictureByDiagramId(GetActiveDocument(application), diagramId);
        }

        public WordPictureReference FindPictureByDiagramId(WordInterop.Document document, string diagramId)
        {
            if (document == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return null;
            }

            try
            {
                int index;
                for (index = 1; index <= document.InlineShapes.Count; index++)
                {
                    WordInterop.InlineShape inlineShape = document.InlineShapes[index];
                    if (HasDiagramId(inlineShape.AlternativeText, diagramId))
                    {
                        return WordPictureReference.FromInlineShape(inlineShape);
                    }
                }
            }
            catch
            {
            }

            try
            {
                int index;
                for (index = 1; index <= document.Shapes.Count; index++)
                {
                    WordInterop.Shape shape = document.Shapes[index];
                    if (HasDiagramId(shape.AlternativeText, diagramId))
                    {
                        return WordPictureReference.FromShape(shape);
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        public void CollectManagedDiagramReferences(WordInterop.Document document, ISet<string> diagramIds)
        {
            if (document == null || diagramIds == null)
            {
                return;
            }

            try
            {
                int index;
                for (index = 1; index <= document.InlineShapes.Count; index++)
                {
                    AddDiagramReference(document.InlineShapes[index].AlternativeText, diagramIds);
                }
            }
            catch
            {
            }

            try
            {
                int index;
                for (index = 1; index <= document.Shapes.Count; index++)
                {
                    AddDiagramReference(document.Shapes[index].AlternativeText, diagramIds);
                }
            }
            catch
            {
            }
        }

        public void SelectPicture(WordPictureReference picture)
        {
            if (picture != null)
            {
                picture.Select();
            }
        }

        private bool HasDiagramId(string metadata, string diagramId)
        {
            DiagramEnvelope envelope;
            return TryReadEnvelope(metadata, out envelope) &&
                string.Equals(envelope.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase);
        }

        private void AddDiagramReference(string metadata, ISet<string> diagramIds)
        {
            DiagramEnvelope envelope;
            if (TryReadEnvelope(metadata, out envelope) && !string.IsNullOrWhiteSpace(envelope.DiagramId))
            {
                diagramIds.Add(envelope.DiagramId);
            }
        }

        private bool TryReadEnvelope(string metadata, out DiagramEnvelope envelope)
        {
            envelope = null;
            if (!_serializer.CanDeserialize(metadata))
            {
                return false;
            }

            envelope = _serializer.Deserialize(metadata);
            return envelope != null;
        }
    }
}
