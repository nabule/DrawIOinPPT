using DrawioPpt.Core;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Services;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class PowerPointShapeSelectionReader
    {
        private readonly IDiagramEnvelopeSerializer _serializer;

        public PowerPointShapeSelectionReader()
            : this(new DiagramEnvelopeSerializer())
        {
        }

        public PowerPointShapeSelectionReader(IDiagramEnvelopeSerializer serializer)
        {
            _serializer = serializer;
        }

        public SelectionContext Read(PptInterop.Selection selection)
        {
            SelectionContext context = new SelectionContext();

            if (selection == null)
            {
                return context;
            }

            if (selection.Type != PptInterop.PpSelectionType.ppSelectionShapes)
            {
                return context;
            }

            context.HasSelection = true;

            if (selection.ShapeRange == null || selection.ShapeRange.Count != 1)
            {
                return context;
            }

            PptInterop.Shape shape = selection.ShapeRange[1];
            context.HasSingleShape = true;
            context.ShapeId = shape.Id;
            context.ShapeName = shape.Name;
            context.AlternativeText = shape.AlternativeText ?? string.Empty;
            context.DiagramId = ReadDiagramId(shape);
            context.IsManagedShape = !string.IsNullOrEmpty(context.DiagramId) || _serializer.CanDeserialize(context.AlternativeText);

            return context;
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
    }
}
