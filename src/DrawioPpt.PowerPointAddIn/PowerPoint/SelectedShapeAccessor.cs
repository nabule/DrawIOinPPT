using System;
using DrawioPpt.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectedShapeAccessor
    {
        public PptInterop.Shape GetSingleSelectedShape(PptInterop.Application application)
        {
            if (application == null || application.ActiveWindow == null)
            {
                return null;
            }

            PptInterop.Selection selection = application.ActiveWindow.Selection;
            if (selection == null)
            {
                return null;
            }

            if (selection.Type != PptInterop.PpSelectionType.ppSelectionShapes)
            {
                return null;
            }

            if (selection.ShapeRange == null || selection.ShapeRange.Count != 1)
            {
                return null;
            }

            return selection.ShapeRange[1];
        }

        public string GetPresentationFullName(PptInterop.Application application)
        {
            if (application == null || application.ActivePresentation == null)
            {
                return string.Empty;
            }

            return application.ActivePresentation.FullName ?? string.Empty;
        }

        public PptInterop.Presentation GetPresentation(PptInterop.Shape shape)
        {
            PptInterop.Slide slide = GetParentSlide(shape);
            if (slide == null)
            {
                return null;
            }

            return slide.Parent as PptInterop.Presentation;
        }

        public PptInterop.Slide GetActiveSlide(PptInterop.Application application, bool createIfMissing)
        {
            if (application == null || application.ActivePresentation == null)
            {
                return null;
            }

            if (application.ActiveWindow != null && application.ActiveWindow.View != null)
            {
                object activeSlide = application.ActiveWindow.View.Slide;
                PptInterop.Slide typedSlide = activeSlide as PptInterop.Slide;
                if (typedSlide != null)
                {
                    return typedSlide;
                }
            }

            PptInterop.Presentation presentation = application.ActivePresentation;
            if (presentation.Slides.Count > 0)
            {
                return presentation.Slides[1];
            }

            if (!createIfMissing)
            {
                return null;
            }

            return presentation.Slides.Add(1, PptInterop.PpSlideLayout.ppLayoutBlank);
        }

        public PptInterop.Slide GetParentSlide(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return null;
            }

            object parent = shape.Parent;
            PptInterop.Slide slide = parent as PptInterop.Slide;
            if (slide != null)
            {
                return slide;
            }

            PptInterop.Shapes shapes = parent as PptInterop.Shapes;
            if (shapes != null)
            {
                return shapes.Parent as PptInterop.Slide;
            }

            return null;
        }

        public PptInterop.Shape FindShapeByDiagramId(PptInterop.Application application, string diagramId)
        {
            if (application == null || application.ActivePresentation == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return null;
            }

            PptInterop.Presentation presentation = application.ActivePresentation;
            int slideIndex;
            for (slideIndex = 1; slideIndex <= presentation.Slides.Count; slideIndex++)
            {
                PptInterop.Slide slide = presentation.Slides[slideIndex];
                PptInterop.Shape shape = FindShapeByDiagramId(slide.Shapes, diagramId);
                if (shape != null)
                {
                    return shape;
                }
            }

            return null;
        }

        private static PptInterop.Shape FindShapeByDiagramId(PptInterop.Shapes shapes, string diagramId)
        {
            if (shapes == null)
            {
                return null;
            }

            int index;
            for (index = 1; index <= shapes.Count; index++)
            {
                PptInterop.Shape shape = shapes[index];
                if (string.Equals(shape.Tags[KnownMetadata.DiagramIdTagName], diagramId, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }

                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    PptInterop.Shape nested = FindShapeByDiagramId(shape.GroupItems, diagramId);
                    if (nested != null)
                    {
                        return nested;
                    }
                }
            }

            return null;
        }

        private static PptInterop.Shape FindShapeByDiagramId(PptInterop.GroupShapes groupShapes, string diagramId)
        {
            if (groupShapes == null)
            {
                return null;
            }

            int index;
            for (index = 1; index <= groupShapes.Count; index++)
            {
                PptInterop.Shape shape = groupShapes[index];
                if (string.Equals(shape.Tags[KnownMetadata.DiagramIdTagName], diagramId, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }

                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    PptInterop.Shape nested = FindShapeByDiagramId(shape.GroupItems, diagramId);
                    if (nested != null)
                    {
                        return nested;
                    }
                }
            }

            return null;
        }
    }
}
