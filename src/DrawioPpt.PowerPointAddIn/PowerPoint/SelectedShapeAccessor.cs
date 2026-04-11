using System;
using System.Collections.Generic;
using DrawioPpt.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectedShapeAccessor
    {
        public PptInterop.Presentation GetActivePresentation(PptInterop.Application application)
        {
            if (application == null)
            {
                return null;
            }

            try
            {
                return application.ActivePresentation;
            }
            catch
            {
                return null;
            }
        }

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
            PptInterop.Presentation presentation = GetActivePresentation(application);
            if (presentation == null)
            {
                return string.Empty;
            }

            return presentation.FullName ?? string.Empty;
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
            PptInterop.Presentation presentation = GetActivePresentation(application);
            if (presentation == null)
            {
                return null;
            }

            PptInterop.Slide activeSlide = TryGetSlideFromActiveView(application);
            if (activeSlide != null)
            {
                return activeSlide;
            }

            if (presentation.Slides.Count > 0)
            {
                PptInterop.Slide firstSlide = presentation.Slides[1];
                TryActivateSlide(firstSlide);
                return firstSlide;
            }

            if (!createIfMissing)
            {
                return null;
            }

            PptInterop.Slide createdSlide = presentation.Slides.Add(1, PptInterop.PpSlideLayout.ppLayoutBlank);
            TryActivateSlide(createdSlide);
            return createdSlide;
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
            PptInterop.Presentation presentation = GetActivePresentation(application);
            if (presentation == null || string.IsNullOrWhiteSpace(diagramId))
            {
                return null;
            }

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

        public void CollectManagedDiagramReferences(PptInterop.Presentation presentation, ISet<string> diagramIds, ISet<string> partIds)
        {
            if (presentation == null)
            {
                return;
            }

            int slideIndex;
            for (slideIndex = 1; slideIndex <= presentation.Slides.Count; slideIndex++)
            {
                PptInterop.Slide slide = presentation.Slides[slideIndex];
                CollectManagedDiagramReferences(slide.Shapes, diagramIds, partIds);
            }
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

        private static void CollectManagedDiagramReferences(PptInterop.Shapes shapes, ISet<string> diagramIds, ISet<string> partIds)
        {
            if (shapes == null)
            {
                return;
            }

            int index;
            for (index = 1; index <= shapes.Count; index++)
            {
                PptInterop.Shape shape = shapes[index];
                AddManagedDiagramReference(shape, diagramIds, partIds);

                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    CollectManagedDiagramReferences(shape.GroupItems, diagramIds, partIds);
                }
            }
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

        private static void CollectManagedDiagramReferences(PptInterop.GroupShapes groupShapes, ISet<string> diagramIds, ISet<string> partIds)
        {
            if (groupShapes == null)
            {
                return;
            }

            int index;
            for (index = 1; index <= groupShapes.Count; index++)
            {
                PptInterop.Shape shape = groupShapes[index];
                AddManagedDiagramReference(shape, diagramIds, partIds);

                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    CollectManagedDiagramReferences(shape.GroupItems, diagramIds, partIds);
                }
            }
        }

        private static void AddManagedDiagramReference(PptInterop.Shape shape, ISet<string> diagramIds, ISet<string> partIds)
        {
            if (shape == null || shape.Tags == null)
            {
                return;
            }

            if (diagramIds != null)
            {
                string diagramId = shape.Tags[KnownMetadata.DiagramIdTagName];
                if (!string.IsNullOrWhiteSpace(diagramId))
                {
                    diagramIds.Add(diagramId);
                }
            }

            if (partIds != null)
            {
                string partId = shape.Tags[KnownMetadata.DiagramPartIdTagName];
                if (!string.IsNullOrWhiteSpace(partId))
                {
                    partIds.Add(partId);
                }
            }
        }

        private static PptInterop.Slide TryGetSlideFromActiveView(PptInterop.Application application)
        {
            if (application == null || application.ActiveWindow == null || application.ActiveWindow.View == null)
            {
                return null;
            }

            try
            {
                object activeSlide = application.ActiveWindow.View.Slide;
                return activeSlide as PptInterop.Slide;
            }
            catch
            {
                return null;
            }
        }

        private static void TryActivateSlide(PptInterop.Slide slide)
        {
            if (slide == null)
            {
                return;
            }

            try
            {
                slide.Select();
            }
            catch
            {
            }
        }
    }
}
