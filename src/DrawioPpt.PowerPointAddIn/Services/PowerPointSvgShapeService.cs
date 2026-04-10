using System;
using DrawioPpt.PowerPointAddIn.PowerPoint;
using Microsoft.Office.Core;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class PowerPointSvgShapeService
    {
        private readonly SelectedShapeAccessor _shapeAccessor;

        public PowerPointSvgShapeService(SelectedShapeAccessor shapeAccessor)
        {
            _shapeAccessor = shapeAccessor;
        }

        public PptInterop.Shape InsertOnActiveSlide(PptInterop.Application application, string svgFilePath, string shapeName)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            PptInterop.Slide slide = _shapeAccessor.GetActiveSlide(application, true);
            if (slide == null)
            {
                throw new InvalidOperationException("No active slide is available.");
            }

            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = application.ActivePresentation.PageSetup.SlideHeight;
            float width = slideWidth * 0.50f;
            float height = slideHeight * 0.38f;
            float left = (slideWidth - width) / 2f;
            float top = (slideHeight - height) / 2f;

            PptInterop.Shape shape = slide.Shapes.AddPicture(
                svgFilePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                left,
                top,
                width,
                height);

            TrySetShapeName(shape, shapeName);
            shape.Select(MsoTriState.msoFalse);
            return shape;
        }

        public PptInterop.Shape Replace(PptInterop.Shape existingShape, string svgFilePath)
        {
            if (existingShape == null)
            {
                throw new ArgumentNullException("existingShape");
            }

            PptInterop.Slide slide = _shapeAccessor.GetParentSlide(existingShape);
            if (slide == null)
            {
                throw new InvalidOperationException("Unable to resolve the parent slide.");
            }

            ShapeSnapshot snapshot = ShapeSnapshot.Capture(existingShape);
            existingShape.Delete();

            PptInterop.Shape newShape = slide.Shapes.AddPicture(
                svgFilePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                snapshot.Left,
                snapshot.Top,
                snapshot.Width,
                snapshot.Height);

            newShape.Rotation = snapshot.Rotation;
            TrySetShapeName(newShape, snapshot.Name);
            MoveToZOrder(newShape, snapshot.ZOrderPosition);
            newShape.Select(MsoTriState.msoFalse);
            return newShape;
        }

        private static void MoveToZOrder(PptInterop.Shape shape, int targetPosition)
        {
            while (shape.ZOrderPosition > targetPosition)
            {
                shape.ZOrder(MsoZOrderCmd.msoSendBackward);
            }
        }

        private static void TrySetShapeName(PptInterop.Shape shape, string shapeName)
        {
            if (shape == null || string.IsNullOrWhiteSpace(shapeName))
            {
                return;
            }

            try
            {
                shape.Name = shapeName;
            }
            catch
            {
            }
        }

        private sealed class ShapeSnapshot
        {
            public float Left { get; private set; }
            public float Top { get; private set; }
            public float Width { get; private set; }
            public float Height { get; private set; }
            public float Rotation { get; private set; }
            public int ZOrderPosition { get; private set; }
            public string Name { get; private set; }

            public static ShapeSnapshot Capture(PptInterop.Shape shape)
            {
                ShapeSnapshot snapshot = new ShapeSnapshot();
                snapshot.Left = shape.Left;
                snapshot.Top = shape.Top;
                snapshot.Width = shape.Width;
                snapshot.Height = shape.Height;
                snapshot.Rotation = shape.Rotation;
                snapshot.ZOrderPosition = shape.ZOrderPosition;
                snapshot.Name = shape.Name;
                return snapshot;
            }
        }
    }
}

