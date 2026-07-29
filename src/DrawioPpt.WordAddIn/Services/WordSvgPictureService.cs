using System;
using System.Drawing;
using System.Runtime.InteropServices;
using DrawioPpt.PowerPointAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Services
{
    public class WordSvgPictureService
    {
        private readonly SelectedPictureAccessor _pictureAccessor;
        private readonly SvgPowerPointSupportService _svgSupportService;

        public WordSvgPictureService(SelectedPictureAccessor pictureAccessor)
            : this(pictureAccessor, new SvgPowerPointSupportService())
        {
        }

        public WordSvgPictureService(SelectedPictureAccessor pictureAccessor, SvgPowerPointSupportService svgSupportService)
        {
            _pictureAccessor = pictureAccessor;
            _svgSupportService = svgSupportService;
        }

        public WordPictureReference InsertAtSelection(WordInterop.Application application, string svgFilePath, string pictureName)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            WordInterop.Document document = _pictureAccessor.GetActiveDocument(application);
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document is available.");
            }

            WordInterop.Selection selection = _pictureAccessor.GetSelection(application);
            if (selection == null)
            {
                throw new InvalidOperationException("No active Word selection is available.");
            }

            float sourceAspectRatio = TryReadSourceAspectRatio(svgFilePath);
            PictureSize size = ResolveDefaultSize(document, sourceAspectRatio);

            object linkToFile = false;
            object saveWithDocument = true;
            WordInterop.Range insertionRange = null;
            WordInterop.InlineShapes inlineShapes = null;
            WordInterop.InlineShape inlineShape = null;
            WordInterop.Shape shape = null;
            bool keepShape = false;
            try
            {
                insertionRange = selection.Range;
                object targetRange = insertionRange;
                inlineShapes = document.InlineShapes;
                inlineShape = inlineShapes.AddPicture(
                    svgFilePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    ref targetRange);
                inlineShape.LockAspectRatio = MsoTriState.msoFalse;
                inlineShape.Width = size.Width;
                inlineShape.Height = size.Height;
                if (sourceAspectRatio > 0f)
                {
                    inlineShape.LockAspectRatio = MsoTriState.msoTrue;
                }

                shape = inlineShape.ConvertToShape();
                SafeReleaseComObject(inlineShape);
                inlineShape = null;
                shape.LockAspectRatio = sourceAspectRatio > 0f
                    ? MsoTriState.msoTrue
                    : MsoTriState.msoFalse;
                TrySetTitle(shape, pictureName);
                TrySetWrapType(shape, WordInterop.WdWrapType.wdWrapFront);
                shape.Select(Type.Missing);
                keepShape = true;
                return WordPictureReference.FromShape(shape);
            }
            catch
            {
                TryDeleteShape(shape);
                TryDeleteInlineShape(inlineShape);
                throw;
            }
            finally
            {
                SafeReleaseComObject(insertionRange);
                SafeReleaseComObject(inlineShape);
                SafeReleaseComObject(inlineShapes);
                if (!keepShape)
                {
                    SafeReleaseComObject(shape);
                }
            }
        }

        public WordPictureReference Replace(WordInterop.Document document, WordPictureReference existingPicture, string svgFilePath)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            if (existingPicture == null)
            {
                throw new ArgumentNullException("existingPicture");
            }

            PictureSnapshot snapshot = PictureSnapshot.Capture(existingPicture);
            float sourceAspectRatio = TryReadSourceAspectRatio(svgFilePath);
            PictureSize maximumSize = ResolveMaximumSize(document);
            snapshot.ApplySourceAspectRatio(
                sourceAspectRatio,
                !existingPicture.IsInline,
                maximumSize.Width,
                maximumSize.Height);
            WordPictureReference replacement;

            if (existingPicture.IsInline)
            {
                replacement = ReplaceInline(document, existingPicture, svgFilePath, snapshot);
            }
            else
            {
                replacement = ReplaceFloating(document, existingPicture, svgFilePath, snapshot);
            }

            replacement.Select();
            return replacement;
        }

        private WordPictureReference ReplaceInline(WordInterop.Document document, WordPictureReference existingPicture, string svgFilePath, PictureSnapshot snapshot)
        {
            WordInterop.Range range = existingPicture.GetRange();
            if (range == null)
            {
                throw new InvalidOperationException("Unable to resolve the selected picture range.");
            }

            object linkToFile = false;
            object saveWithDocument = true;
            WordInterop.InlineShapes inlineShapes = null;
            WordInterop.InlineShape inlineShape = null;
            WordInterop.Shape shape = null;
            bool keepShape = false;
            try
            {
                range.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
                object targetRange = range;
                inlineShapes = document.InlineShapes;
                inlineShape = inlineShapes.AddPicture(
                    svgFilePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    ref targetRange);
                inlineShape.LockAspectRatio = MsoTriState.msoFalse;
                if (snapshot.Width > 0f)
                {
                    inlineShape.Width = snapshot.Width;
                }

                if (snapshot.Height > 0f)
                {
                    inlineShape.Height = snapshot.Height;
                }

                shape = inlineShape.ConvertToShape();
                SafeReleaseComObject(inlineShape);
                inlineShape = null;
                shape.LockAspectRatio = snapshot.LockAspectRatio;
                TrySetName(shape, snapshot.Name);
                TrySetTitle(shape, snapshot.Title);
                TrySetWrapType(shape, WordInterop.WdWrapType.wdWrapFront);

                DeleteExistingPicture(existingPicture);
                keepShape = true;
                return WordPictureReference.FromShape(shape);
            }
            catch
            {
                TryDeleteShape(shape);
                TryDeleteInlineShape(inlineShape);
                throw;
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(inlineShape);
                SafeReleaseComObject(inlineShapes);
                if (!keepShape)
                {
                    SafeReleaseComObject(shape);
                }
            }
        }

        private WordPictureReference ReplaceFloating(WordInterop.Document document, WordPictureReference existingPicture, string svgFilePath, PictureSnapshot snapshot)
        {
            WordInterop.Range anchor = existingPicture.GetRange();
            if (anchor == null)
            {
                throw new InvalidOperationException("Unable to resolve the selected picture anchor.");
            }

            object linkToFile = false;
            object saveWithDocument = true;
            object left = snapshot.Left;
            object top = snapshot.Top;
            object width = snapshot.Width;
            object height = snapshot.Height;
            object targetAnchor = anchor;
            WordInterop.Shapes shapes = null;
            WordInterop.Shape shape = null;
            bool keepShape = false;
            try
            {
                shapes = document.Shapes;
                shape = shapes.AddPicture(
                    svgFilePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    ref left,
                    ref top,
                    ref width,
                    ref height,
                    ref targetAnchor);
                shape.LockAspectRatio = snapshot.LockAspectRatio;
                TrySetName(shape, snapshot.Name);
                TrySetTitle(shape, snapshot.Title);
                TrySetWrapType(shape, WordInterop.WdWrapType.wdWrapFront);
                TrySetRelativeHorizontalPosition(shape, snapshot.RelativeHorizontalPosition);
                TrySetRelativeVerticalPosition(shape, snapshot.RelativeVerticalPosition);
                TrySetLeft(shape, snapshot.Left);
                TrySetTop(shape, snapshot.Top);

                DeleteExistingPicture(existingPicture);
                keepShape = true;
                return WordPictureReference.FromShape(shape);
            }
            catch
            {
                TryDeleteShape(shape);
                throw;
            }
            finally
            {
                SafeReleaseComObject(anchor);
                SafeReleaseComObject(shapes);
                if (!keepShape)
                {
                    SafeReleaseComObject(shape);
                }
            }
        }

        private float TryReadSourceAspectRatio(string svgFilePath)
        {
            if (_svgSupportService == null)
            {
                return TryReadRasterAspectRatio(svgFilePath);
            }

            float sourceWidth;
            float sourceHeight;
            if (_svgSupportService.TryReadSize(svgFilePath, out sourceWidth, out sourceHeight) &&
                sourceWidth > 0f &&
                sourceHeight > 0f)
            {
                return sourceWidth / sourceHeight;
            }

            return TryReadRasterAspectRatio(svgFilePath);
        }

        private static float TryReadRasterAspectRatio(string imageFilePath)
        {
            if (string.IsNullOrWhiteSpace(imageFilePath))
            {
                return 0f;
            }

            try
            {
                using (Image image = Image.FromFile(imageFilePath))
                {
                    if (image.Width > 0 && image.Height > 0)
                    {
                        return (float)image.Width / image.Height;
                    }
                }
            }
            catch
            {
            }

            return 0f;
        }

        private static PictureSize ResolveDefaultSize(
            WordInterop.Document document,
            float sourceAspectRatio)
        {
            PictureSize maximumSize = ResolveMaximumSize(document);
            float width = maximumSize.Width * 0.72f;
            float maximumHeight = maximumSize.Height * (0.55f / 0.90f);
            float height = sourceAspectRatio > 0f
                ? width / sourceAspectRatio
                : width * 0.62f;
            if (height > maximumHeight)
            {
                height = maximumHeight;
                width = sourceAspectRatio > 0f
                    ? height * sourceAspectRatio
                    : height / 0.62f;
            }

            return new PictureSize(width, height);
        }

        private static PictureSize ResolveMaximumSize(
            WordInterop.Document document)
        {
            float maximumWidth = 500f;
            float maximumHeight = 630f;
            WordInterop.PageSetup pageSetup = null;
            try
            {
                pageSetup = document.PageSetup;
                float usableWidth =
                    pageSetup.PageWidth -
                    pageSetup.LeftMargin -
                    pageSetup.RightMargin;
                float usableHeight =
                    pageSetup.PageHeight -
                    pageSetup.TopMargin -
                    pageSetup.BottomMargin;
                if (usableWidth > 0f)
                {
                    maximumWidth = usableWidth;
                }

                if (usableHeight > 0f)
                {
                    maximumHeight = usableHeight * 0.90f;
                }
            }
            catch
            {
            }
            finally
            {
                SafeReleaseComObject(pageSetup);
            }

            return new PictureSize(maximumWidth, maximumHeight);
        }

        private static void TrySetTitle(WordInterop.InlineShape shape, string title)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Title = title ?? string.Empty;
            }
            catch
            {
            }
        }

        private static void TrySetTitle(WordInterop.Shape shape, string title)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Title = title ?? string.Empty;
            }
            catch
            {
            }
        }

        private static void TrySetName(WordInterop.Shape shape, string name)
        {
            if (shape == null || string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            try
            {
                shape.Name = name;
            }
            catch
            {
            }
        }

        private static void TrySetWrapType(WordInterop.Shape shape, WordInterop.WdWrapType wrapType)
        {
            if (shape == null)
            {
                return;
            }

            WordInterop.WrapFormat wrapFormat = null;
            try
            {
                wrapFormat = shape.WrapFormat;
                wrapFormat.Type = wrapType;
            }
            catch
            {
            }
            finally
            {
                SafeReleaseComObject(wrapFormat);
            }
        }

        private static void TrySetRelativeHorizontalPosition(WordInterop.Shape shape, WordInterop.WdRelativeHorizontalPosition position)
        {
            try
            {
                shape.RelativeHorizontalPosition = position;
            }
            catch
            {
            }
        }

        private static void TrySetRelativeVerticalPosition(WordInterop.Shape shape, WordInterop.WdRelativeVerticalPosition position)
        {
            try
            {
                shape.RelativeVerticalPosition = position;
            }
            catch
            {
            }
        }

        private static void TrySetLeft(WordInterop.Shape shape, float left)
        {
            try
            {
                shape.Left = left;
            }
            catch
            {
            }
        }

        private static void TrySetTop(WordInterop.Shape shape, float top)
        {
            try
            {
                shape.Top = top;
            }
            catch
            {
            }
        }

        private static void DeleteExistingPicture(
            WordPictureReference picture)
        {
            if (picture.IsInline && picture.InlineShape != null)
            {
                picture.InlineShape.Delete();
                return;
            }

            if (picture.Shape != null)
            {
                picture.Shape.Delete();
                return;
            }

            throw new InvalidOperationException(
                "The existing Word picture is unavailable.");
        }

        private static void TryDeleteInlineShape(
            WordInterop.InlineShape inlineShape)
        {
            if (inlineShape == null)
            {
                return;
            }

            try
            {
                inlineShape.Delete();
            }
            catch
            {
            }
        }

        private static void TryDeleteShape(WordInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Delete();
            }
            catch
            {
            }
        }

        private static void SafeReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch
            {
            }
        }

        private struct PictureSize
        {
            public PictureSize(float width, float height)
            {
                Width = width;
                Height = height;
            }

            public float Width;
            public float Height;
        }

        private sealed class PictureSnapshot
        {
            public string Name { get; private set; }
            public string Title { get; private set; }
            public float Width { get; private set; }
            public float Height { get; private set; }
            public float Left { get; private set; }
            public float Top { get; private set; }
            public MsoTriState LockAspectRatio { get; private set; }
            public WordInterop.WdWrapType WrapType { get; private set; }
            public WordInterop.WdRelativeHorizontalPosition RelativeHorizontalPosition { get; private set; }
            public WordInterop.WdRelativeVerticalPosition RelativeVerticalPosition { get; private set; }

            public static PictureSnapshot Capture(WordPictureReference picture)
            {
                PictureSnapshot snapshot = new PictureSnapshot();
                snapshot.Name = picture.Name;
                snapshot.Title = picture.Title;
                snapshot.Width = picture.Width;
                snapshot.Height = picture.Height;
                snapshot.LockAspectRatio = picture.LockAspectRatio;
                snapshot.Left = 0f;
                snapshot.Top = 0f;
                snapshot.WrapType = WordInterop.WdWrapType.wdWrapInline;
                snapshot.RelativeHorizontalPosition = WordInterop.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn;
                snapshot.RelativeVerticalPosition = WordInterop.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;

                if (!picture.IsInline && picture.Shape != null)
                {
                    WordInterop.WrapFormat wrapFormat = null;
                    try
                    {
                        snapshot.Left = picture.Shape.Left;
                        snapshot.Top = picture.Shape.Top;
                        wrapFormat = picture.Shape.WrapFormat;
                        snapshot.WrapType = wrapFormat.Type;
                        snapshot.RelativeHorizontalPosition = picture.Shape.RelativeHorizontalPosition;
                        snapshot.RelativeVerticalPosition = picture.Shape.RelativeVerticalPosition;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        SafeReleaseComObject(wrapFormat);
                    }
                }

                if (snapshot.Width <= 0f)
                {
                    snapshot.Width = 360f;
                }

                if (snapshot.Height <= 0f)
                {
                    snapshot.Height = 220f;
                }

                return snapshot;
            }

            public void ApplySourceAspectRatio(
                float sourceAspectRatio,
                bool preserveFloatingCenter,
                float maximumWidth,
                float maximumHeight)
            {
                if (sourceAspectRatio <= 0f || Width <= 0f)
                {
                    return;
                }

                float originalWidth = Width;
                float originalHeight = Height;
                float correctedWidth = Width;
                float correctedHeight = correctedWidth / sourceAspectRatio;
                if (maximumWidth > 0f && correctedWidth > maximumWidth)
                {
                    correctedWidth = maximumWidth;
                    correctedHeight = correctedWidth / sourceAspectRatio;
                }

                if (maximumHeight > 0f &&
                    correctedHeight > maximumHeight)
                {
                    correctedHeight = maximumHeight;
                    correctedWidth = correctedHeight * sourceAspectRatio;
                }

                if (correctedWidth <= 0f || correctedHeight <= 0f)
                {
                    return;
                }

                Width = correctedWidth;
                Height = correctedHeight;
                LockAspectRatio = MsoTriState.msoTrue;

                if (preserveFloatingCenter)
                {
                    if (originalWidth > 0f)
                    {
                        Left += (originalWidth - correctedWidth) / 2f;
                    }

                    if (originalHeight > 0f)
                    {
                        Top += (originalHeight - correctedHeight) / 2f;
                    }
                }
            }
        }
    }
}
