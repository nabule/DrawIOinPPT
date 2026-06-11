using System;
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

            float width = ResolveDefaultWidth(document);
            float height = width * 0.62f;
            string renderSvgPath = PrepareWordSvg(svgFilePath, width, height);

            object linkToFile = false;
            object saveWithDocument = true;
            object range = selection.Range;
            WordInterop.InlineShape inlineShape = document.InlineShapes.AddPicture(renderSvgPath, ref linkToFile, ref saveWithDocument, ref range);
            inlineShape.LockAspectRatio = MsoTriState.msoFalse;
            inlineShape.Width = width;
            inlineShape.Height = height;
            TrySetTitle(inlineShape, pictureName);
            inlineShape.Select();
            return WordPictureReference.FromInlineShape(inlineShape);
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
            string renderSvgPath = PrepareWordSvg(svgFilePath, snapshot.Width, snapshot.Height);
            WordPictureReference replacement;

            if (existingPicture.IsInline)
            {
                replacement = ReplaceInline(document, existingPicture, renderSvgPath, snapshot);
            }
            else
            {
                replacement = ReplaceFloating(document, existingPicture, renderSvgPath, snapshot);
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

            existingPicture.Delete();
            object linkToFile = false;
            object saveWithDocument = true;
            object targetRange = range;
            WordInterop.InlineShape inlineShape = document.InlineShapes.AddPicture(svgFilePath, ref linkToFile, ref saveWithDocument, ref targetRange);
            inlineShape.LockAspectRatio = snapshot.LockAspectRatio;
            if (snapshot.Width > 0f)
            {
                inlineShape.Width = snapshot.Width;
            }

            if (snapshot.Height > 0f)
            {
                inlineShape.Height = snapshot.Height;
            }

            TrySetTitle(inlineShape, snapshot.Title);
            return WordPictureReference.FromInlineShape(inlineShape);
        }

        private WordPictureReference ReplaceFloating(WordInterop.Document document, WordPictureReference existingPicture, string svgFilePath, PictureSnapshot snapshot)
        {
            WordInterop.Range anchor = existingPicture.GetRange();
            if (anchor == null)
            {
                throw new InvalidOperationException("Unable to resolve the selected picture anchor.");
            }

            existingPicture.Delete();
            object linkToFile = false;
            object saveWithDocument = true;
            object left = snapshot.Left;
            object top = snapshot.Top;
            object width = snapshot.Width;
            object height = snapshot.Height;
            object targetAnchor = anchor;
            WordInterop.Shape shape = document.Shapes.AddPicture(svgFilePath, ref linkToFile, ref saveWithDocument, ref left, ref top, ref width, ref height, ref targetAnchor);
            shape.LockAspectRatio = snapshot.LockAspectRatio;
            TrySetName(shape, snapshot.Name);
            TrySetTitle(shape, snapshot.Title);
            TrySetWrapType(shape, snapshot.WrapType);
            TrySetRelativeHorizontalPosition(shape, snapshot.RelativeHorizontalPosition);
            TrySetRelativeVerticalPosition(shape, snapshot.RelativeVerticalPosition);
            return WordPictureReference.FromShape(shape);
        }

        private string PrepareWordSvg(string svgFilePath, float targetWidth, float targetHeight)
        {
            if (_svgSupportService == null)
            {
                return svgFilePath;
            }

            return _svgSupportService.PrepareForPresentation(svgFilePath, targetWidth, targetHeight);
        }

        private static float ResolveDefaultWidth(WordInterop.Document document)
        {
            try
            {
                float usableWidth = document.PageSetup.PageWidth - document.PageSetup.LeftMargin - document.PageSetup.RightMargin;
                if (usableWidth > 0f)
                {
                    return usableWidth * 0.72f;
                }
            }
            catch
            {
            }

            return 360f;
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

            try
            {
                shape.WrapFormat.Type = wrapType;
            }
            catch
            {
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
                    try
                    {
                        snapshot.Left = picture.Shape.Left;
                        snapshot.Top = picture.Shape.Top;
                        snapshot.WrapType = picture.Shape.WrapFormat.Type;
                        snapshot.RelativeHorizontalPosition = picture.Shape.RelativeHorizontalPosition;
                        snapshot.RelativeVerticalPosition = picture.Shape.RelativeVerticalPosition;
                    }
                    catch
                    {
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
        }
    }
}
