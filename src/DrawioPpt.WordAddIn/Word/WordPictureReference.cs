using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class WordPictureReference
    {
        private readonly WordInterop.InlineShape _inlineShape;
        private readonly WordInterop.Shape _shape;

        private WordPictureReference(WordInterop.InlineShape inlineShape, WordInterop.Shape shape)
        {
            _inlineShape = inlineShape;
            _shape = shape;
        }

        public static WordPictureReference FromInlineShape(WordInterop.InlineShape inlineShape)
        {
            return inlineShape == null ? null : new WordPictureReference(inlineShape, null);
        }

        public static WordPictureReference FromShape(WordInterop.Shape shape)
        {
            return shape == null ? null : new WordPictureReference(null, shape);
        }

        public bool IsInline
        {
            get { return _inlineShape != null; }
        }

        public WordInterop.InlineShape InlineShape
        {
            get { return _inlineShape; }
        }

        public WordInterop.Shape Shape
        {
            get { return _shape; }
        }

        public string ReferenceKey
        {
            get
            {
                if (_shape != null)
                {
                    try
                    {
                        return "shape:" + _shape.ID.ToString();
                    }
                    catch
                    {
                        return "shape";
                    }
                }

                if (_inlineShape != null)
                {
                    WordInterop.Range sourceRange = null;
                    try
                    {
                        sourceRange = _inlineShape.Range;
                        return "inline:" +
                            sourceRange.Start.ToString();
                    }
                    catch
                    {
                        return "inline";
                    }
                    finally
                    {
                        SafeReleaseComObject(sourceRange);
                    }
                }

                return string.Empty;
            }
        }

        public string Name
        {
            get
            {
                if (_shape != null)
                {
                    try
                    {
                        return _shape.Name ?? string.Empty;
                    }
                    catch
                    {
                        return string.Empty;
                    }
                }

                return "Inline Picture";
            }
            set
            {
                if (_shape == null || string.IsNullOrWhiteSpace(value))
                {
                    return;
                }

                try
                {
                    _shape.Name = value;
                }
                catch
                {
                }
            }
        }

        public string AlternativeText
        {
            get
            {
                try
                {
                    return _shape != null ? (_shape.AlternativeText ?? string.Empty) : (_inlineShape.AlternativeText ?? string.Empty);
                }
                catch
                {
                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    if (_shape != null)
                    {
                        _shape.AlternativeText = value ?? string.Empty;
                    }
                    else if (_inlineShape != null)
                    {
                        _inlineShape.AlternativeText = value ?? string.Empty;
                    }
                }
                catch
                {
                }
            }
        }

        public string Title
        {
            get
            {
                try
                {
                    return _shape != null ? (_shape.Title ?? string.Empty) : (_inlineShape.Title ?? string.Empty);
                }
                catch
                {
                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    if (_shape != null)
                    {
                        _shape.Title = value ?? string.Empty;
                    }
                    else if (_inlineShape != null)
                    {
                        _inlineShape.Title = value ?? string.Empty;
                    }
                }
                catch
                {
                }
            }
        }

        public float Width
        {
            get
            {
                try
                {
                    return _shape != null ? _shape.Width : _inlineShape.Width;
                }
                catch
                {
                    return 0f;
                }
            }
            set
            {
                try
                {
                    if (_shape != null)
                    {
                        _shape.Width = value;
                    }
                    else if (_inlineShape != null)
                    {
                        _inlineShape.Width = value;
                    }
                }
                catch
                {
                }
            }
        }

        public float Height
        {
            get
            {
                try
                {
                    return _shape != null ? _shape.Height : _inlineShape.Height;
                }
                catch
                {
                    return 0f;
                }
            }
            set
            {
                try
                {
                    if (_shape != null)
                    {
                        _shape.Height = value;
                    }
                    else if (_inlineShape != null)
                    {
                        _inlineShape.Height = value;
                    }
                }
                catch
                {
                }
            }
        }

        public MsoTriState LockAspectRatio
        {
            get
            {
                try
                {
                    return _shape != null ? _shape.LockAspectRatio : _inlineShape.LockAspectRatio;
                }
                catch
                {
                    return MsoTriState.msoTrue;
                }
            }
            set
            {
                try
                {
                    if (_shape != null)
                    {
                        _shape.LockAspectRatio = value;
                    }
                    else if (_inlineShape != null)
                    {
                        _inlineShape.LockAspectRatio = value;
                    }
                }
                catch
                {
                }
            }
        }

        public WordInterop.Range GetRange()
        {
            WordInterop.Range sourceRange = null;
            try
            {
                if (_inlineShape != null)
                {
                    sourceRange = _inlineShape.Range;
                    return sourceRange.Duplicate;
                }

                if (_shape != null)
                {
                    sourceRange = _shape.Anchor;
                    if (sourceRange != null)
                    {
                        return sourceRange.Duplicate;
                    }
                }
            }
            catch
            {
            }
            finally
            {
                SafeReleaseComObject(sourceRange);
            }

            return null;
        }

        private static void SafeReleaseComObject(
            object comObject)
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

        public void Delete()
        {
            try
            {
                if (_shape != null)
                {
                    _shape.Delete();
                }
                else if (_inlineShape != null)
                {
                    _inlineShape.Delete();
                }
            }
            catch
            {
            }
        }

        public void Select()
        {
            try
            {
                if (_shape != null)
                {
                    _shape.Select(Type.Missing);
                }
                else if (_inlineShape != null)
                {
                    _inlineShape.Select();
                }
            }
            catch
            {
            }
        }
    }
}
