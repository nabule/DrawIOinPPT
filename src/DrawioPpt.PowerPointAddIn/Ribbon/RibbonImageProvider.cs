using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DrawioPpt.PowerPointAddIn.Ribbon
{
    public class RibbonImageProvider
    {
        private readonly Dictionary<string, Bitmap> _cache;

        public RibbonImageProvider()
        {
            _cache = new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        }

        public object GetButtonImage(string controlId)
        {
            if (string.IsNullOrWhiteSpace(controlId))
            {
                return null;
            }

            Bitmap image;
            if (!_cache.TryGetValue(controlId, out image))
            {
                image = CreateImage(controlId);
                _cache[controlId] = image;
            }

            return PictureDispConverter.ToPictureDisp(image);
        }

        private static Bitmap CreateImage(string controlId)
        {
            Bitmap bitmap = new Bitmap(48, 48);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);

                Color accent = ResolveAccent(controlId);
                using (GraphicsPath background = BuildRoundedRectangle(new RectangleF(4f, 4f, 40f, 40f), 12f))
                using (SolidBrush brush = new SolidBrush(accent))
                using (Pen border = new Pen(Color.FromArgb(35, 0, 0, 0), 1.2f))
                {
                    graphics.FillPath(brush, background);
                    graphics.DrawPath(border, background);
                }

                using (Pen primaryPen = new Pen(Color.White, 3f))
                using (Pen secondaryPen = new Pen(Color.FromArgb(230, 255, 255, 255), 2.2f))
                using (SolidBrush primaryBrush = new SolidBrush(Color.White))
                {
                    primaryPen.StartCap = LineCap.Round;
                    primaryPen.EndCap = LineCap.Round;
                    secondaryPen.StartCap = LineCap.Round;
                    secondaryPen.EndCap = LineCap.Round;

                    if (controlId.EndsWith(".New", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawNewIcon(graphics, primaryPen, secondaryPen, primaryBrush);
                    }
                    else if (controlId.EndsWith(".Edit", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawEditIcon(graphics, primaryPen, secondaryPen, primaryBrush);
                    }
                    else if (controlId.EndsWith(".Refresh", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawRefreshIcon(graphics, primaryPen, primaryBrush);
                    }
                    else if (controlId.EndsWith(".Bind", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawBindIcon(graphics, primaryPen);
                    }
                    else if (controlId.EndsWith(".Clear", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawClearIcon(graphics, primaryPen, secondaryPen);
                    }
                    else if (controlId.EndsWith(".Settings", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawSettingsIcon(graphics, primaryPen, primaryBrush);
                    }
                    else if (controlId.EndsWith(".AutoOpen", StringComparison.OrdinalIgnoreCase))
                    {
                        DrawAutoOpenIcon(graphics, primaryPen, secondaryPen, primaryBrush);
                    }
                    else
                    {
                        graphics.FillEllipse(primaryBrush, 18f, 18f, 12f, 12f);
                    }
                }
            }

            return bitmap;
        }

        private static void DrawNewIcon(Graphics graphics, Pen primaryPen, Pen secondaryPen, SolidBrush primaryBrush)
        {
            graphics.DrawRectangle(secondaryPen, 12f, 11f, 17f, 22f);
            graphics.DrawLine(secondaryPen, 23f, 11f, 29f, 17f);
            graphics.DrawLine(secondaryPen, 22f, 18f, 29f, 18f);
            graphics.DrawLine(primaryPen, 31f, 29f, 39f, 29f);
            graphics.DrawLine(primaryPen, 35f, 25f, 35f, 33f);
            graphics.FillEllipse(primaryBrush, 32.5f, 26.5f, 5f, 5f);
        }

        private static void DrawEditIcon(Graphics graphics, Pen primaryPen, Pen secondaryPen, SolidBrush primaryBrush)
        {
            graphics.DrawRectangle(secondaryPen, 11f, 28f, 18f, 8f);
            graphics.DrawLine(primaryPen, 16f, 31f, 31f, 16f);
            graphics.DrawLine(primaryPen, 29f, 14f, 34f, 19f);
            PointF[] tip =
            {
                new PointF(30.5f, 14.5f),
                new PointF(35.5f, 19.5f),
                new PointF(37.5f, 12.5f)
            };
            graphics.FillPolygon(primaryBrush, tip);
        }

        private static void DrawRefreshIcon(Graphics graphics, Pen primaryPen, SolidBrush primaryBrush)
        {
            graphics.DrawArc(primaryPen, 11f, 11f, 26f, 26f, 35f, 250f);
            PointF[] arrow =
            {
                new PointF(34f, 12f),
                new PointF(38f, 20f),
                new PointF(29f, 18f)
            };
            graphics.FillPolygon(primaryBrush, arrow);
            graphics.DrawArc(primaryPen, 12f, 12f, 24f, 24f, 215f, 70f);
        }

        private static void DrawBindIcon(Graphics graphics, Pen primaryPen)
        {
            graphics.DrawArc(primaryPen, 10f, 16f, 14f, 14f, 300f, 250f);
            graphics.DrawArc(primaryPen, 24f, 16f, 14f, 14f, 120f, 250f);
            graphics.DrawLine(primaryPen, 20f, 20f, 28f, 20f);
            graphics.DrawLine(primaryPen, 20f, 26f, 28f, 26f);
        }

        private static void DrawClearIcon(Graphics graphics, Pen primaryPen, Pen secondaryPen)
        {
            graphics.DrawArc(secondaryPen, 10f, 16f, 14f, 14f, 300f, 250f);
            graphics.DrawArc(secondaryPen, 24f, 16f, 14f, 14f, 120f, 250f);
            graphics.DrawLine(primaryPen, 18f, 30f, 31f, 17f);
            graphics.DrawLine(primaryPen, 18f, 17f, 31f, 30f);
        }

        private static void DrawSettingsIcon(Graphics graphics, Pen primaryPen, SolidBrush primaryBrush)
        {
            graphics.FillEllipse(primaryBrush, 20f, 20f, 8f, 8f);
            graphics.DrawEllipse(primaryPen, 15f, 15f, 18f, 18f);
            graphics.DrawLine(primaryPen, 24f, 10f, 24f, 14f);
            graphics.DrawLine(primaryPen, 24f, 34f, 24f, 38f);
            graphics.DrawLine(primaryPen, 10f, 24f, 14f, 24f);
            graphics.DrawLine(primaryPen, 34f, 24f, 38f, 24f);
            graphics.DrawLine(primaryPen, 14f, 14f, 17f, 17f);
            graphics.DrawLine(primaryPen, 31f, 31f, 34f, 34f);
            graphics.DrawLine(primaryPen, 31f, 17f, 34f, 14f);
            graphics.DrawLine(primaryPen, 14f, 34f, 17f, 31f);
        }

        private static void DrawAutoOpenIcon(Graphics graphics, Pen primaryPen, Pen secondaryPen, SolidBrush primaryBrush)
        {
            graphics.DrawRectangle(secondaryPen, 10f, 13f, 13f, 22f);
            graphics.DrawRectangle(secondaryPen, 26f, 16f, 11f, 16f);
            graphics.DrawLine(primaryPen, 18f, 24f, 28f, 24f);
            PointF[] arrow =
            {
                new PointF(26f, 19f),
                new PointF(33f, 24f),
                new PointF(26f, 29f)
            };
            graphics.FillPolygon(primaryBrush, arrow);
        }

        private static GraphicsPath BuildRoundedRectangle(RectangleF bounds, float radius)
        {
            float diameter = radius * 2f;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180f, 90f);
            path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270f, 90f);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0f, 90f);
            path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90f, 90f);
            path.CloseFigure();
            return path;
        }

        private static Color ResolveAccent(string controlId)
        {
            if (controlId.EndsWith(".New", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(44, 138, 89);
            }

            if (controlId.EndsWith(".Edit", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(38, 112, 196);
            }

            if (controlId.EndsWith(".Refresh", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(14, 138, 143);
            }

            if (controlId.EndsWith(".Bind", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(215, 132, 38);
            }

            if (controlId.EndsWith(".Clear", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(194, 68, 68);
            }

            if (controlId.EndsWith(".Settings", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(93, 102, 115);
            }

            if (controlId.EndsWith(".AutoOpen", StringComparison.OrdinalIgnoreCase))
            {
                return Color.FromArgb(88, 103, 215);
            }

            return Color.FromArgb(38, 112, 196);
        }

        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter()
                : base("00020400-0000-0000-C000-000000000046")
            {
            }

            public static object ToPictureDisp(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
