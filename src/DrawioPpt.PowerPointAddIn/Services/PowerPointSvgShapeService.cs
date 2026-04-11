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
            TryPickupFormatting(existingShape);
            TryPickupAnimation(existingShape);
            existingShape.Delete();

            PptInterop.Shape newShape = slide.Shapes.AddPicture(
                svgFilePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                snapshot.Left,
                snapshot.Top,
                snapshot.Width,
                snapshot.Height);

            TryApplyFormatting(newShape);
            TryApplyAnimation(newShape);
            newShape.Rotation = snapshot.Rotation;
            TrySetShapeName(newShape, snapshot.Name);
            TrySetTitle(newShape, snapshot.Title);
            TrySetVisible(newShape, snapshot.Visible);
            TrySetLockAspectRatio(newShape, snapshot.LockAspectRatio);
            TrySetBlackWhiteMode(newShape, snapshot.BlackWhiteMode);
            TryApplyActionSetting(newShape, PptInterop.PpMouseActivation.ppMouseClick, snapshot.ClickAction);
            TryApplyActionSetting(newShape, PptInterop.PpMouseActivation.ppMouseOver, snapshot.MouseOverAction);
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

        private static void TrySetTitle(PptInterop.Shape shape, string title)
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

        private static void TrySetVisible(PptInterop.Shape shape, MsoTriState visible)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Visible = visible;
            }
            catch
            {
            }
        }

        private static void TrySetLockAspectRatio(PptInterop.Shape shape, MsoTriState lockAspectRatio)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.LockAspectRatio = lockAspectRatio;
            }
            catch
            {
            }
        }

        private static void TrySetBlackWhiteMode(PptInterop.Shape shape, MsoBlackWhiteMode blackWhiteMode)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.BlackWhiteMode = blackWhiteMode;
            }
            catch
            {
            }
        }

        private static void TryPickupFormatting(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.PickUp();
            }
            catch
            {
            }
        }

        private static void TryApplyFormatting(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Apply();
            }
            catch
            {
            }
        }

        private static void TryPickupAnimation(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.PickupAnimation();
            }
            catch
            {
            }
        }

        private static void TryApplyAnimation(PptInterop.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.ApplyAnimation();
            }
            catch
            {
            }
        }

        private static void TryApplyActionSetting(PptInterop.Shape shape, PptInterop.PpMouseActivation activation, ActionSettingSnapshot snapshot)
        {
            if (shape == null || snapshot == null)
            {
                return;
            }

            try
            {
                PptInterop.ActionSetting actionSetting = shape.ActionSettings[activation];
                TrySetActionType(actionSetting, snapshot.Action);
                TrySetAnimateAction(actionSetting, snapshot.AnimateAction);
                TrySetActionVerb(actionSetting, snapshot.ActionVerb);
                TrySetRun(actionSetting, snapshot.Run);
                TrySetShowAndReturn(actionSetting, snapshot.ShowAndReturn);
                TrySetSlideShowName(actionSetting, snapshot.SlideShowName);

                if (snapshot.Hyperlink != null)
                {
                    PptInterop.Hyperlink hyperlink = actionSetting.Hyperlink;
                    TrySetHyperlinkAddress(hyperlink, snapshot.Hyperlink.Address);
                    TrySetHyperlinkSubAddress(hyperlink, snapshot.Hyperlink.SubAddress);
                    TrySetHyperlinkScreenTip(hyperlink, snapshot.Hyperlink.ScreenTip);
                    TrySetHyperlinkTextToDisplay(hyperlink, snapshot.Hyperlink.TextToDisplay);
                    TrySetHyperlinkEmailSubject(hyperlink, snapshot.Hyperlink.EmailSubject);
                    TrySetHyperlinkShowAndReturn(hyperlink, snapshot.Hyperlink.ShowAndReturn);
                }
            }
            catch
            {
            }
        }

        private static void TrySetActionType(PptInterop.ActionSetting actionSetting, PptInterop.PpActionType action)
        {
            try
            {
                actionSetting.Action = action;
            }
            catch
            {
            }
        }

        private static void TrySetAnimateAction(PptInterop.ActionSetting actionSetting, MsoTriState animateAction)
        {
            try
            {
                actionSetting.AnimateAction = animateAction;
            }
            catch
            {
            }
        }

        private static void TrySetActionVerb(PptInterop.ActionSetting actionSetting, string actionVerb)
        {
            if (actionSetting == null || string.IsNullOrEmpty(actionVerb))
            {
                return;
            }

            try
            {
                actionSetting.ActionVerb = actionVerb;
            }
            catch
            {
            }
        }

        private static void TrySetRun(PptInterop.ActionSetting actionSetting, string run)
        {
            if (actionSetting == null || string.IsNullOrEmpty(run))
            {
                return;
            }

            try
            {
                actionSetting.Run = run;
            }
            catch
            {
            }
        }

        private static void TrySetShowAndReturn(PptInterop.ActionSetting actionSetting, MsoTriState showAndReturn)
        {
            try
            {
                actionSetting.ShowAndReturn = showAndReturn;
            }
            catch
            {
            }
        }

        private static void TrySetSlideShowName(PptInterop.ActionSetting actionSetting, string slideShowName)
        {
            if (actionSetting == null || string.IsNullOrEmpty(slideShowName))
            {
                return;
            }

            try
            {
                actionSetting.SlideShowName = slideShowName;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkAddress(PptInterop.Hyperlink hyperlink, string address)
        {
            if (hyperlink == null || string.IsNullOrEmpty(address))
            {
                return;
            }

            try
            {
                hyperlink.Address = address;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkSubAddress(PptInterop.Hyperlink hyperlink, string subAddress)
        {
            if (hyperlink == null || string.IsNullOrEmpty(subAddress))
            {
                return;
            }

            try
            {
                hyperlink.SubAddress = subAddress;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkScreenTip(PptInterop.Hyperlink hyperlink, string screenTip)
        {
            if (hyperlink == null || string.IsNullOrEmpty(screenTip))
            {
                return;
            }

            try
            {
                hyperlink.ScreenTip = screenTip;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkTextToDisplay(PptInterop.Hyperlink hyperlink, string textToDisplay)
        {
            if (hyperlink == null || string.IsNullOrEmpty(textToDisplay))
            {
                return;
            }

            try
            {
                hyperlink.TextToDisplay = textToDisplay;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkEmailSubject(PptInterop.Hyperlink hyperlink, string emailSubject)
        {
            if (hyperlink == null || string.IsNullOrEmpty(emailSubject))
            {
                return;
            }

            try
            {
                hyperlink.EmailSubject = emailSubject;
            }
            catch
            {
            }
        }

        private static void TrySetHyperlinkShowAndReturn(PptInterop.Hyperlink hyperlink, MsoTriState showAndReturn)
        {
            try
            {
                hyperlink.ShowAndReturn = showAndReturn;
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
            public string Title { get; private set; }
            public MsoTriState Visible { get; private set; }
            public MsoTriState LockAspectRatio { get; private set; }
            public MsoBlackWhiteMode BlackWhiteMode { get; private set; }
            public ActionSettingSnapshot ClickAction { get; private set; }
            public ActionSettingSnapshot MouseOverAction { get; private set; }

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
                snapshot.Title = TryGetTitle(shape);
                snapshot.Visible = TryGetVisible(shape);
                snapshot.LockAspectRatio = TryGetLockAspectRatio(shape);
                snapshot.BlackWhiteMode = TryGetBlackWhiteMode(shape);
                snapshot.ClickAction = ActionSettingSnapshot.Capture(shape, PptInterop.PpMouseActivation.ppMouseClick);
                snapshot.MouseOverAction = ActionSettingSnapshot.Capture(shape, PptInterop.PpMouseActivation.ppMouseOver);
                return snapshot;
            }

            private static string TryGetTitle(PptInterop.Shape shape)
            {
                try
                {
                    return shape.Title ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            private static MsoTriState TryGetVisible(PptInterop.Shape shape)
            {
                try
                {
                    return shape.Visible;
                }
                catch
                {
                    return MsoTriState.msoTrue;
                }
            }

            private static MsoTriState TryGetLockAspectRatio(PptInterop.Shape shape)
            {
                try
                {
                    return shape.LockAspectRatio;
                }
                catch
                {
                    return MsoTriState.msoFalse;
                }
            }

            private static MsoBlackWhiteMode TryGetBlackWhiteMode(PptInterop.Shape shape)
            {
                try
                {
                    return shape.BlackWhiteMode;
                }
                catch
                {
                    return MsoBlackWhiteMode.msoBlackWhiteMixed;
                }
            }
        }

        private sealed class ActionSettingSnapshot
        {
            public PptInterop.PpActionType Action { get; private set; }
            public string ActionVerb { get; private set; }
            public MsoTriState AnimateAction { get; private set; }
            public string Run { get; private set; }
            public MsoTriState ShowAndReturn { get; private set; }
            public string SlideShowName { get; private set; }
            public HyperlinkSnapshot Hyperlink { get; private set; }

            public static ActionSettingSnapshot Capture(PptInterop.Shape shape, PptInterop.PpMouseActivation activation)
            {
                if (shape == null)
                {
                    return null;
                }

                ActionSettingSnapshot snapshot = new ActionSettingSnapshot();

                try
                {
                    PptInterop.ActionSetting actionSetting = shape.ActionSettings[activation];
                    snapshot.Action = TryGetActionType(actionSetting);
                    snapshot.ActionVerb = TryGetActionVerb(actionSetting);
                    snapshot.AnimateAction = TryGetAnimateAction(actionSetting);
                    snapshot.Run = TryGetRun(actionSetting);
                    snapshot.ShowAndReturn = TryGetShowAndReturn(actionSetting);
                    snapshot.SlideShowName = TryGetSlideShowName(actionSetting);
                    snapshot.Hyperlink = HyperlinkSnapshot.Capture(actionSetting.Hyperlink);
                }
                catch
                {
                    return null;
                }

                return snapshot;
            }

            private static PptInterop.PpActionType TryGetActionType(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.Action;
                }
                catch
                {
                    return PptInterop.PpActionType.ppActionNone;
                }
            }

            private static string TryGetActionVerb(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.ActionVerb ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            private static MsoTriState TryGetAnimateAction(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.AnimateAction;
                }
                catch
                {
                    return MsoTriState.msoFalse;
                }
            }

            private static string TryGetRun(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.Run ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            private static MsoTriState TryGetShowAndReturn(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.ShowAndReturn;
                }
                catch
                {
                    return MsoTriState.msoFalse;
                }
            }

            private static string TryGetSlideShowName(PptInterop.ActionSetting actionSetting)
            {
                try
                {
                    return actionSetting.SlideShowName ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        private sealed class HyperlinkSnapshot
        {
            public string Address { get; private set; }
            public string SubAddress { get; private set; }
            public string ScreenTip { get; private set; }
            public string TextToDisplay { get; private set; }
            public string EmailSubject { get; private set; }
            public MsoTriState ShowAndReturn { get; private set; }

            public static HyperlinkSnapshot Capture(PptInterop.Hyperlink hyperlink)
            {
                if (hyperlink == null)
                {
                    return null;
                }

                HyperlinkSnapshot snapshot = new HyperlinkSnapshot();

                try
                {
                    snapshot.Address = hyperlink.Address ?? string.Empty;
                }
                catch
                {
                    snapshot.Address = string.Empty;
                }

                try
                {
                    snapshot.SubAddress = hyperlink.SubAddress ?? string.Empty;
                }
                catch
                {
                    snapshot.SubAddress = string.Empty;
                }

                try
                {
                    snapshot.ScreenTip = hyperlink.ScreenTip ?? string.Empty;
                }
                catch
                {
                    snapshot.ScreenTip = string.Empty;
                }

                try
                {
                    snapshot.TextToDisplay = hyperlink.TextToDisplay ?? string.Empty;
                }
                catch
                {
                    snapshot.TextToDisplay = string.Empty;
                }

                try
                {
                    snapshot.EmailSubject = hyperlink.EmailSubject ?? string.Empty;
                }
                catch
                {
                    snapshot.EmailSubject = string.Empty;
                }

                try
                {
                    snapshot.ShowAndReturn = hyperlink.ShowAndReturn;
                }
                catch
                {
                    snapshot.ShowAndReturn = MsoTriState.msoFalse;
                }

                return snapshot;
            }
        }
    }
}
