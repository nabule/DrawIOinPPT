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
    }
}

