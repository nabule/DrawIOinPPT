using System;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectionDoubleClickEventArgs : EventArgs
    {
        public SelectionDoubleClickEventArgs(PptInterop.Selection selection)
        {
            this.Selection = selection;
            this.Cancel = false;
        }

        public PptInterop.Selection Selection { get; private set; }
        public bool Cancel { get; set; }
    }
}

