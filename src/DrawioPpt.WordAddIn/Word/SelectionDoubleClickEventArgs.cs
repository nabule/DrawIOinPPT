using System;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class SelectionDoubleClickEventArgs : EventArgs
    {
        public SelectionDoubleClickEventArgs(WordInterop.Selection selection)
        {
            this.Selection = selection;
            this.Cancel = false;
        }

        public WordInterop.Selection Selection { get; private set; }
        public bool Cancel { get; set; }
    }
}
