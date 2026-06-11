using System;

namespace DrawioPpt.WordAddIn.Word
{
    public class SelectionContextChangedEventArgs : EventArgs
    {
        public SelectionContextChangedEventArgs(SelectionContext context)
        {
            this.Context = context ?? new SelectionContext();
        }

        public SelectionContext Context { get; private set; }
    }
}
