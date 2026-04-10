using System;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectionContextChangedEventArgs : EventArgs
    {
        public SelectionContextChangedEventArgs(SelectionContext context)
        {
            this.Context = context;
        }

        public SelectionContext Context { get; private set; }
    }
}

