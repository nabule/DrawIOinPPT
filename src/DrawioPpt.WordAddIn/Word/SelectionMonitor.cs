using System;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class SelectionMonitor : IDisposable
    {
        private readonly WordInterop.Application _application;
        private bool _started;

        public SelectionMonitor(WordInterop.Application application)
        {
            _application = application;
        }

        public event EventHandler<SelectionDoubleClickEventArgs> SelectionDoubleClicked;

        public void Start()
        {
            if (_started || _application == null)
            {
                return;
            }

            _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
            _started = true;
        }

        public void Dispose()
        {
            if (_started && _application != null)
            {
                _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick;
                _started = false;
            }
        }

        private void OnWindowBeforeDoubleClick(WordInterop.Selection selection, ref bool cancel)
        {
            EventHandler<SelectionDoubleClickEventArgs> handler = this.SelectionDoubleClicked;
            if (handler == null)
            {
                return;
            }

            SelectionDoubleClickEventArgs args = new SelectionDoubleClickEventArgs(selection);
            handler(this, args);
            cancel = args.Cancel;
        }
    }
}
