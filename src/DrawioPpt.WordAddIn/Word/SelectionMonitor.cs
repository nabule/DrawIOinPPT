using System;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class SelectionMonitor : IDisposable
    {
        private readonly WordInterop.Application _application;
        private readonly WordPictureSelectionReader _reader;
        private bool _started;

        public SelectionMonitor(WordInterop.Application application, WordPictureSelectionReader reader)
        {
            _application = application;
            _reader = reader;
        }

        public event EventHandler<SelectionContextChangedEventArgs> SelectionChanged;
        public event EventHandler<SelectionDoubleClickEventArgs> SelectionDoubleClicked;

        public void Start()
        {
            if (_started || _application == null)
            {
                return;
            }

            _application.WindowSelectionChange += OnWindowSelectionChange;
            _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
            _started = true;
        }

        public void Dispose()
        {
            if (_started && _application != null)
            {
                _application.WindowSelectionChange -= OnWindowSelectionChange;
                _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick;
                _started = false;
            }
        }

        private void OnWindowSelectionChange(WordInterop.Selection selection)
        {
            EventHandler<SelectionContextChangedEventArgs> handler = this.SelectionChanged;
            if (handler == null)
            {
                return;
            }

            SelectionContext context = _reader.Read(selection);
            handler(this, new SelectionContextChangedEventArgs(context));
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
