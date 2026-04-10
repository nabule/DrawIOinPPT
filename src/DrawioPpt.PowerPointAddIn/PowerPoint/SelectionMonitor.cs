using System;
using PptInterop = Microsoft.Office.Interop.PowerPoint;

namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectionMonitor : IDisposable
    {
        private readonly PptInterop.Application _application;
        private readonly PowerPointShapeSelectionReader _reader;
        private bool _started;

        public SelectionMonitor(PptInterop.Application application, PowerPointShapeSelectionReader reader)
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

        private void OnWindowSelectionChange(PptInterop.Selection selection)
        {
            EventHandler<SelectionContextChangedEventArgs> handler = this.SelectionChanged;
            if (handler == null)
            {
                return;
            }

            SelectionContext context = _reader.Read(selection);
            handler(this, new SelectionContextChangedEventArgs(context));
        }

        private void OnWindowBeforeDoubleClick(PptInterop.Selection selection, ref bool cancel)
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
