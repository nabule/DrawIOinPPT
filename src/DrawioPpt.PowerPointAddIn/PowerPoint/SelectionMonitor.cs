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

        public void Start()
        {
            if (_started || _application == null)
            {
                return;
            }

            _application.WindowSelectionChange += OnWindowSelectionChange;
            _started = true;
        }

        public void Dispose()
        {
            if (_started && _application != null)
            {
                _application.WindowSelectionChange -= OnWindowSelectionChange;
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
    }
}
