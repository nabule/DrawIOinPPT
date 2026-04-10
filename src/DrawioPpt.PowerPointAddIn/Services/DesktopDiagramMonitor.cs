using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class DesktopDiagramMonitor : IDisposable
    {
        private readonly Action<string, string> _refreshAction;
        private readonly Dictionary<string, MonitoredDiagram> _monitoredDiagrams;
        private readonly Timer _timer;

        public DesktopDiagramMonitor(Action<string, string> refreshAction)
        {
            _refreshAction = refreshAction;
            _monitoredDiagrams = new Dictionary<string, MonitoredDiagram>(StringComparer.OrdinalIgnoreCase);
            _timer = new Timer();
            _timer.Interval = 1500;
            _timer.Tick += OnTimerTick;
            _timer.Start();
        }

        public void Track(string diagramId, string filePath)
        {
            if (string.IsNullOrWhiteSpace(diagramId) || string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            MonitoredDiagram item = new MonitoredDiagram();
            item.DiagramId = diagramId;
            item.FilePath = filePath;
            item.LastSeenWriteUtc = File.Exists(filePath) ? File.GetLastWriteTimeUtc(filePath) : DateTime.MinValue;
            _monitoredDiagrams[diagramId] = item;
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Tick -= OnTimerTick;
            _timer.Dispose();
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            List<MonitoredDiagram> snapshot = new List<MonitoredDiagram>(_monitoredDiagrams.Values);
            int index;
            for (index = 0; index < snapshot.Count; index++)
            {
                MonitoredDiagram item = snapshot[index];
                if (!File.Exists(item.FilePath))
                {
                    continue;
                }

                DateTime lastWriteUtc = File.GetLastWriteTimeUtc(item.FilePath);
                if (lastWriteUtc <= item.LastSeenWriteUtc)
                {
                    continue;
                }

                try
                {
                    _refreshAction(item.DiagramId, item.FilePath);
                    item.LastSeenWriteUtc = lastWriteUtc;
                }
                catch
                {
                }
            }
        }

        private sealed class MonitoredDiagram
        {
            public string DiagramId { get; set; }
            public string FilePath { get; set; }
            public DateTime LastSeenWriteUtc { get; set; }
        }
    }
}
