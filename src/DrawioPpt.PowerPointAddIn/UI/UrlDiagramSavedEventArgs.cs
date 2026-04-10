using System;

namespace DrawioPpt.PowerPointAddIn.UI
{
    public class UrlDiagramSavedEventArgs : EventArgs
    {
        public UrlDiagramSavedEventArgs(string xml, string svgMarkup, bool exitRequested)
        {
            this.Xml = xml ?? string.Empty;
            this.SvgMarkup = svgMarkup ?? string.Empty;
            this.ExitRequested = exitRequested;
        }

        public string Xml { get; private set; }
        public string SvgMarkup { get; private set; }
        public bool ExitRequested { get; private set; }
    }
}
