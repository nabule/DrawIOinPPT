using System;
using System.IO;
using System.Text;
using DrawioPpt.Core.Models;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgMarkupFileStore
    {
        public string Write(DiagramEnvelope envelope, string svgMarkup)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            if (string.IsNullOrWhiteSpace(svgMarkup))
            {
                throw new ArgumentException("SVG markup is required.", "svgMarkup");
            }

            string targetDirectory = Path.Combine(Path.GetTempPath(), "DrawioPpt", "url-export");
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            string safeId = string.IsNullOrWhiteSpace(envelope.DiagramId) ? Guid.NewGuid().ToString("N") : envelope.DiagramId;
            string filePath = Path.Combine(targetDirectory, safeId + ".svg");
            File.WriteAllText(filePath, svgMarkup, Encoding.UTF8);
            return filePath;
        }
    }
}
