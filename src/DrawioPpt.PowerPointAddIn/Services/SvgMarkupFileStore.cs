using System;
using System.IO;
using System.Text;
using DrawioPpt.Core.Models;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class SvgMarkupFileStore
    {
        private readonly SvgDiagramDataEmbedder _embedder;
        private readonly SvgPowerPointSupportService _svgSupportService;

        public SvgMarkupFileStore()
            : this(new SvgDiagramDataEmbedder(), new SvgPowerPointSupportService())
        {
        }

        public SvgMarkupFileStore(SvgDiagramDataEmbedder embedder, SvgPowerPointSupportService svgSupportService)
        {
            _embedder = embedder;
            _svgSupportService = svgSupportService;
        }

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
            string finalMarkup = _embedder != null
                ? _embedder.EnsureEmbedded(svgMarkup, envelope.DrawioXml)
                : svgMarkup;
            if (_svgSupportService != null)
            {
                finalMarkup = _svgSupportService.SanitizeMarkup(finalMarkup);
            }

            File.WriteAllText(filePath, finalMarkup, Encoding.UTF8);
            return filePath;
        }
    }
}
