using System;

namespace DrawioPpt.Core.Models
{
    public class DiagramEnvelope
    {
        public DiagramEnvelope()
        {
            this.FormatVersion = KnownMetadata.EnvelopeVersion;
            this.UpdatedUtc = DateTime.UtcNow;
            this.EditorMode = EditorMode.Desktop;
        }

        public string FormatVersion { get; set; }
        public string DiagramId { get; set; }
        public string DiagramName { get; set; }
        public EditorMode EditorMode { get; set; }
        public string EditorTarget { get; set; }
        public string SidecarPath { get; set; }
        public DateTime UpdatedUtc { get; set; }
        public string DrawioXml { get; set; }
    }
}

