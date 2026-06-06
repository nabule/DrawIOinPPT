namespace DrawioPpt.Core.Models
{
    public class PluginSettings
    {
        public PluginSettings()
        {
            this.EditorMode = EditorMode.Desktop;
            this.DesktopEditorPath = string.Empty;
            this.EditorUrl = "https://embed.diagrams.net/?embed=1&proto=json&spin=1";
            this.UseOfficeCompatibleSvgLabels = true;
            this.AutoOpenOnSelection = false;
            this.AutoUpdateOnSave = true;
            this.KeepSidecarFile = true;
            this.SidecarFolderName = "drawio";
            this.ShowDiagramInfoDialog = true;
        }

        public EditorMode EditorMode { get; set; }
        public string DesktopEditorPath { get; set; }
        public string EditorUrl { get; set; }
        public bool UseOfficeCompatibleSvgLabels { get; set; }
        public bool AutoOpenOnSelection { get; set; }
        public bool AutoUpdateOnSave { get; set; }
        public bool KeepSidecarFile { get; set; }
        public string SidecarFolderName { get; set; }
        public bool ShowDiagramInfoDialog { get; set; }
    }
}
