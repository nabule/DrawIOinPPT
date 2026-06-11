namespace DrawioPpt.WordAddIn.Word
{
    public class SelectionContext
    {
        public SelectionContext()
        {
            this.HasSelection = false;
            this.HasSinglePicture = false;
            this.IsManagedPicture = false;
            this.PictureName = string.Empty;
            this.DiagramId = string.Empty;
            this.AlternativeText = string.Empty;
            this.ReferenceKey = string.Empty;
        }

        public bool HasSelection { get; set; }
        public bool HasSinglePicture { get; set; }
        public bool IsManagedPicture { get; set; }
        public string PictureName { get; set; }
        public string DiagramId { get; set; }
        public string AlternativeText { get; set; }
        public string ReferenceKey { get; set; }
    }
}
