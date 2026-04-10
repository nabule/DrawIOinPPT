namespace DrawioPpt.PowerPointAddIn.PowerPoint
{
    public class SelectionContext
    {
        public SelectionContext()
        {
            this.HasSelection = false;
            this.HasSingleShape = false;
            this.IsManagedShape = false;
            this.ShapeId = 0;
            this.ShapeName = string.Empty;
            this.DiagramId = string.Empty;
            this.AlternativeText = string.Empty;
        }

        public bool HasSelection { get; set; }
        public bool HasSingleShape { get; set; }
        public bool IsManagedShape { get; set; }
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public string DiagramId { get; set; }
        public string AlternativeText { get; set; }
    }
}

