using Microsoft.Office.Core;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.PowerPointAddIn.Ribbon
{
    public class RibbonController
    {
        private readonly AddInHost _host;
        private IRibbonUI _ribbonUi;

        public RibbonController(AddInHost host)
        {
            _host = host;
            _host.SelectionStateChanged += OnSelectionStateChanged;
        }

        public void OnRibbonLoad(IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
        }

        public void OnNewDiagram(IRibbonControl control)
        {
            _host.CreateNewDiagram();
        }

        public void OnEditSelectedDiagram(IRibbonControl control)
        {
            _host.EditSelectedDiagram();
        }

        public void OnBindSelectedShape(IRibbonControl control)
        {
            _host.BindSelectedShape();
        }

        public void OnClearSelectedShape(IRibbonControl control)
        {
            _host.ClearSelectedShapeBinding();
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            _host.OpenSettings();
        }

        public bool GetSingleShapeEnabled(IRibbonControl control)
        {
            return _host.HasSingleShapeSelection;
        }

        public bool GetEditEnabled(IRibbonControl control)
        {
            return _host.HasEditableSelection;
        }

        public string GetSelectionStatus(IRibbonControl control)
        {
            return _host.GetSelectionSummary();
        }

        private void OnSelectionStateChanged(object sender, System.EventArgs e)
        {
            if (_ribbonUi != null)
            {
                _ribbonUi.Invalidate();
            }
        }
    }
}
