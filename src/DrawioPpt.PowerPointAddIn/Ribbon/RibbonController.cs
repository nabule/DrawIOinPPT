using Microsoft.Office.Core;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.PowerPointAddIn.Ribbon
{
    public class RibbonController
    {
        private readonly AddInHost _host;
        private readonly RibbonImageProvider _imageProvider;
        private IRibbonUI _ribbonUi;

        public RibbonController(AddInHost host)
        {
            _host = host;
            _imageProvider = new RibbonImageProvider();
            _host.SelectionStateChanged += OnSelectionStateChanged;
        }

        public void OnRibbonLoad(IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
            if (_ribbonUi != null)
            {
                _ribbonUi.Invalidate();
            }
        }

        public void OnNewDiagram(IRibbonControl control)
        {
            _host.CreateNewDiagram();
        }

        public void OnEditSelectedDiagram(IRibbonControl control)
        {
            _host.EditSelectedDiagram();
        }

        public void OnRefreshSelectedDiagram(IRibbonControl control)
        {
            _host.RefreshSelectedDiagram();
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

        public bool GetBindEnabled(IRibbonControl control)
        {
            return _host.CanBindSelection;
        }

        public bool GetManagedShapeEnabled(IRibbonControl control)
        {
            return _host.HasEditableSelection;
        }

        public bool GetEditEnabled(IRibbonControl control)
        {
            return _host.HasEditableSelection;
        }

        public string GetSelectionStatus(IRibbonControl control)
        {
            return _host.GetSelectionSummary();
        }

        public string GetSelectionDetailSummary(IRibbonControl control)
        {
            return _host.GetSelectionDetailSummary();
        }

        public string GetEditorModeSummary(IRibbonControl control)
        {
            return _host.GetEditorModeSummary();
        }

        public string GetEditButtonLabel(IRibbonControl control)
        {
            return _host.GetEditButtonLabel();
        }

        public bool GetAutoOpenPressed(IRibbonControl control)
        {
            return _host.IsAutoOpenOnSelectionEnabled;
        }

        public void OnToggleAutoOpen(IRibbonControl control, bool pressed)
        {
            _host.SetAutoOpenOnSelection(pressed);
        }

        public object GetButtonImage(IRibbonControl control)
        {
            if (control == null)
            {
                return null;
            }

            return _imageProvider.GetButtonImage(control.Id);
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
