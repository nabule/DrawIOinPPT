using DrawioPpt.PowerPointAddIn.Ribbon;
using DrawioPpt.WordAddIn.Services;
using Microsoft.Office.Core;

namespace DrawioPpt.WordAddIn.Ribbon
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
            if (_host != null)
            {
                _host.SelectionStateChanged += OnHostSelectionStateChanged;
            }
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
            if (_host != null)
            {
                _host.CreateNewDiagram();
            }
        }

        public void OnEditSelectedDiagram(IRibbonControl control)
        {
            if (_host != null)
            {
                _host.EditSelectedDiagram();
            }
        }

        public void OnRefreshSelectedDiagram(IRibbonControl control)
        {
            if (_host != null)
            {
                _host.RefreshSelectedDiagram();
            }
        }

        public void OnBindSelectedPicture(IRibbonControl control)
        {
            if (_host != null)
            {
                _host.BindSelectedPicture();
            }
        }

        public void OnClearSelectedPicture(IRibbonControl control)
        {
            if (_host != null)
            {
                _host.ClearSelectedPictureBinding();
            }
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            if (_host != null)
            {
                _host.OpenSettings();
            }
        }

        public bool GetEditEnabled(IRibbonControl control)
        {
            return _host != null;
        }

        public string GetSelectionStatus(IRibbonControl control)
        {
            return _host == null ? "对象：初始化中" : _host.GetSelectionSummary();
        }

        public bool GetManagedPictureEnabled(IRibbonControl control)
        {
            return _host != null;
        }

        public bool GetSinglePictureEnabled(IRibbonControl control)
        {
            return _host != null;
        }

        public bool GetBindEnabled(IRibbonControl control)
        {
            return _host != null;
        }

        public string GetEditorModeSummary(IRibbonControl control)
        {
            return _host == null ? "模式：初始化中" : _host.GetEditorModeSummary();
        }

        public string GetSelectionDetailSummary(IRibbonControl control)
        {
            return _host == null ? "状态：初始化中" : _host.GetSelectionDetailSummary();
        }

        public string GetEditButtonLabel(IRibbonControl control)
        {
            return _host == null ? "编辑" : _host.GetEditButtonLabel();
        }

        public object GetButtonImage(IRibbonControl control)
        {
            return _imageProvider.GetButtonImage(control == null ? string.Empty : control.Id);
        }

        private void OnHostSelectionStateChanged(object sender, System.EventArgs e)
        {
            if (_ribbonUi != null)
            {
                _ribbonUi.Invalidate();
            }
        }
    }
}
