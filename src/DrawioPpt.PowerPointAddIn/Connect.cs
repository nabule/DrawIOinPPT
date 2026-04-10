using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PptInterop = Microsoft.Office.Interop.PowerPoint;
using DrawioPpt.PowerPointAddIn.Ribbon;
using DrawioPpt.PowerPointAddIn.Services;

namespace DrawioPpt.PowerPointAddIn
{
    [ComVisible(true)]
    [Guid("0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31")]
    [ProgId("Greensoft.DrawioPptAddIn")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private PptInterop.Application _application;
        private AddInHost _host;
        private RibbonController _ribbonController;

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXmlProvider.GetCustomUi();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _application = application as PptInterop.Application;
            if (_application == null)
            {
                return;
            }

            _host = new AddInHost(_application);
            _ribbonController = new RibbonController(_host);
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (_host != null)
            {
                _host.Dispose();
            }

            _ribbonController = null;
            _host = null;
            _application = null;
        }

        public void OnStartupComplete(ref Array custom)
        {
            if (_host != null)
            {
                _host.Start();
            }
        }

        public void OnRibbonLoad(IRibbonUI ribbonUi)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnRibbonLoad(ribbonUi);
            }
        }

        public void OnNewDiagram(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnNewDiagram(control);
            }
        }

        public void OnEditSelectedDiagram(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnEditSelectedDiagram(control);
            }
        }

        public void OnRefreshSelectedDiagram(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnRefreshSelectedDiagram(control);
            }
        }

        public void OnBindSelectedShape(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnBindSelectedShape(control);
            }
        }

        public void OnClearSelectedShape(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnClearSelectedShape(control);
            }
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnOpenSettings(control);
            }
        }

        public bool GetEditEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetEditEnabled(control);
        }

        public string GetSelectionStatus(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return "Draw.io: 初始化中";
            }

            return _ribbonController.GetSelectionStatus(control);
        }

        public bool GetManagedShapeEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetManagedShapeEnabled(control);
        }

        public bool GetSingleShapeEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetSingleShapeEnabled(control);
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            string keyName = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\" + type.FullName;
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyName))
            {
                if (key == null)
                {
                    throw new InvalidOperationException("Unable to create PowerPoint Add-in registry key.");
                }

                key.SetValue("FriendlyName", "DrawioPpt");
                key.SetValue("Description", "Edit Draw.io diagrams inside PowerPoint.");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            string keyName = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\" + type.FullName;
            Registry.CurrentUser.DeleteSubKeyTree(keyName, false);
        }
    }
}
