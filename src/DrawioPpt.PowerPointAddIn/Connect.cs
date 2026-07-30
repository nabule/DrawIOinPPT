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
    [Guid("9C0C05A5-FA1D-4B34-B7DD-7C5C8592D844")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPowerPointAddInAutomation
    {
        string GetVersionSummary();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PowerPointAddInAutomation : MarshalByRefObject, IPowerPointAddInAutomation
    {
        private readonly AddInHost _host;

        public PowerPointAddInAutomation(AddInHost host)
        {
            _host = host;
        }

        public string GetVersionSummary()
        {
            return _host == null ? "版本：初始化中" : _host.GetVersionSummary();
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }
    }

    [ComVisible(true)]
    [Guid("0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31")]
    [ProgId(AddInProgId)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public const string AddInProgId = "Greensoft.DrawioPptAddIn";

        private const string OfficeAddinsRegistryRoot = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\";
        private PptInterop.Application _application;
        private COMAddIn _comAddIn;
        private PowerPointAddInAutomation _automation;
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

            _comAddIn = addInInst as COMAddIn;
            _host = new AddInHost(_application);
            _ribbonController = new RibbonController(_host);
            _automation = new PowerPointAddInAutomation(_host);
            if (_comAddIn != null)
            {
                _comAddIn.Object = _automation;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (_host != null)
            {
                _host.Dispose();
            }

            if (_comAddIn != null)
            {
                try
                {
                    _comAddIn.Object = null;
                }
                catch
                {
                }
            }

            _ribbonController = null;
            _host = null;
            _automation = null;
            _comAddIn = null;
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

        public bool GetBindEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetBindEnabled(control);
        }

        public string GetEditorModeSummary(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return "当前模式：初始化中";
            }

            return _ribbonController.GetEditorModeSummary(control);
        }

        public string GetVersionSummary(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return "版本：初始化中";
            }

            return _ribbonController.GetVersionSummary(control);
        }

        public string GetSelectionDetailSummary(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return "状态：初始化中";
            }

            return _ribbonController.GetSelectionDetailSummary(control);
        }

        public string GetEditButtonLabel(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return "编辑";
            }

            return _ribbonController.GetEditButtonLabel(control);
        }

        public bool GetAutoOpenPressed(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetAutoOpenPressed(control);
        }

        public void OnToggleAutoOpen(IRibbonControl control, bool pressed)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnToggleAutoOpen(control, pressed);
            }
        }

        public object GetButtonImage(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return null;
            }

            return _ribbonController.GetButtonImage(control);
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            string keyName = OfficeAddinsRegistryRoot + AddInProgId;
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
            string keyName = OfficeAddinsRegistryRoot + AddInProgId;
            Registry.CurrentUser.DeleteSubKeyTree(keyName, false);
        }
    }
}
