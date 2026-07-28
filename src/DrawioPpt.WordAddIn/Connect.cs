using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DrawioPpt.WordAddIn.Ribbon;
using DrawioPpt.WordAddIn.Services;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn
{
    [ComVisible(true)]
    [Guid("3348835A-5E22-44BE-9C00-60328868C5FA")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IWordAddInAutomation
    {
        string GetInteractionMode();
        void EditSelectedDiagram();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class WordAddInAutomation : MarshalByRefObject, IWordAddInAutomation
    {
        private readonly AddInHost _host;
        private readonly Control _dispatcher;

        public WordAddInAutomation(AddInHost host, Control dispatcher)
        {
            _host = host;
            _dispatcher = dispatcher;
        }

        public void EditSelectedDiagram()
        {
            if (_host == null || _dispatcher == null || _dispatcher.IsDisposed)
            {
                return;
            }

            if (_dispatcher.InvokeRequired)
            {
                _dispatcher.Invoke(new Action(_host.EditSelectedDiagram));
                return;
            }

            _host.EditSelectedDiagram();
        }

        public string GetInteractionMode()
        {
            return "ExplicitSelectionCommand";
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }
    }

    [ComVisible(true)]
    [Guid("F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D")]
    [ProgId(AddInProgId)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public const string AddInProgId = "Greensoft.DrawioWordAddIn";

        private const string OfficeAddinsRegistryRoot = "Software\\Microsoft\\Office\\Word\\Addins\\";
        private WordInterop.Application _application;
        private COMAddIn _comAddIn;
        private Control _automationDispatcher;
        private WordAddInAutomation _automation;
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
            _application = application as WordInterop.Application;
            if (_application == null)
            {
                return;
            }

            _comAddIn = addInInst as COMAddIn;
            _host = new AddInHost(_application);
            _ribbonController = new RibbonController(_host);
            _automationDispatcher = new Control();
            _automationDispatcher.CreateControl();
            _automation = new WordAddInAutomation(_host, _automationDispatcher);
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

            _ribbonController = null;
            _host = null;
            _application = null;
            if (_comAddIn != null)
            {
                _comAddIn.Object = null;
                _comAddIn = null;
            }
            _automation = null;
            if (_automationDispatcher != null)
            {
                _automationDispatcher.Dispose();
                _automationDispatcher = null;
            }
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

        public void EditSelectedDiagram()
        {
            OnEditSelectedDiagram(null);
        }

        public void OnRefreshSelectedDiagram(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnRefreshSelectedDiagram(control);
            }
        }

        public void OnBindSelectedPicture(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnBindSelectedPicture(control);
            }
        }

        public void OnClearSelectedPicture(IRibbonControl control)
        {
            if (_ribbonController != null)
            {
                _ribbonController.OnClearSelectedPicture(control);
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

        public bool GetManagedPictureEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetManagedPictureEnabled(control);
        }

        public bool GetSinglePictureEnabled(IRibbonControl control)
        {
            if (_ribbonController == null)
            {
                return false;
            }

            return _ribbonController.GetSinglePictureEnabled(control);
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
                    throw new InvalidOperationException("Unable to create Word Add-in registry key.");
                }

                key.SetValue("FriendlyName", "DrawioWord");
                key.SetValue("Description", "Edit Draw.io diagrams inside Word.");
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
