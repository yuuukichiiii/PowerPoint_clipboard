using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Core = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ShapePalette.UI;

namespace ShapePalette
{
    /// <summary>
    /// アドイン本体。PowerPoint への接続、リボン、作業ウィンドウ（カスタムタスクペイン）を担う。
    /// </summary>
    // AutoDispatch: リボンの onAction 等のコールバック（OnOpenPalette）を
    // クラスの IDispatch 経由で名前解決できるようにする。None だと Office が呼べない。
    [ComVisible(true)]
    [Guid("352F4B34-C24C-42FB-AF36-FE303D48BB57")]
    [ProgId("ShapePalette.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Core.IRibbonExtensibility, Core.ICustomTaskPaneConsumer
    {
        private PowerPoint.Application _app;
        private Core.ICTPFactory _ctpFactory;
        private Core.CustomTaskPane _taskPane;

        public Connect() { Log.Write("Connect ctor"); }

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try { _app = Application as PowerPoint.Application; Log.Write("OnConnection app=" + (_app != null)); }
            catch (Exception ex) { Log.Write("OnConnection EX: " + ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_taskPane != null) { Marshal.ReleaseComObject(_taskPane); _taskPane = null; }
            }
            catch { }
            _ctpFactory = null;
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                var xml = ReadEmbedded("ShapePalette.Ribbon.xml");
                Log.Write("GetCustomUI -> " + xml.Length + " chars");
                return xml;
            }
            catch (Exception ex)
            {
                Log.Write("GetCustomUI EX: " + ex);
                return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" />";
            }
        }

        // リボンボタンの onAction
        public void OnOpenPalette(Core.IRibbonControl control)
        {
            try { Log.Write("OnOpenPalette"); TogglePalette(); }
            catch (Exception ex) { Log.Write("OnOpenPalette EX: " + ex); }
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Core.ICTPFactory CTPFactoryInst)
        {
            try { _ctpFactory = CTPFactoryInst; Log.Write("CTPFactoryAvailable factory=" + (CTPFactoryInst != null)); }
            catch (Exception ex) { Log.Write("CTPFactoryAvailable EX: " + ex); }
        }

        #endregion

        private void TogglePalette()
        {
            if (_ctpFactory == null) { Log.Write("TogglePalette: factory null"); return; }

            if (_taskPane == null)
            {
                Log.Write("CreateCTP...");
                _taskPane = _ctpFactory.CreateCTP("ShapePalette.PaletteHost", "Shape Palette", Type.Missing);
                _taskPane.DockPosition = Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 360;

                var host = _taskPane.ContentControl as PaletteHost;
                Log.Write("CTP created host=" + (host != null));
                if (host != null) host.Initialize(_app);
            }

            _taskPane.Visible = !_taskPane.Visible;
            Log.Write("TogglePalette visible=" + _taskPane.Visible);
        }

        private static string ReadEmbedded(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    Log.Write("ReadEmbedded: resource not found: " + resourceName +
                              " (available: " + string.Join(",", asm.GetManifestResourceNames()) + ")");
                    return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" />";
                }
                using (var reader = new StreamReader(stream)) { return reader.ReadToEnd(); }
            }
        }

        #region COM 登録（PowerPoint アドインキー）

        private const string AddInKeyPath =
            @"Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect";

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AddInKeyPath))
            {
                key.SetValue("FriendlyName", "Shape Palette");
                key.SetValue("Description", "図形をタブ別に並べ、クリックでスライドに挿入するパレット");
                key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try { Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree(AddInKeyPath, false); }
            catch { }
        }

        #endregion
    }
}
