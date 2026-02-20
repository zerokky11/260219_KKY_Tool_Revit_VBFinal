using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace KKY_Tool_Revit.UI.Hub
{
    public sealed class HubHostWindow : Window
    {
        private const string BaseTitle = "KKY Tool Hub";

        private static HubHostWindow _instance;
        private static readonly object Gate = new object();

        private readonly WebView2 _web = new WebView2();
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        private UIApplication _uiApp;
        private string _currentDocName = string.Empty;
        private string _currentDocPath = string.Empty;
        private bool _initStarted;
        private bool _isClosing;

        public WebView2 Web => _web;

        public bool IsClosing => _isClosing;

        public static void ShowSingleton(UIApplication uiApp)
        {
            if (uiApp == null)
            {
                return;
            }

            lock (Gate)
            {
                if (_instance != null && !_instance.IsClosing)
                {
                    _instance.AttachTo(uiApp);
                    UiBridgeExternalEvent.Initialize(_instance);
                    if (_instance.WindowState == WindowState.Minimized)
                    {
                        _instance.WindowState = WindowState.Normal;
                    }

                    _instance.Activate();
                    _instance.Focus();
                    return;
                }

                var wnd = new HubHostWindow(uiApp);
                UiBridgeExternalEvent.Initialize(wnd);
                _instance = wnd;
                wnd.Show();
            }
        }

        public static void NotifyActiveDocumentChanged(Document doc)
        {
            var inst = _instance;
            if (inst == null || inst.IsClosing)
            {
                return;
            }

            inst.UpdateActiveDocument(doc);
        }

        public static void NotifyDocumentListChanged()
        {
            var inst = _instance;
            if (inst == null || inst.IsClosing)
            {
                return;
            }

            inst.BroadcastDocumentList();
        }

        public HubHostWindow(UIApplication uiApp)
        {
            _uiApp = uiApp;
            Title = BaseTitle;
            var workArea = SystemParameters.WorkArea;
            const double desiredWidth = 1400;
            const double desiredHeight = 900;
            Width = Math.Min(desiredWidth, workArea.Width * 0.93);
            Height = Math.Min(desiredHeight, workArea.Height * 0.93);
            MinWidth = Math.Min(1100, Width);
            MinHeight = Math.Min(720, Height);
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Content = _web;

            Loaded += OnLoaded;
            Closing += OnWindowClosing;
            Closed += OnWindowClosed;

            UpdateActiveDocument(GetActiveDocument());
        }

        public void AttachTo(UIApplication uiApp)
        {
            _uiApp = uiApp;
            UpdateActiveDocument(GetActiveDocument());
            BroadcastDocumentList();
        }

        private Document GetActiveDocument()
        {
            try
            {
                if (_uiApp == null)
                {
                    return null;
                }

                var uidoc = _uiApp.ActiveUIDocument;
                if (uidoc == null)
                {
                    return null;
                }

                return uidoc.Document;
            }
            catch
            {
                return null;
            }
        }

        private void UpdateActiveDocument(Document doc)
        {
            var name = string.Empty;
            var path = string.Empty;

            if (doc != null)
            {
                try
                {
                    name = doc.Title;
                }
                catch
                {
                    // ignore
                }

                try
                {
                    path = doc.PathName;
                }
                catch
                {
                    // ignore
                }
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                path = name;
            }

            _currentDocName = name;
            _currentDocPath = path;

            UpdateWindowTitle();
            SendActiveDocument();
        }

        private void UpdateWindowTitle()
        {
            Title = string.IsNullOrWhiteSpace(_currentDocName)
                ? BaseTitle
                : $"{BaseTitle} - {_currentDocName}";
        }

        private string ResolveUiFolder()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var baseDir = Path.GetDirectoryName(asm.Location);
                var ui = Path.Combine(baseDir, "Resources", "HubUI");
                if (Directory.Exists(ui))
                {
                    return Path.GetFullPath(ui);
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_initStarted)
            {
                return;
            }

            _initStarted = true;
            try
            {
                var userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KKY_Tool_Revit",
                    "WebView2UserData");
                Directory.CreateDirectory(userData);

                var env = await CoreWebView2Environment.CreateAsync(null, userData, null);
                await _web.EnsureCoreWebView2Async(env);
                var core = _web.CoreWebView2;

                core.Settings.AreDefaultContextMenusEnabled = false;
                core.Settings.IsStatusBarEnabled = false;
#if DEBUG
                core.Settings.AreDevToolsEnabled = true;
#else
                core.Settings.AreDevToolsEnabled = false;
#endif

                var uiFolder = ResolveUiFolder();
                if (string.IsNullOrEmpty(uiFolder))
                {
                    throw new DirectoryNotFoundException(@"Resources\HubUI 폴더를 찾을 수 없습니다.");
                }

                core.SetVirtualHostNameToFolderMapping(
                    "hub.local",
                    uiFolder,
                    CoreWebView2HostResourceAccessKind.Allow);

                core.WebMessageReceived += OnWebMessage;
                _web.Source = new Uri("https://hub.local/index.html");

                SendToWeb("host:topmost", new { on = Topmost });
                SendActiveDocument();
                BroadcastDocumentList();
            }
            catch (Exception ex)
            {
                var hr = Runtime.InteropServices.Marshal.GetHRForException(ex);
                MessageBox.Show(
                    $"WebView 초기화 실패 (0x{hr:X8}) : {ex.Message}",
                    "KKY Tool",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OnWindowClosing(object sender, CancelEventArgs e)
        {
            _isClosing = true;
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
            lock (Gate)
            {
                if (ReferenceEquals(_instance, this))
                {
                    _instance = null;
                }
            }
        }

        private void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var root = _serializer.Deserialize<Dictionary<string, object>>(e.WebMessageAsJson);

                string name = null;
                if (root != null)
                {
                    if (root.TryGetValue("ev", out var evObj) && evObj != null)
                    {
                        name = Convert.ToString(evObj);
                    }
                    else if (root.TryGetValue("name", out var nameObj) && nameObj != null)
                    {
                        name = Convert.ToString(nameObj);
                    }
                }

                if (string.IsNullOrEmpty(name))
                {
                    return;
                }

                object payload = null;
                if (root != null && root.TryGetValue("payload", out var payloadObj))
                {
                    payload = payloadObj;
                }

                switch (name)
                {
                    case "ui:ping":
                        SendToWeb("host:pong", new { t = DateTime.Now.Ticks });
                        break;

                    case "ui:toggle-topmost":
                        Topmost = !Topmost;
                        SendToWeb("host:topmost", new { on = Topmost });
                        break;

                    case "ui:query-topmost":
                        SendToWeb("host:topmost", new { on = Topmost });
                        break;

                    default:
                        UiBridgeExternalEvent.Raise(name, payload);
                        break;
                }
            }
            catch (Exception ex)
            {
                SendToWeb("host:error", new { ex.Message });
            }
        }

        private void BroadcastDocumentList()
        {
            var docs = new List<object>();
            try
            {
                if (_uiApp != null && _uiApp.Application != null)
                {
                    foreach (Document d in _uiApp.Application.Documents)
                    {
                        try
                        {
                            var name = d.Title;
                            var path = d.PathName;
                            if (string.IsNullOrWhiteSpace(path))
                            {
                                path = name;
                            }

                            docs.Add(new { name, path });
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }

            SendToWeb("host:doc-list", docs);
        }

        private void SendActiveDocument()
        {
            SendToWeb("host:doc-changed", new { name = _currentDocName, path = _currentDocPath });
        }

        public void SendToWeb(string ev, object payload)
        {
            var core = _web.CoreWebView2;
            if (core == null)
            {
                return;
            }

            var msg = new Dictionary<string, object>
            {
                ["ev"] = ev,
                ["name"] = ev,
                ["payload"] = payload
            };
            var json = _serializer.Serialize(msg);
            core.PostWebMessageAsJson(json);
        }
    }
}
