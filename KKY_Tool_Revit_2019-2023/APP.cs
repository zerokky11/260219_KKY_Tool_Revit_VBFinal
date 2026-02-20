using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KKY_Tool_Revit.UI.Hub;

namespace KKY_Tool_Revit.KKY_Tool_Revit
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            const string TAB_NAME = "KKY Tools";
            const string PANEL_NAME = "Hub";

            try { a.CreateRibbonTab(TAB_NAME); } catch { }

            var ribbonPanel = a.GetRibbonPanels(TAB_NAME).FirstOrDefault(p => p.Name == PANEL_NAME);
            if (ribbonPanel == null)
            {
                ribbonPanel = a.CreateRibbonPanel(TAB_NAME, PANEL_NAME);
            }

            var asmPath = Assembly.GetExecutingAssembly().Location;
            var cmdFullName = typeof(DuplicateExport).FullName;
            var pbd = new PushButtonData("KKY_Hub_Button", "KKY Hub", asmPath, cmdFullName);

            var btn = ribbonPanel.AddItem(pbd) as PushButton;
            if (btn != null)
            {
                btn.ToolTip = "KKY Tool 허브 열기";

                var smallImg = LoadHubIcon(
                    "KKY_Tool_Revit.Resources.Icons.KKY_Tool_16.png",
                    "KKY_Tool_Revit.Resources.Icons.KKY_Hub_16.png",
                    "KKY_Tool_Revit.Resources.icons.hub_16.png");

                var largeImg = LoadHubIcon(
                    "KKY_Tool_Revit.Resources.Icons.KKY_Tool_32.png",
                    "KKY_Tool_Revit.Resources.Icons.KKY_Hub_32.png",
                    "KKY_Tool_Revit.Resources.icons.hub_32.png");

                btn.Image = smallImg;
                btn.LargeImage = largeImg;
            }

            a.ViewActivated += OnViewActivated;
            a.ControlledApplication.DocumentOpened += OnDocumentListChanged;
            a.ControlledApplication.DocumentClosed += OnDocumentListChanged;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            try
            {
                a.ViewActivated -= OnViewActivated;
                a.ControlledApplication.DocumentOpened -= OnDocumentListChanged;
                a.ControlledApplication.DocumentClosed -= OnDocumentListChanged;
            }
            catch
            {
                // ignore
            }

            return Result.Succeeded;
        }

        private void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {
            try
            {
                HubHostWindow.NotifyActiveDocumentChanged(e.Document);
            }
            catch
            {
                // ignore
            }
        }

        private void OnDocumentListChanged(object sender, EventArgs e)
        {
            try
            {
                HubHostWindow.NotifyDocumentListChanged();
            }
            catch
            {
                // ignore
            }
        }

        private ImageSource LoadHubIcon(params string[] resourceNames)
        {
            foreach (var name in resourceNames)
            {
                var img = LoadPngImageSource(name);
                if (img != null) return img;
            }
            return null;
        }

        private ImageSource LoadPngImageSource(string resName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resolvedName = ResolveResourceName(asm, resName);
            if (string.IsNullOrEmpty(resolvedName))
            {
                Debug.WriteLine($"[KKY_Tool_Revit] 아이콘 리소스 로드 실패: {resName}");
                return null;
            }

            using (var s = asm.GetManifestResourceStream(resolvedName))
            {
                if (s == null)
                {
                    Debug.WriteLine($"[KKY_Tool_Revit] 아이콘 리소스 로드 실패: {resName}");
                    return null;
                }

                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = s;
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
        }

        private string ResolveResourceName(Assembly asm, string desired)
        {
            var names = asm.GetManifestResourceNames();

            var exact = names.FirstOrDefault(n => string.Equals(n, desired, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(exact)) return exact;

            var fileName = Path.GetFileName(desired);
            var byFile = names.FirstOrDefault(n => n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(byFile)) return byFile;

            return null;
        }
    }
}
