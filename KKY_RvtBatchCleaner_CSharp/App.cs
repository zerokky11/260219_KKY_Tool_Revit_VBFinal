using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.Models;
using KKY_Tool_Revit.Services;

namespace KKY_Tool_Revit
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        private const string RibbonTabName = "KKY Tool";
        private const string RibbonPanelName = "RVT Cleaner";

        private static readonly object SharedStateLock = new object();
        private static BatchPrepareSession _sharedPreparedSession;
        private static BatchCleanSettings _sharedLastSettings;
        private static readonly List<string> _sharedLogLines = new List<string>();

        public static BatchPrepareSession SharedPreparedSession
        {
            get { lock (SharedStateLock) { return _sharedPreparedSession; } }
            set { lock (SharedStateLock) { _sharedPreparedSession = value; } }
        }

        public static BatchCleanSettings SharedLastSettings
        {
            get { lock (SharedStateLock) { return _sharedLastSettings != null ? _sharedLastSettings.Clone() : null; } }
            set { lock (SharedStateLock) { _sharedLastSettings = value != null ? value.Clone() : null; } }
        }

        public static IList<string> GetSharedLogLinesSnapshot()
        {
            lock (SharedStateLock)
            {
                return new List<string>(_sharedLogLines);
            }
        }

        public static void AppendSharedLog(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return;
            lock (SharedStateLock)
            {
                _sharedLogLines.Add(line);
                if (_sharedLogLines.Count > 4000)
                {
                    _sharedLogLines.RemoveRange(0, _sharedLogLines.Count - 4000);
                }
            }
        }

        public static void ClearSharedLog()
        {
            lock (SharedStateLock)
            {
                _sharedLogLines.Clear();
            }
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                EnsureRibbonTab(application, RibbonTabName);
                RibbonPanel panel = GetOrCreatePanel(application, RibbonTabName, RibbonPanelName);

                string assemblyPath = typeof(App).Assembly.Location;
                var buttonData = new PushButtonData(
                    "KKY_RvtBatchCleaner_Button",
                    "RVT\r\n정리",
                    assemblyPath,
                    "KKY_Tool_Revit.Commands.CmdOpenBatchCleaner");

                PushButton button = panel.AddItem(buttonData) as PushButton;
                if (button != null)
                {
                    button.ToolTip = "다중 RVT 파일을 Detach, 정리, 퍼지, 저장까지 단계별로 처리합니다.";
                    button.LongDescription = "Detach + workset discard, 링크 정리, 뷰 재구성, 필터/파라미터 적용 후 정리-퍼지-저장을 분리해서 진행합니다.";
                }

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try { PurgeUiBatchService.Stop(); } catch { }
            try { PurgeProgressWindowHost.CloseWindow(); } catch { }
            return Result.Succeeded;
        }

        private static void EnsureRibbonTab(UIControlledApplication application, string tabName)
        {
            try { application.CreateRibbonTab(tabName); } catch { }
        }

        private static RibbonPanel GetOrCreatePanel(UIControlledApplication application, string tabName, string panelName)
        {
            foreach (RibbonPanel panel in application.GetRibbonPanels(tabName))
            {
                if (string.Equals(panel.Name, panelName, StringComparison.OrdinalIgnoreCase)) return panel;
            }
            return application.CreateRibbonPanel(tabName, panelName);
        }
    }
}
