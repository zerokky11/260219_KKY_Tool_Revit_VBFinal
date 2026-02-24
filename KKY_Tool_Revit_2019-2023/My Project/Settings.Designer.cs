using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace KKY_Tool_Revit.My
{
    [CompilerGenerated]
    [GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "11.0.0.0")]
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal sealed partial class MySettings : ApplicationSettingsBase
    {
        private static MySettings defaultInstance =
            (MySettings)Synchronized(new MySettings());

        public static MySettings Default
        {
            get { return defaultInstance; }
        }
    }

    [CompilerGenerated]
    [DebuggerNonUserCode]
    internal static class MySettingsProperty
    {
        [System.ComponentModel.Design.HelpKeyword("My.Settings")]
        internal static MySettings Settings
        {
            get { return MySettings.Default; }
        }
    }
}
