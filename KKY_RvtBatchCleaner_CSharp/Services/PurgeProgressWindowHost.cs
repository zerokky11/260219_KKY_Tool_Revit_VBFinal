using System;
using System.Threading;
using WinForms = System.Windows.Forms;
using KKY_Tool_Revit.UI;

namespace KKY_Tool_Revit.Services
{
    public static class PurgeProgressWindowHost
    {
        private static readonly object SyncRoot = new object();
        private static Thread _uiThread;
        private static PurgeProgressForm _form;

        public static void ShowWindow()
        {
            lock (SyncRoot)
            {
                if (_form != null && !_form.IsDisposed)
                {
                    try
                    {
                        _form.BeginInvoke(new Action(() =>
                        {
                            if (!_form.Visible) _form.Show();
                            _form.WindowState = WinForms.FormWindowState.Normal;
                            _form.BringToFront();
                        }));
                    }
                    catch
                    {
                    }
                    return;
                }

                _uiThread = new Thread(ThreadMain);
                _uiThread.IsBackground = true;
                _uiThread.Name = "KKY_PurgeProgressWindow";
                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.Start();
            }
        }

        public static void CloseWindow()
        {
            lock (SyncRoot)
            {
                if (_form == null || _form.IsDisposed)
                {
                    _form = null;
                    return;
                }

                try
                {
                    _form.BeginInvoke(new Action(() =>
                    {
                        if (!_form.IsDisposed)
                        {
                            _form.Close();
                        }
                    }));
                }
                catch
                {
                }
            }
        }

        private static void ThreadMain()
        {
            using (var form = new PurgeProgressForm())
            {
                lock (SyncRoot)
                {
                    _form = form;
                }

                form.FormClosed += (_, __) =>
                {
                    lock (SyncRoot)
                    {
                        _form = null;
                        _uiThread = null;
                    }
                };

                WinForms.Application.Run(form);
            }
        }
    }
}
