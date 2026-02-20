using System;
using System.Collections.Generic;

namespace KKY_Tool_Revit.UI.Hub
{
    public static class ExcelProgressReporter
    {
        private sealed class ProgressState
        {
            public DateTime LastSent { get; set; } = DateTime.MinValue;
            public int LastRow { get; set; }
        }

        private static readonly Dictionary<string, ProgressState> States = new Dictionary<string, ProgressState>(StringComparer.OrdinalIgnoreCase);
        private static readonly object Gate = new object();

        public static void Reset(string channel)
        {
            if (string.IsNullOrWhiteSpace(channel))
            {
                return;
            }

            lock (Gate)
            {
                States[channel] = new ProgressState();
            }
        }

        public static void Report(
            string channel,
            string phase,
            string message,
            int current,
            int total,
            double? percentOverride = null,
            bool force = false)
        {
            if (string.IsNullOrWhiteSpace(channel))
            {
                return;
            }

            var shouldSend = force;
            var now = DateTime.UtcNow;

            lock (Gate)
            {
                if (!States.TryGetValue(channel, out var st))
                {
                    st = new ProgressState();
                    States[channel] = st;
                }

                var deltaMs = (now - st.LastSent).TotalMilliseconds;
                var deltaRows = Math.Abs(current - st.LastRow);

                if (force || deltaMs >= 200.0 || deltaRows >= 200)
                {
                    st.LastSent = now;
                    st.LastRow = current;
                    shouldSend = true;
                }
            }

            if (!shouldSend)
            {
                return;
            }

            var phaseProgress = ComputePhaseProgress(phase, current, total);
            UiBridgeExternalEvent.SendToWeb(channel, new
            {
                phase,
                message,
                current,
                total,
                phaseProgress,
                percent = percentOverride
            });
        }

        private static double ComputePhaseProgress(string phase, int current, int total)
        {
            var norm = (phase ?? string.Empty).Trim().ToUpperInvariant();
            if (norm == "EXCEL_WRITE" || norm == "AUTOFIT")
            {
                if (total <= 0) return 0.0;
                return Math.Min(1.0, Math.Max(0.0, (double)current / total));
            }

            if (norm == "DONE") return 1.0;
            if (norm == "EXCEL_SAVE") return 1.0;
            return 0.0;
        }
    }
}
