namespace KKY_Tool_Revit.Models
{
    public class PurgeBatchProgressSnapshot
    {
        public bool IsRunning { get; set; }
        public bool IsCompleted { get; set; }
        public bool IsFaulted { get; set; }
        public int CurrentFileIndex { get; set; }
        public int TotalFiles { get; set; }
        public int CurrentIteration { get; set; }
        public int TotalIterations { get; set; }
        public string CurrentFileName { get; set; }
        public string StateName { get; set; }
        public string Message { get; set; }
    }
}
