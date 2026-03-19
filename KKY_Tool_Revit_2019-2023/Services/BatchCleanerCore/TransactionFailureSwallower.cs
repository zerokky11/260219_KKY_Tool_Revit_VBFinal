using Autodesk.Revit.DB;

namespace KKY_Tool_Revit.Services
{
    public class TransactionFailureSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            foreach (var message in failuresAccessor.GetFailureMessages())
            {
                if (message.GetSeverity() == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(message);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}
