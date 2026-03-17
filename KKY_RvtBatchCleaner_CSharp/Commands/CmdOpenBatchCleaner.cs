using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KKY_Tool_Revit.UI;

namespace KKY_Tool_Revit.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CmdOpenBatchCleaner : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var form = new BatchCleanerForm(commandData.Application);
            form.ShowDialog();
            return Result.Succeeded;
        }
    }
}
