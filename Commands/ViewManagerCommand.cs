using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PlaceViews.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ViewManagerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("View Manager", "View Manager is not available yet.");
            return Result.Succeeded;
        }
    }
}
