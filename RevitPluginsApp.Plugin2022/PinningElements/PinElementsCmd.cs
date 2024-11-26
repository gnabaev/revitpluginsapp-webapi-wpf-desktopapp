using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitPluginsApp.Plugin.PinningElements
{
    [Transaction(TransactionMode.Manual)]
    public class PinElementsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;

            Document doc = uiDoc.Document;

            var window = new PinElementsWnd(doc);
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
