using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using PlaceViews.Models;
using PlaceViews.Services;
using System.Collections.Generic;

namespace PlaceViews.ExternalEvents
{
    public class ViewPlacementHandler : IExternalEventHandler
    {
        public List<ViewItem> Views { get; set; }
        public List<SheetItem> Sheets { get; set; }

        public event System.Action<string, string> OnPlacementFinished;

        public void Execute(UIApplication app)
        {
            if (Views == null || Sheets == null) return;
            
            Document doc = app.ActiveUIDocument.Document;
            var service = new RevitService(doc);

            // Execute placement
            string result = service.PlaceViewsOnSheets(Views, Sheets);
            
            OnPlacementFinished?.Invoke("Placement Results", result);
        }

        public string GetName()
        {
            return "Place Views Handler";
        }
    }
}
