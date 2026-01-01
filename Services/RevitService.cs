using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.Services
{
    public class RevitService
    {
        private Document _doc;

        public RevitService(Document doc)
        {
            _doc = doc;
        }

        public List<SheetItem> GetSheets()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate)
                .Select(s => new SheetItem(s))
                .OrderBy(s => s.SheetNumber)
                .ToList();
        }
    }
}
