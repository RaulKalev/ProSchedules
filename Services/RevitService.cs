using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PlaceViews.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PlaceViews.Services
{
    public class RevitService
    {
        private Document _doc;

        public RevitService(Document doc)
        {
            _doc = doc;
        }

        public List<ViewItem> GetViews()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.ViewType != ViewType.DrawingSheet && v.ViewType != ViewType.Internal && v.ViewType != ViewType.ProjectBrowser) // Basic filtering
                .Where(v => v.CanBePrinted) // Usually good proxy for user views
                .Select(v => new ViewItem(v))
                .OrderBy(v => v.Name)
                .ToList();
        }

        public List<SheetItem> GetSheets()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate) // Should probably check IsPlaceholder too if relevant
                .Select(s => new SheetItem(s))
                .OrderBy(s => s.SheetNumber)
                .ToList();
        }

        public string PlaceViewsOnSheets(List<ViewItem> views, List<SheetItem> sheets)
        {
            using (Transaction t = new Transaction(_doc, "Place Views"))
            {
                t.Start();
                int successCount = 0;
                int skipCount = 0;
                var matchLog = new System.Text.StringBuilder();

                matchLog.AppendLine($"Attempting to match {sheets.Count} sheet(s) with {views.Count} view(s):\n");

                foreach (var sheetItem in sheets)
                {
                    // Find best matching view using fuzzy case-insensitive matching
                    var matchingView = FindBestMatch(sheetItem.Name, views);
                    
                    if (matchingView != null)
                    {
                        matchLog.AppendLine($"✓ Sheet: '{sheetItem.Name}' → View: '{matchingView.Name}'");
                        
                        try
                        {
                            ViewSheet sheet = _doc.GetElement(sheetItem.Id) as ViewSheet;
                            View view = _doc.GetElement(matchingView.Id) as View;

                            if (sheet != null && view != null)
                            {
                                // Check if view can be added to sheet
                                if (Viewport.CanAddViewToSheet(_doc, sheet.Id, view.Id))
                                {
                                    // Get sheet center point for viewport placement
                                    BoundingBoxUV outline = sheet.Outline;
                                    XYZ center = new XYZ(
                                        (outline.Min.U + outline.Max.U) / 2,
                                        (outline.Min.V + outline.Max.V) / 2,
                                        0);

                                    Viewport.Create(_doc, sheet.Id, view.Id, center);
                                    successCount++;
                                }
                                else
                                {
                                    matchLog.AppendLine($"  ⚠ Cannot place (already on sheet or incompatible)");
                                    skipCount++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            matchLog.AppendLine($"  ⚠ Error: {ex.Message}");
                            skipCount++;
                        }
                    }
                    else
                    {
                        matchLog.AppendLine($"✗ Sheet: '{sheetItem.Name}' → No matching view found");
                    }
                }
                
                t.Commit();
                
                // Show result message with details
                string message = $"Placement completed!\n\nSuccessfully placed: {successCount} view(s)";
                if (skipCount > 0)
                    message += $"\nSkipped: {skipCount} view(s)";
                message += $"\n\n{matchLog}";
                
                return message;
            }
        }

        private ViewItem FindBestMatch(string sheetName, List<ViewItem> views)
        {
            if (string.IsNullOrWhiteSpace(sheetName) || views == null || !views.Any())
                return null;

            // First try exact match (case-insensitive)
            var exactMatch = views.FirstOrDefault(v => 
                string.Equals(v.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            
            if (exactMatch != null)
                return exactMatch;

            // Extract key parts from sheet name (e.g., "Heli B01-c" -> ["Heli", "B01", "c"])
            var sheetParts = ExtractKeyParts(sheetName);
            
            if (sheetParts.Count == 0)
                return null;

            // Score each view and find the best match
            var scoredViews = views.Select(v => new
            {
                View = v,
                Score = CalculateMatchScore(sheetName, sheetParts, v.Name)
            })
            .Where(x => x.Score > 0)
            .OrderByDescending(x => x.Score)
            .ToList();

            return scoredViews.FirstOrDefault()?.View;
        }

        private int CalculateMatchScore(string sheetName, List<string> sheetParts, string viewName)
        {
            var viewNameLower = viewName.ToLowerInvariant();
            var sheetNameLower = sheetName.ToLowerInvariant();
            int score = 0;

            // Check if all sheet parts are in the view name (required for any match)
            if (!sheetParts.All(part => viewNameLower.Contains(part.ToLowerInvariant())))
                return 0; // No match

            // Base score for containing all parts
            score += 10;

            // Bonus points for matching the last part (suffix) - this is critical for "Heli B01-c" vs "Heli B01-d"
            if (sheetParts.Count > 0 && viewName.Length > 0)
            {
                var lastSheetPart = sheetParts.Last().ToLowerInvariant();
                
                // Check if view ends with the same suffix (highest priority)
                if (viewNameLower.EndsWith(lastSheetPart) || viewNameLower.EndsWith($"- {lastSheetPart}"))
                {
                    score += 100; // Very high bonus for exact suffix match
                }
                // Check if the last part appears near the end of the view name
                else if (viewNameLower.LastIndexOf(lastSheetPart) > viewNameLower.Length - 10)
                {
                    score += 50; // Medium bonus for suffix appearing near end
                }
            }

            // Bonus for each matching part
            score += sheetParts.Count(part => viewNameLower.Contains(part.ToLowerInvariant())) * 5;

            return score;
        }

        private List<string> ExtractKeyParts(string name)
        {
            // Split by common separators and filter out empty/short parts
            var separators = new[] { ' ', '-', '_', ':', '/', '\\', '|' };
            return name.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                       .Where(part => part.Length > 0) // Keep all parts including single letters
                       .Select(part => part.Trim())
                       .Where(part => !string.IsNullOrWhiteSpace(part))
                       .ToList();
        }
    }
}
