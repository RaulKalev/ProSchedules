using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using PlaceViews.Commands;

namespace PlaceViews
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private const string RibbonTabName = "RK Tools";
        private const string ToolsPanelName = "Tools";
        private const string ViewManagerId = "ViewManagerCommand";
        private const string PlaceViewsId = "PlaceViewsCommand";
        private const string DuplicateSheetsId = "DuplicateSheetsCommand";

        private RibbonPanel ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            // Define the custom tab name
            string tabName = RibbonTabName;
            string toolsPanelName = ToolsPanelName;
            string viewManagerId = ViewManagerId;
            string placeViewsId = PlaceViewsId;
            string duplicateSheetsId = DuplicateSheetsId;

            // Try to create the custom tab (avoid exception if it already exists)
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // Tab already exists; continue without throwing an error
            }

            // Create Ribbon Panel on the custom tab
            ribbonPanel = application.CreateOrSelectPanel(tabName, toolsPanelName);

            // Create PushButton Data
            var viewManagerData = new PushButtonData(
                viewManagerId,
                "View\nManager",
                typeof(ViewManagerCommand).Assembly.Location,
                typeof(ViewManagerCommand).FullName
            )
            {
                ToolTip = "Manage project views.",
                LongDescription = "Open View Manager to organize, filter, and review views.",
            };

            var placeViewsData = new PushButtonData(
                placeViewsId, 
                "Place\nViews",
                typeof(MainCommand).Assembly.Location, 
                typeof(MainCommand).FullName
            )
            {
                ToolTip = "Place views on sheets.",
                LongDescription = "Batch place selected views onto sheets, review pending changes, and apply updates in one step.",
            };

            var duplicateSheetsData = new PushButtonData(
                duplicateSheetsId, 
                "Sheet\nManager",
                typeof(DuplicateSheetsCommand).Assembly.Location, 
                typeof(DuplicateSheetsCommand).FullName
            )
            {
                ToolTip = "Manage sheet duplication and batch renaming.",
                LongDescription = "Duplicate sheets in bulk, rename with find/replace, prefixes/suffixes, and preview changes before applying.",
            };

            // Add stacked items
            var items = ribbonPanel.AddStackedItems(viewManagerData, placeViewsData, duplicateSheetsData);

            // Assign Images (using the embedded resource via Pack URI or similar if accessible, 
            // but standard ImageSource helper is easiest)
            if (items.Count == 3)
            {
                if (items[0] is PushButton viewBtn)
                {
                   viewBtn.Image = GetImageSource("Assets/PlaceViews.tiff"); // Small image
                }
                if (items[1] is PushButton placeBtn)
                {
                   placeBtn.Image = GetImageSource("Assets/PlaceViews.tiff"); // Small image
                   placeBtn.SetContextualHelp("https://raulkalev.github.io/rktools/");
                }
                if (items[2] is PushButton dupBtn)
                {
                   dupBtn.Image = GetImageSource("Assets/PlaceViews.tiff"); // Reusing same icon for now
                   dupBtn.SetContextualHelp("https://raulkalev.github.io/rktools/");
                }
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Trigger the update check
            ribbonPanel?.Remove();
            return Result.Succeeded;
        }

        private System.Windows.Media.ImageSource GetImageSource(string resourcePath)
        {
            try
            {
                // Assuming resourcePath is relative to project root in embedded resources or strictly in output if Pack URI works
                // The Pack URI needs the assembly name.
                var assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                var uri = new Uri($"pack://application:,,,/{assemblyName};component/{resourcePath}", UriKind.Absolute);
                return new System.Windows.Media.Imaging.BitmapImage(uri);
            }
            catch
            {
                return null;
            }
        }
    }
}

