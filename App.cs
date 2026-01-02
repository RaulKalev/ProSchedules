using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using ProSchedules.Commands;

namespace ProSchedules
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private RibbonPanel ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            // Define the custom tab name
            string tabName = "RK Tools";

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
            ribbonPanel = application.CreateOrSelectPanel(tabName, "Tools");

            // Create PushButton with embedded resource
            var duplicateSheetsButton = ribbonPanel.CreatePushButton<ProSheetsCommand>()
                .SetLargeImage("pack://application:,,,/ProSchedules;component/Assets/ProSchedules.tiff")
                .SetText("Pro\r\nSchedules")
                .SetToolTip("Manage sheet duplication and batch renaming.")
                .SetLongDescription("Duplicate sheets in bulk, rename with find/replace, prefixes/suffixes, and preview changes before applying.")
                .SetContextualHelp("https://github.com/RaulKalev/ProSchedules");

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Trigger the update check
            ribbonPanel?.Remove();
            return Result.Succeeded;
        }

    }
}

