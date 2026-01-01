using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using ProSchedules.Commands;

namespace ProSchedules
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private const string RibbonTabName = "RK Tools";
        private const string ToolsPanelName = "Tools";
        private const string DuplicateSheetsId = "DuplicateSheetsCommand";

        private RibbonPanel ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = RibbonTabName;
            string toolsPanelName = ToolsPanelName;
            string duplicateSheetsId = DuplicateSheetsId;

            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
            }

            ribbonPanel = application.CreateOrSelectPanel(tabName, toolsPanelName);

            var duplicateSheetsData = new PushButtonData(
                duplicateSheetsId, 
                "Pro\nSchedules",
                typeof(DuplicateSheetsCommand).Assembly.Location, 
                typeof(DuplicateSheetsCommand).FullName
            )
            {
                ToolTip = "Manage sheet duplication and batch renaming.",
                LongDescription = "Duplicate sheets in bulk, rename with find/replace, prefixes/suffixes, and preview changes before applying.",
            };

            var item = ribbonPanel.AddItem(duplicateSheetsData);

            if (item is PushButton dupBtn)
            {
                dupBtn.LargeImage = GetImageSource("Assets/ProSchedules.tiff");
                dupBtn.Image = GetImageSource("Assets/ProSchedules.tiff");
                dupBtn.SetContextualHelp("https://github.com/RaulKalev/ProSchedules");
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
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

