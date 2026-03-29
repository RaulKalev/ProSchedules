using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using System.Windows.Media.Imaging;
using ProSchedules.Commands;
using AdWindows = Autodesk.Windows;

namespace ProSchedules
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private RibbonPanel ribbonPanel;
        private PushButton _pushButton;

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

            // Create PushButton — icon set after based on current theme
            _pushButton = ribbonPanel.CreatePushButton<ProSheetsCommand>()
                .SetText("Pro\r\nSchedules")
                .SetToolTip("Manage sheet duplication and batch renaming.")
                .SetLongDescription("Pro Schedules allows you to duplicate sheets in bulk, rename with find/replace, prefixes/suffixes, and preview changes before applying.")
                .SetContextualHelp("https://github.com/RaulKalev/ProSchedules");

            UpdateRibbonIcon();

            try { AdWindows.ComponentManager.PropertyChanged += OnComponentManagerPropertyChanged; } catch { }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try { AdWindows.ComponentManager.PropertyChanged -= OnComponentManagerPropertyChanged; } catch { }
            ribbonPanel?.Remove();
            return Result.Succeeded;
        }

        private void OnComponentManagerPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == AdWindows.ComponentManager.CurrentThemePropertyName)
                UpdateRibbonIcon();
        }

        private void UpdateRibbonIcon()
        {
            if (_pushButton == null) return;
            bool isDark = AdWindows.ComponentManager.CurrentTheme?.Name == "Dark";
            string asset = isDark
                ? "pack://application:,,,/ProSchedules;component/Assets/Dark%20-%20ProSchedules.tiff"
                : "pack://application:,,,/ProSchedules;component/Assets/Light%20-%20ProSchedules.tiff";
            var image = new BitmapImage(new Uri(asset, UriKind.Absolute));
            _pushButton.LargeImage = image;
            _pushButton.Image = image;
        }
    }
}

