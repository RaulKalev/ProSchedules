using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PlaceViews.Models;
using PlaceViews.Services;
using PlaceViews.ExternalEvents;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace PlaceViews
{
    public partial class MainWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\PlaceViews\config.json";
        private const string WindowLeftKey = "MainWindow.Left";
        private const string WindowTopKey = "MainWindow.Top";
        private const string WindowWidthKey = "MainWindow.Width";
        private const string WindowHeightKey = "MainWindow.Height";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Revit state / UI state

        private UIDocument _uiDoc;
        private Document _doc;
        private View _currentView;

        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isDataLoaded;

        public ObservableCollection<ViewItem> Views { get; set; } = new ObservableCollection<ViewItem>();
        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<ViewItem> FilteredViews { get; set; } = new ObservableCollection<ViewItem>();
        public ObservableCollection<SheetItem> FilteredSheets { get; set; } = new ObservableCollection<SheetItem>();

        private readonly RevitService _revitService;
        private ViewPlacementHandler _requestHandler;
        private ExternalEvent _exEvent;

        #endregion

        #region Ctor / Init

        public MainWindow(UIDocument uiDoc, Document doc, View currentView)
        {
            InitializeComponent();

            _uiDoc = uiDoc;
            _doc = doc;
            _currentView = currentView;
            _revitService = new RevitService(_doc);

            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            DeferWindowShow();

            // Window infrastructure
            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;

            // Window-level mouse hooks for resizing
            MouseMove += Window_MouseMove;
            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            // External Event Setup
            _requestHandler = new ViewPlacementHandler();
            _exEvent = ExternalEvent.Create(_requestHandler);

            // Theme + DataContext
            _requestHandler.OnPlacementFinished += (title, msg) => 
            {
                // Ensure UI thread
                Dispatcher.Invoke(() => ShowPopup(title, msg, () => Close()));
            };
            
            LoadThemeState();
            LoadWindowState();
            // Theme resources are now inlined in MainWindow.xaml
            DataContext = this;

            LoadData();
        }

        private void LoadData()
        {
            try
            {
                var views = _revitService.GetViews()
                    .OrderBy(v => v.ViewType)
                    .ThenBy(v => v.Name);

                Views.Clear();
                FilteredViews.Clear();
                foreach (var v in views)
                {
                    Views.Add(v);
                    FilteredViews.Add(v);
                }

                var sheets = _revitService.GetSheets()
                    .OrderBy(s => s.Name); // Optional: sort sheets by name too for consistency

                Sheets.Clear();
                FilteredSheets.Clear();
                foreach (var s in sheets)
                {
                    Sheets.Add(s);
                    FilteredSheets.Add(s);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}");
            }
            finally
            {
                _isDataLoaded = true;
                TryShowWindow();
            }
        }

        #endregion

        #region Actions

        private void PlaceViews_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedViews = Views.Where(v => v.IsSelected).ToList();
                var selectedSheets = Sheets.Where(s => s.IsSelected).ToList();

                if (selectedViews.Count == 0 && selectedSheets.Count == 0)
                {
                    ShowPopup("No Selection", "Please select at least one View and one Sheet.");
                    return;
                }

                if (selectedViews.Count == 0)
                {
                    ShowPopup("No Views Selected", "Please select at least one View.");
                    return;
                }

                if (selectedSheets.Count == 0)
                {
                    ShowPopup("No Sheets Selected", "Please select at least one Sheet.");
                    return;
                }

                _requestHandler.Views = selectedViews;
                _requestHandler.Sheets = selectedSheets;
                _exEvent.Raise();
                
                // Do not Close() here; wait for result event
            }
            catch (Exception ex)
            {
                ShowPopup("Error", $"Error: {ex.Message}");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ViewsDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                var dataGrid = sender as DataGrid;
                if (dataGrid?.SelectedItems != null && dataGrid.SelectedItems.Count > 0)
                {
                    e.Handled = true; // Prevent default spacebar behavior first
                    
                    // Toggle all selected items
                    var selectedViews = dataGrid.SelectedItems.Cast<ViewItem>().ToList();
                    
                    // Determine new state based on first item
                    bool newState = !selectedViews.First().IsSelected;
                    
                    foreach (var view in selectedViews)
                    {
                        view.IsSelected = newState;
                    }
                    
                    // Force refresh
                    dataGrid.Items.Refresh();
                }
            }
        }

        private void SheetsDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                var dataGrid = sender as DataGrid;
                if (dataGrid?.SelectedItems != null && dataGrid.SelectedItems.Count > 0)
                {
                    e.Handled = true; // Prevent default spacebar behavior first
                    
                    // Toggle all selected items
                    var selectedSheets = dataGrid.SelectedItems.Cast<SheetItem>().ToList();
                    
                    // Determine new state based on first item
                    bool newState = !selectedSheets.First().IsSelected;
                    
                    foreach (var sheet in selectedSheets)
                    {
                        sheet.IsSelected = newState;
                    }
                    
                    // Force refresh
                    dataGrid.Items.Refresh();
                }
            }
        }

        private void DataGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Check if the click is on a checkbox
            var dep = e.OriginalSource as DependencyObject;
            
            while (dep != null && !(dep is DataGridCell) && !(dep is DataGridColumnHeader))
            {
                if (dep is CheckBox)
                {
                    // Clicking on checkbox - don't change selection
                    e.Handled = true;
                    
                    // Let the checkbox handle the click
                    var checkbox = dep as CheckBox;
                    checkbox.IsChecked = !checkbox.IsChecked;
                    return;
                }
                dep = VisualTreeHelper.GetParent(dep);
            }
        }

        private void ViewSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = (sender as System.Windows.Controls.TextBox)?.Text?.ToLowerInvariant() ?? "";
            
            var filtered = Views.Where(v => 
                string.IsNullOrEmpty(searchText) || 
                v.Name.ToLowerInvariant().Contains(searchText))
                .OrderBy(v => v.ViewType)
                .ThenBy(v => v.Name);

            FilteredViews.Clear();
            foreach (var view in filtered)
            {
                FilteredViews.Add(view);
            }
        }

        private void SheetSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = (sender as System.Windows.Controls.TextBox)?.Text?.ToLowerInvariant() ?? "";
            
            FilteredSheets.Clear();
            foreach (var sheet in Sheets.Where(s => 
                string.IsNullOrEmpty(searchText) || 
                s.Name.ToLowerInvariant().Contains(searchText) ||
                s.SheetNumber.ToLowerInvariant().Contains(searchText)))
            {
                FilteredSheets.Add(sheet);
            }
        }

        private void ClearViewSearch_Click(object sender, RoutedEventArgs e)
        {
            ViewSearchBox.Clear();
        }

        private void ClearSheetSearch_Click(object sender, RoutedEventArgs e)
        {
            SheetSearchBox.Clear();
        }

        private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            // Cancel edit to prevent checkbox column from triggering edit mode
            // This keeps the row selection intact when clicking checkboxes
            e.Cancel = true;
        }

        private void ViewCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is ViewItem clickedItem)
            {
                // Prevent the row selection from changing
                e.Handled = true;
                
                // Toggle the state manually
                bool newState = !(checkBox.IsChecked ?? false);
                checkBox.IsChecked = newState;

                // Sync the item property (binding might not fire yet due to Handled=true?) 
                // Actually binding is TwoWay so setting IsChecked updates the property.
                
                // If the clicked item is part of the selection, apply the change to all selected items
                if (ViewsDataGrid.SelectedItems.Contains(clickedItem))
                {
                    foreach (ViewItem item in ViewsDataGrid.SelectedItems)
                    {
                        if (item != clickedItem)
                        {
                            item.IsSelected = newState;
                        }
                    }
                }
            }
        }

        private void SheetCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is SheetItem clickedItem)
            {
                // Prevent the row selection from changing
                e.Handled = true;
                
                // Toggle the state manually
                bool newState = !(checkBox.IsChecked ?? false);
                checkBox.IsChecked = newState;

                // If the clicked item is part of the selection, apply the change to all selected items
                if (SheetsDataGrid.SelectedItems.Contains(clickedItem))
                {
                    foreach (SheetItem item in SheetsDataGrid.SelectedItems)
                    {
                        if (item != clickedItem)
                        {
                            item.IsSelected = newState;
                        }
                    }
                }
            }
        }

        private void SelectAllViews_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var view in FilteredViews)
            {
                view.IsSelected = true;
            }
        }

        private void SelectAllViews_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var view in FilteredViews)
            {
                view.IsSelected = false;
            }
        }

        private void SelectAllSheets_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in FilteredSheets)
            {
                sheet.IsSelected = true;
            }
        }

        private void SelectAllSheets_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in FilteredSheets)
            {
                sheet.IsSelected = false;
            }
        }

        #endregion

        #region Theme

        private ResourceDictionary _currentThemeDictionary;

        private void LoadTheme()
        {
            try
            {
                var themeUri = new Uri(_isDarkMode 
                    ? "pack://application:,,,/PlaceViews;component/UI/Themes/DarkTheme.xaml" 
                    : "pack://application:,,,/PlaceViews;component/UI/Themes/LightTheme.xaml", UriKind.Absolute);
                
                var newDict = new ResourceDictionary { Source = themeUri };
                
                if (_currentThemeDictionary != null)
                {
                    Resources.MergedDictionaries.Remove(_currentThemeDictionary);
                }
                
                Resources.MergedDictionaries.Add(newDict);
                _currentThemeDictionary = newDict;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading theme: {ex.Message}");
            }
        }

        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = ThemeToggleButton.IsChecked == true;
            LoadTheme();
            SaveThemeState();

            var icon = ThemeToggleButton?.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                       as MaterialDesignThemes.Wpf.PackIcon;
            if (icon != null)
            {
                icon.Kind = _isDarkMode
                    ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                    : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
            }
        }

        private void LoadThemeState()
        {
            try
            {
                var config = LoadConfig();
                if (TryGetBool(config, "IsDarkMode", out var isDark))
                {
                    _isDarkMode = isDark;
                }
            }
            catch (Exception)
            {
                // Ignore config load errors
            }

            // reflect UI
                if (ThemeToggleButton != null)
            {
                ThemeToggleButton.IsChecked = _isDarkMode;
                var icon = ThemeToggleButton.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                           as MaterialDesignThemes.Wpf.PackIcon;
                if (icon != null)
                {
                    icon.Kind = _isDarkMode
                        ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                        : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
                }
            }
            
            LoadTheme();
        }

        private void SaveThemeState()
        {
            try
            {
                var config = LoadConfig();
                config["IsDarkMode"] = _isDarkMode;
                SaveConfig(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Custom Popup

        
        private Action _onPopupClose;

        public void ShowPopup(string title, string message, Action onCloseAction = null)
        {
            PopupTitle.Text = title;
            PopupMessage.Text = message;
            _onPopupClose = onCloseAction;
            PopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ClosePopup()
        {
            PopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onPopupClose != null)
            {
                var action = _onPopupClose;
                _onPopupClose = null;
                action.Invoke();
            }
        }

        private void PopupClose_Click(object sender, RoutedEventArgs e)
        {
            ClosePopup();
        }

        private void PopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Optional: Close on background click?
            // Confirm with user preference. For now, we'll allow it as it's standard UX.
            // ClosePopup(); 
            // Actually, keep it modal-like (explicit close) unless requested otherwise.
        }

        #endregion

        #region Window chrome / resize handlers

        private void TitleBar_Loaded(object sender, RoutedEventArgs e)
        {
            // Optional initialization for TitleBar
        }

        private void LeftEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void RightEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void BottomEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNS;
        private void Edge_MouseLeave(object sender, MouseEventArgs e) => Cursor = Cursors.Arrow;
        private void BottomLeftCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNESW;
        private void BottomRightCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNWSE;

        private void Window_MouseMove(object sender, MouseEventArgs e) => _windowResizer.ResizeWindow(e);
        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) => _windowResizer.StopResizing();
        private void LeftEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Left);
        private void RightEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Right);
        private void BottomEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Bottom);
        private void BottomLeftCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomLeft);
        private void BottomRightCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomRight);

        #endregion

        #region Cleanup / disposal

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveWindowState();
        }

        #endregion

        #region Window Startup

        private void DeferWindowShow()
        {
            Opacity = 0;
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            TryShowWindow();
        }

        private void TryShowWindow()
        {
            if (!_isDataLoaded)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                Opacity = 1;
            }), DispatcherPriority.Render);
        }

        #endregion

        #region Window State

        private void LoadWindowState()
        {
            try
            {
                var config = LoadConfig();
                bool hasLeft = TryGetDouble(config, WindowLeftKey, out var left);
                bool hasTop = TryGetDouble(config, WindowTopKey, out var top);
                bool hasWidth = TryGetDouble(config, WindowWidthKey, out var width);
                bool hasHeight = TryGetDouble(config, WindowHeightKey, out var height);

                bool hasSize = hasWidth && hasHeight && width > 0 && height > 0;
                bool hasPos = hasLeft && hasTop && !double.IsNaN(left) && !double.IsNaN(top);

                if (!hasSize && !hasPos)
                {
                    return;
                }

                WindowStartupLocation = WindowStartupLocation.Manual;

                if (hasSize)
                {
                    Width = Math.Max(MinWidth, width);
                    Height = Math.Max(MinHeight, height);
                }

                if (hasPos)
                {
                    Left = left;
                    Top = top;
                }
            }
            catch (Exception)
            {
            }
        }

        private void SaveWindowState()
        {
            try
            {
                var config = LoadConfig();
                var bounds = WindowState == WindowState.Normal
                    ? new Rect(Left, Top, Width, Height)
                    : RestoreBounds;

                config[WindowLeftKey] = bounds.Left;
                config[WindowTopKey] = bounds.Top;
                config[WindowWidthKey] = bounds.Width;
                config[WindowHeightKey] = bounds.Height;

                SaveConfig(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save window state: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Config Helpers

        private Dictionary<string, object> LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (config != null)
                    {
                        return config;
                    }
                }
            }
            catch (Exception)
            {
            }

            return new Dictionary<string, object>();
        }

        private void SaveConfig(Dictionary<string, object> config)
        {
            var dir = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
        }

        private static bool TryGetBool(Dictionary<string, object> config, string key, out bool value)
        {
            value = false;
            if (!config.TryGetValue(key, out var raw) || raw == null)
            {
                return false;
            }

            if (raw is bool boolValue)
            {
                value = boolValue;
                return true;
            }

            if (raw is JToken token && token.Type == JTokenType.Boolean)
            {
                value = token.Value<bool>();
                return true;
            }

            if (raw is string text && bool.TryParse(text, out var parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        private static bool TryGetDouble(Dictionary<string, object> config, string key, out double value)
        {
            value = 0;
            if (!config.TryGetValue(key, out var raw) || raw == null)
            {
                return false;
            }

            switch (raw)
            {
                case double doubleValue:
                    value = doubleValue;
                    return true;
                case float floatValue:
                    value = floatValue;
                    return true;
                case decimal decimalValue:
                    value = (double)decimalValue;
                    return true;
                case long longValue:
                    value = longValue;
                    return true;
                case int intValue:
                    value = intValue;
                    return true;
                case JToken token when token.Type == JTokenType.Float || token.Type == JTokenType.Integer:
                    value = token.Value<double>();
                    return true;
                case string text when double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed):
                    value = parsed;
                    return true;
            }

            return false;
        }

        #endregion
    }
}
