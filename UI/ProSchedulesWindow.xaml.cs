using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using ProSchedules.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Data;
using ProSchedules.ExternalEvents;

namespace ProSchedules.UI
{
    public class SortItem : System.ComponentModel.INotifyPropertyChanged
    {
        private string _selectedColumn;
        public string SelectedColumn
        {
            get => _selectedColumn;
            set { _selectedColumn = value; OnPropertyChanged(nameof(SelectedColumn)); }
        }

        private bool _isAscending = true;
        public bool IsAscending
        {
            get => _isAscending;
            set { _isAscending = value; OnPropertyChanged(nameof(IsAscending)); }
        }
        
        // Visual placeholders matching screenshot
        public bool ShowHeader { get; set; }

        private bool _showFooter;
        public bool ShowFooter
        {
            get => _showFooter;
            set { _showFooter = value; OnPropertyChanged(nameof(ShowFooter)); }
        }

        private string _footerOption = "Title, count, and totals";
        public string FooterOption
        {
            get => _footerOption;
            set { _footerOption = value; OnPropertyChanged(nameof(FooterOption)); }
        }

        public bool ShowBlankLine { get; set; }

        public SortItem Clone()
        {
            return new SortItem
            {
                SelectedColumn = this.SelectedColumn,
                IsAscending = this.IsAscending,
                ShowHeader = this.ShowHeader,
                ShowFooter = this.ShowFooter,
                FooterOption = this.FooterOption,
                ShowBlankLine = this.ShowBlankLine
            };
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
    }

    public class InverseBooleanConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool?) && targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return !(bool)value;
        }
    }

    /// <summary>
    /// Represents an option in the Rename Parameter dropdown.
    /// </summary>
    public class RenameParameterOption
    {
        public string Name { get; set; }
        public bool IsTypeParameter { get; set; }
        public bool IsBuiltInParameter { get; set; } // True for Sheet Number/Sheet Name
    }

    public partial class ProSchedulesWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\ProSchedules\config.json";
        private const string WindowLeftKey = "ProSchedulesWindow.Left";
        private const string WindowTopKey = "ProSchedulesWindow.Top";
        private const string WindowWidthKey = "ProSchedulesWindow.Width";
        private const string WindowHeightKey = "ProSchedulesWindow.Height";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Revit state / UI state

        public static readonly DependencyProperty IsCopyModeProperty =
            DependencyProperty.Register("IsCopyMode", typeof(bool), typeof(ProSchedulesWindow), 
            new PropertyMetadata(false, (d, e) => ((ProSchedulesWindow)d).UpdateSelectionAdorner()));

        public bool IsCopyMode
        {
            get { return (bool)GetValue(IsCopyModeProperty); }
            set { SetValue(IsCopyModeProperty, value); }
        }

        private List<SheetItem> _allSheets;
        private Action _onPopupClose;
        private Action _onConfirmAction;
        private Action _onCancelAction;
        private ExternalEvent _externalEvent;
        private ExternalEvents.SheetDuplicationHandler _handler;
        private ExternalEvent _editExternalEvent;
        private ExternalEvents.SheetEditHandler _editHandler;


        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isDataLoaded;
        private UIApplication _uiApplication;
        private RevitService _revitService;
        private ExternalEvent _parameterRenameExternalEvent;
        private ExternalEvents.ParameterRenameHandler _parameterRenameHandler;
        private ExternalEvent _scheduleFieldsExternalEvent;
        private ExternalEvents.ScheduleFieldsHandler _scheduleFieldsHandler;
        private ExternalEvent _parameterLoadExternalEvent;
        private ExternalEvents.ParameterDataLoadHandler _parameterLoadHandler;
        private ExternalEvent _highlightInModelExternalEvent;
        private ExternalEvents.HighlightInModelHandler _highlightInModelHandler;

        private ExternalEvent _parameterValueUpdateExternalEvent;
        private ExternalEvents.ParameterValueUpdateHandler _parameterValueUpdateHandler;
        private ExternalEvents.SaveUserSettingsHandler _saveSettingsHandler;
        private ExternalEvent _saveSettingsExternalEvent;

        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<SheetItem> FilteredSheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<RenamePreviewItem> RenamePreviewItems { get; set; } = new ObservableCollection<RenamePreviewItem>();
        public ObservableCollection<SortItem> SortCriteria { get; set; } = new ObservableCollection<SortItem>();
        public ObservableCollection<string> AvailableSortColumns { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<Models.FilterItem> FilterCriteria { get; set; } = new ObservableCollection<Models.FilterItem>();
        public ObservableCollection<string> AvailableFilterColumns { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<ScheduleRenameItem> ScheduleRenamePreviewItems { get; set; } = new ObservableCollection<ScheduleRenameItem>();
        public ObservableCollection<RenameParameterOption> RenameParameterOptions { get; set; } = new ObservableCollection<RenameParameterOption>();

        private ProSchedules.Models.ScheduleData _currentScheduleData;
        private Dictionary<ElementId, bool> _scheduleItemizeSettings = new Dictionary<ElementId, bool>();
        private System.Data.DataTable _rawScheduleData;
        private Dictionary<ElementId, ObservableCollection<SortItem>> _scheduleSortSettings = new Dictionary<ElementId, ObservableCollection<SortItem>>();
        private Dictionary<ElementId, ObservableCollection<Models.FilterItem>> _scheduleFilterSettings = new Dictionary<ElementId, ObservableCollection<Models.FilterItem>>();
        private string _lastSelectedScheduleName;

        // Manual update mode
        private bool _isManualMode = false;
        private List<PendingParameterChange> _pendingParameterChanges = new List<PendingParameterChange>();
        private HashSet<string> _highlightedCells = new HashSet<string>(); // "rowIndex:columnName"

        #endregion

        #region Manual Mode Data

        private class PendingParameterChange
        {
            public string ElementIdStr { get; set; }
            public ElementId ParameterId { get; set; }
            public string NewValue { get; set; }
            public string OldValue { get; set; }
            public string ColumnName { get; set; }
            public int RowIndex { get; set; }
        }

        #endregion

        #region Ctor / Init

        public ProSchedulesWindow(UIApplication app)
        {
            _uiApplication = app;
            InitializeComponent();
            DataContext = this;

            // Create duplication handler
            _handler = new ExternalEvents.SheetDuplicationHandler();
            _handler.OnDuplicationFinished += OnDuplicationFinished;
            _externalEvent = ExternalEvent.Create(_handler);

            // Create edit handler
            _editHandler = new ExternalEvents.SheetEditHandler();
            _editHandler.OnEditFinished += OnEditFinished;
            _editExternalEvent = ExternalEvent.Create(_editHandler);

            _parameterRenameHandler = new ExternalEvents.ParameterRenameHandler();
            _parameterRenameHandler.OnRenameFinished += OnParameterRenameFinished;
            _parameterRenameExternalEvent = ExternalEvent.Create(_parameterRenameHandler);

            // Create schedule fields handler
            _scheduleFieldsHandler = new ExternalEvents.ScheduleFieldsHandler();
            _scheduleFieldsHandler.OnUpdateFinished += OnScheduleFieldsUpdateFinished;
            _scheduleFieldsExternalEvent = ExternalEvent.Create(_scheduleFieldsHandler);

            // Create parameter load handler
            _parameterLoadHandler = new ExternalEvents.ParameterDataLoadHandler();
            _parameterLoadHandler.OnDataLoaded += OnParameterDataLoaded;
            _parameterLoadExternalEvent = ExternalEvent.Create(_parameterLoadHandler);

            // Create parameter value update handler
            _parameterValueUpdateHandler = new ExternalEvents.ParameterValueUpdateHandler();
            _parameterValueUpdateHandler.OnUpdateFinished += OnParameterValueUpdateFinished;
            _parameterValueUpdateExternalEvent = ExternalEvent.Create(_parameterValueUpdateHandler);

            // Create highlight in model handler
            _highlightInModelHandler = new ExternalEvents.HighlightInModelHandler();
            _highlightInModelExternalEvent = ExternalEvent.Create(_highlightInModelHandler);

            // Create save user settings handler (extensible storage)
            _saveSettingsHandler = new ExternalEvents.SaveUserSettingsHandler();
            _saveSettingsExternalEvent = ExternalEvent.Create(_saveSettingsHandler);


            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            DeferWindowShow();

            // Window infrastructure
            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;

            // Enforce Cell selection (Excel-like)
            ScheduleDataGrid.SelectionUnit = DataGridSelectionUnit.Cell;
            ScheduleDataGrid.SelectionMode = DataGridSelectionMode.Extended;
            
            // Selection Adorner Events
            ScheduleDataGrid.SelectedCellsChanged += (s, e) => UpdateSelectionAdorner();
            ScheduleDataGrid.LayoutUpdated += (s, e) => UpdateSelectionAdorner();
            ScheduleDataGrid.SizeChanged += (s, e) => UpdateSelectionAdorner();
            ScheduleDataGrid.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, e) => UpdateSelectionAdorner()));

            // Debug: Show selection count in title
            ScheduleDataGrid.SelectedCellsChanged += (s, e) => {
                // this.Title = $"Debug: Selected Cells = {ScheduleDataGrid.SelectedCells.Count}";
            };

            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            // Theme
            LoadThemeState();
            LoadWindowState();
            LoadManualModeState();
            DataContext = this;

            // Manual mode: cell highlighting via LoadingRow
            ScheduleDataGrid.LoadingRow += ScheduleDataGrid_LoadingRow_ManualMode;

            // Load persistent settings (must be before LoadData so they're available when schedule is restored)
            LoadSettingsFromStorage(app.ActiveUIDocument.Document);
            
            LoadData(app.ActiveUIDocument.Document);

            // Check for updates after window loads
            Loaded += (s, e) => 
            {
                Dispatcher.BeginInvoke(new Action(() => 
                {
                    Services.UpdateLogService.CheckAndShow(this);
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            };
        }


        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveSettingsToStorage();
            SaveWindowState();
            SaveManualModeState();
        }

        private void LoadData(Document doc)
        {
            _revitService = new RevitService(doc);
            
            var collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));
            _allSheets = new List<SheetItem>();

            foreach (ViewSheet sheet in collector)
            {
                if (sheet.IsPlaceholder) continue;
                var sheetItem = new SheetItem(sheet);
                sheetItem.State = SheetItemState.ExistingInRevit;
                sheetItem.OriginalSheetNumber = sheet.SheetNumber;
                sheetItem.OriginalName = sheet.Name;
                sheetItem.PropertyChanged += OnItemPropertyChanged;
                _allSheets.Add(sheetItem);
            }

            // Initially show all
            Sheets.Clear();
            FilteredSheets.Clear();
            foreach (var s in _allSheets.OrderBy(s => s.SheetNumber).ThenBy(s => s.Name))
            {
                Sheets.Add(s);
                FilteredSheets.Add(s);
            }

            // Load Schedules
            var schedules = _revitService.GetSchedules();
            var comboItems = new List<ScheduleOption>();
            comboItems.Add(new ScheduleOption { Name = "No Schedules Selected", Id = ElementId.InvalidElementId, Schedule = null });
            foreach(var s in schedules)
            {
                comboItems.Add(new ScheduleOption { Name = s.Name, Id = s.Id, Schedule = s });
            }
            SchedulesComboBox.ItemsSource = comboItems;
            
            // Try to restore last selected schedule from config
            string savedScheduleName = GetSavedScheduleName();
            if (!string.IsNullOrEmpty(savedScheduleName))
            {
                var savedItem = comboItems.FirstOrDefault(x => x.Name == savedScheduleName);
                if (savedItem != null)
                {
                    SchedulesComboBox.SelectedItem = savedItem;
                }
                else
                {
                    SchedulesComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                SchedulesComboBox.SelectedIndex = 0;
            }

            UpdateButtonStates();
            _isDataLoaded = true;
            TryShowWindow();
        }

        private void SchedulesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Guard: warn about pending parameter changes in manual mode
            if (_isManualMode && _pendingParameterChanges.Count > 0)
            {
                var previousItem = e.RemovedItems.Count > 0 ? e.RemovedItems[0] : null;
                ShowConfirmationPopup("Pending Changes",
                    "Switching schedules will discard pending parameter changes. Continue?",
                    () =>
                    {
                        _pendingParameterChanges.Clear();
                        ClearAllCellHighlights();
                        // Re-trigger selection change now that pending changes are cleared
                        SchedulesComboBox_SelectionChanged(sender, e);
                    },
                    () =>
                    {
                        // Revert selection
                        if (previousItem != null)
                        {
                            SchedulesComboBox.SelectionChanged -= SchedulesComboBox_SelectionChanged;
                            SchedulesComboBox.SelectedItem = previousItem;
                            SchedulesComboBox.SelectionChanged += SchedulesComboBox_SelectionChanged;
                        }
                    });
                return;
            }

            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;

            // Clear search and filters when switching schedules
            if (ScheduleSearchBox != null)
                ScheduleSearchBox.Clear();

            // Save selected schedule name for persistence
            if (selectedItem != null)
            {
                SaveScheduleName(selectedItem.Name);
            }
            
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
                {
                    // Save previous sorting (Deep Copy)
                    var list = new ObservableCollection<SortItem>();
                    foreach(var item in SortCriteria) list.Add(item.Clone());
                    _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
                }

                LoadScheduleData(selectedItem.Schedule);
                
                // Load sorting if exists (Deep Copy Restore)
                SortCriteria.Clear();
                if (_scheduleSortSettings.ContainsKey(selectedItem.Id))
                {
                    foreach(var item in _scheduleSortSettings[selectedItem.Id]) SortCriteria.Add(item.Clone());
                }

                // Load filters if exists (Deep Copy Restore)
                FilterCriteria.Clear();
                if (_scheduleFilterSettings.ContainsKey(selectedItem.Id))
                {
                    foreach(var item in _scheduleFilterSettings[selectedItem.Id]) FilterCriteria.Add(item.Clone());
                }

                // Restore Itemize Setting
                bool itemize = true;
                if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                {
                    itemize = _scheduleItemizeSettings[selectedItem.Id];
                }
                else
                {
                     _scheduleItemizeSettings[selectedItem.Id] = true; // Default
                }
                
                // Set CheckBox (this triggers checked/unchecked events which call RefreshScheduleView)
                // We need to avoid double refresh if possible, but simplest is to just set it.
                // However, the event handler calls ApplyCurrentSortLogic? No, we added that.
                
                // Temporarily detach events if we want to manually control flow, or just let it trigger.
                // Let's just set IsChecked. The event handler calls RefreshScheduleView(itemize) AND ApplyCurrentSortLogic().
                // This is exactly what we want.
                
                if (ItemizeCheckBox != null)
                {
                    ItemizeCheckBox.IsChecked = itemize;
                }
                
                // If the check state didn't change (already matched), the event won't fire.
                // In that case we must ensure data is loaded.
                // LoadScheduleData does NOT refresh the view passed the initial setup.
                
                // Force apply if event didn't trigger?
                // Actually, LoadScheduleData populates _rawScheduleData.
                // We need to call RefreshScheduleView at least once.
                
                // Let's force it manually if we suspect it might not trigger.
                // But changing source calls LoadScheduleData first. 
                
                // Better approach: 
                // 1. Set check box (might fire event)
                // 2. Ensure ApplyCurrentSortLogic works.
                
                // If I set IsChecked = itemize, and it was already itemize, no event fires.
                // We need to force refresh then.
                
                if (ItemizeCheckBox != null && ItemizeCheckBox.IsChecked == itemize)
                {
                   // Event won't fire, manually refresh
                   RefreshScheduleView(itemize);
                   ApplyCurrentSortLogic();
                   ApplyFilterLogic();
                }
            }
            else
            {
                // No schedule selected - show empty DataGrid
                ScheduleDataGrid.Columns.Clear();
                ScheduleDataGrid.ItemsSource = null;
                _currentScheduleData = null;
                AvailableSortColumns.Clear();
            }


        }

        private void RestoreDataView()
        {
            ScheduleDataGrid.ItemsSource = null;
            ScheduleDataGrid.Columns.Clear();
            
            ScheduleDataGrid.ItemsSource = FilteredSheets;
            ScheduleDataGrid.AutoGenerateColumns = false;
            
            InitializeDataColumns();
        }

        private void InitializeDataColumns()
        {
            ScheduleDataGrid.Columns.Clear();
            
            var checkBoxColumn = CreateCheckBoxColumn();
            ScheduleDataGrid.Columns.Add(checkBoxColumn);
            
            var numberCol = new DataGridTextColumn
            {
                Header = "Sheet Number",
                Binding = new System.Windows.Data.Binding("SheetNumber"),
                Width = new DataGridLength(150)
            };
            ScheduleDataGrid.Columns.Add(numberCol);
            
            var nameCol = new DataGridTextColumn
            {
                Header = "Sheet Name",
                Binding = new System.Windows.Data.Binding("Name"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            };
            ScheduleDataGrid.Columns.Add(nameCol);
        }

        private DataGridTemplateColumn CreateCheckBoxColumn()
        {
            var headerCheckBox = new CheckBox
            {
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Style = (Style)FindResource("CustomCheckBoxStyle")
            };
            headerCheckBox.Checked += HeaderCheckBox_Checked;
            headerCheckBox.Unchecked += HeaderCheckBox_Unchecked;

            var baseStyle = FindResource("ExcelLikeCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center));
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center));

            // Hide checkbox cell content on blank separator rows
            var blankRowTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "BlankLine"
            };
            blankRowTrigger.Setters.Add(new Setter(DataGridCell.ContentTemplateProperty, new DataTemplate()));
            blankRowTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.IsEnabledProperty, false));
            blankRowTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty,
                new System.Windows.DynamicResourceExtension("BlankLineBrush")));
            cellStyle.Triggers.Add(blankRowTrigger);

            // Hide checkbox cell content on footer rows
            var footerRowTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "FooterLine"
            };
            footerRowTrigger.Setters.Add(new Setter(DataGridCell.ContentTemplateProperty, new DataTemplate()));
            footerRowTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.IsEnabledProperty, false));
            footerRowTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty,
                new System.Windows.DynamicResourceExtension("FooterLineBrush")));
            cellStyle.Triggers.Add(footerRowTrigger);

            // Create template for checkbox cells
            var cellTemplate = new DataTemplate();
            var checkBoxFactory = new FrameworkElementFactory(typeof(CheckBox));
            checkBoxFactory.SetValue(CheckBox.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.StyleProperty, FindResource("CustomCheckBoxStyle"));
            checkBoxFactory.SetBinding(CheckBox.IsCheckedProperty, new System.Windows.Data.Binding("IsSelected") 
            { 
                Mode = System.Windows.Data.BindingMode.TwoWay, 
                UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged 
            });
            checkBoxFactory.AddHandler(CheckBox.ClickEvent, new RoutedEventHandler(RowCheckBox_Click));
            cellTemplate.VisualTree = checkBoxFactory;


            var col = new DataGridTemplateColumn
            {
                Header = headerCheckBox,
                CellTemplate = cellTemplate,
                Width = new DataGridLength(40),
                CellStyle = cellStyle
            };
            return col;
        }

        private void RowCheckBox_Click(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as CheckBox;
            if (checkBox == null) return;
            bool isChecked = checkBox.IsChecked == true;

            // 1. Handle "SelectedItems" (Full Row Selection mode - though we are in Cell mode, this might still be populated if user selects full rows)
            if (ScheduleDataGrid.SelectedItems.Count > 1)
            {
               foreach (var selectedItem in ScheduleDataGrid.SelectedItems)
                {
                    SetRowSelection(selectedItem, isChecked);
                }
            }

            // 2. Handle "SelectedCells" (Cell Selection mode)
            // If the user selected a range of cells in the first column (Checkbox column), we want to toggle all of them.
            // Checkbox column is usually index 0.
            if (ScheduleDataGrid.SelectedCells.Count > 1)
            {
                foreach(var cellInfo in ScheduleDataGrid.SelectedCells)
                {
                    // Check if this cell is in the Checkbox column
                    // We can check Column.DisplayIndex or checking the content. 
                    // Our Checkbox column is a DataGridTemplateColumn created in code.
                    // Let's assume it's the one with index 0.
                    if (cellInfo.Column.DisplayIndex == 0)
                    {
                        SetRowSelection(cellInfo.Item, isChecked);
                    }
                }
            }
        }

        private void SetRowSelection(object item, bool isSelected)
        {
            if (item is System.Data.DataRowView rowView)
            {
                rowView["IsSelected"] = isSelected;
            }
            else if (item is SheetItem sheet)
            {
                sheet.IsSelected = isSelected;
            }
        }

        private void HeaderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var view = ScheduleDataGrid.ItemsSource;
            if (view is System.Data.DataView dataView)
            {
                foreach (System.Data.DataRowView row in dataView)
                {
                    row["IsSelected"] = true;
                }
            }
            else if (view is ObservableCollection<SheetItem> sheets)
            {
                foreach (var sheet in sheets)
                {
                    sheet.IsSelected = true;
                }
            }
        }

        private void HeaderCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            var view = ScheduleDataGrid.ItemsSource;
            if (view is System.Data.DataView dataView)
            {
                foreach (System.Data.DataRowView row in dataView)
                {
                    row["IsSelected"] = false;
                }
            }
            else if (view is ObservableCollection<SheetItem> sheets)
            {
                foreach (var sheet in sheets)
                {
                    sheet.IsSelected = false;
                }
            }
        }



        private void LoadScheduleData(ViewSchedule schedule)
        {
            try
            {
                var data = _revitService.GetScheduleData(schedule);
                _currentScheduleData = data;
                
                if (data == null) throw new Exception("Failed to retrieve schedule data from Revit.");

                if (!_scheduleItemizeSettings.ContainsKey(schedule.Id))
                {
                    _scheduleItemizeSettings[schedule.Id] = true;
                }
                bool isItemized = _scheduleItemizeSettings[schedule.Id];
                
                // Removed UI checkbox update from here to prevent premature event firing.
                // It is now handled in SelectionChanged after SortCriteria is restored.
                
                var dt = new System.Data.DataTable();
                dt.Columns.Add("IsSelected", typeof(bool)).DefaultValue = false;
                dt.Columns.Add("RowState", typeof(string)).DefaultValue = "Unchanged";
                dt.Columns.Add("Count", typeof(int));

                var newSortColumns = new List<string> { "(none)" };

                // Detect column types
                for(int i = 0; i < data.Columns.Count; i++)
                {
                    string safeName = data.Columns[i];
                    int dupIdx = 1;
                    while(dt.Columns.Contains(safeName))
                    {
                        safeName = $"{data.Columns[i]} ({dupIdx++})";
                    }

                    // Filter out internal columns from sorting options
                    if (safeName != "ElementId" && safeName != "TypeName" && !safeName.StartsWith("Count"))
                    {
                        newSortColumns.Add(safeName);
                    }


                    // Check if column is numeric
                    bool isNumeric = true;
                    bool hasValue = false;

                    if (safeName == "ElementId")
                    {
                        isNumeric = false;
                        hasValue = true; // Force string creation
                    }
                    else
                    {
                        foreach(var r in data.Rows)
                        {
                            string val = r[i];
                            if (string.IsNullOrWhiteSpace(val)) continue;
                            hasValue = true;
                            // Values with leading zeros (e.g. "01") must stay as strings
                            if (val.Length > 1 && val[0] == '0')
                            {
                                isNumeric = false;
                                break;
                            }
                            if (!double.TryParse(val, out _))
                            {
                                isNumeric = false;
                                break;
                            }
                        }
                    }

                    if (isNumeric && hasValue)
                    {
                        dt.Columns.Add(safeName, typeof(double));
                    }
                    else
                    {
                        dt.Columns.Add(safeName, typeof(string));
                    }
                }

                // Smart Update AvailableSortColumns to preserve bindings
                bool updateNeeded = AvailableSortColumns.Count != newSortColumns.Count;
                if (!updateNeeded)
                {
                    for (int i = 0; i < newSortColumns.Count; i++)
                    {
                        if (AvailableSortColumns[i] != newSortColumns[i])
                        {
                            updateNeeded = true;
                            break;
                        }
                    }
                }

                if (updateNeeded)
                {
                    AvailableSortColumns.Clear();
                    foreach (var col in newSortColumns)
                    {
                        AvailableSortColumns.Add(col);
                    }
                }

                // Update AvailableFilterColumns
                var skipFilterCols = new[] { "IsSelected", "RowState", "ElementId", "TypeName", "Count" };
                AvailableFilterColumns.Clear();
                AvailableFilterColumns.Add("(none)");
                foreach (System.Data.DataColumn col in dt.Columns)
                {
                    if (!skipFilterCols.Contains(col.ColumnName))
                        AvailableFilterColumns.Add(col.ColumnName);
                }

                foreach(var row in data.Rows)
                {
                    var newRow = dt.NewRow();
                    newRow["IsSelected"] = false;
                    newRow["RowState"] = "Unchanged";
                    newRow["Count"] = 1;
                    
                    for(int i = 0; i < data.Columns.Count; i++)
                    {
                        string val = row[i];
                        // If column is numeric, parse it
                        if (dt.Columns[i + 3].DataType == typeof(double))
                        {
                            if (double.TryParse(val, out double dVal))
                            {
                                newRow[i + 3] = dVal;
                            }
                            else
                            {
                                newRow[i + 3] = DBNull.Value;
                            }
                        }
                        else
                        {
                            newRow[i + 3] = val;
                        }
                    }
                    
                    dt.Rows.Add(newRow);
                }
                
                _rawScheduleData = dt;
                // RefreshScheduleView(isItemized); <-- Removed to prevent refresh with stale SortCriteria
            }
            catch (Exception ex)
            {
                ShowPopup("Error Loading Schedule", ex.Message);
            }
        }

        private void RefreshScheduleView(bool itemize)
        {
            System.Data.DataTable viewTable;

            if (HasAnySpecialRows())
            {
                // Pre-sorted table with blank/footer rows baked in
                viewTable = BuildSortedTableWithBlanks(itemize);
            }
            else if (!itemize && _rawScheduleData != null)
            {
                viewTable = _rawScheduleData.Clone();
                // Group by sorting rules
                var validSorts = SortCriteria.Where(s => s.SelectedColumn != "(none)" && !string.IsNullOrEmpty(s.SelectedColumn))
                                             .Select(s => s.SelectedColumn)
                                             .ToList();

                var grouped = _rawScheduleData.AsEnumerable()
                    .GroupBy(r => 
                    {
                        if (validSorts.Count == 0) return ""; // Group all if no sort
                        
                        // Create composite key
                        return string.Join("||", validSorts.Select(col => 
                            r.Table.Columns.Contains(col) ? r[col]?.ToString() ?? "" : ""));
                    });
                
                foreach(var grp in grouped)
                {
                    var grpList = grp.ToList();
                    var firstRow = grpList[0];
                    var newRow = viewTable.NewRow();
                    newRow.ItemArray = (object[])firstRow.ItemArray.Clone();
                    newRow["Count"] = grpList.Count;

                    // Aggregate ElementIds if present
                    if (viewTable.Columns.Contains("ElementId"))
                    {
                        var ids = grpList.Select(r => r["ElementId"]?.ToString())
                                     .Where(s => !string.IsNullOrEmpty(s));
                        newRow["ElementId"] = string.Join(",", ids);
                    }

                    ApplyVaries(grpList, newRow, validSorts);
                    viewTable.Rows.Add(newRow);
                }
            }
            else
            {
                viewTable = _rawScheduleData;
            }

            ScheduleDataGrid.ItemsSource = null;
            ScheduleDataGrid.Columns.Clear();
            ScheduleDataGrid.AutoGenerateColumns = false;
            
            if (viewTable == null) return;
            
            var baseStyle = FindResource("ExcelLikeCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            var cellTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "Pending"
            };
            cellTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(Colors.Yellow) { Opacity = 0.5 }));
            cellStyle.Triggers.Add(cellTrigger);

            var blankLineTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "BlankLine"
            };
            var blankBrush = new System.Windows.DynamicResourceExtension("BlankLineBrush");
            blankLineTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, blankBrush));
            blankLineTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.IsEnabledProperty, false));
            cellStyle.Triggers.Add(blankLineTrigger);

            var footerLineTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "FooterLine"
            };
            var footerBrush = new System.Windows.DynamicResourceExtension("FooterLineBrush");
            footerLineTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, footerBrush));
            footerLineTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.IsEnabledProperty, false));
            cellStyle.Triggers.Add(footerLineTrigger);

            // Row-level style: colors the whole row (including checkbox column) for blank/footer rows
            var baseRowStyle = FindResource("ExcelLikeRowStyle") as Style;
            var rowStyle = new Style(typeof(DataGridRow), baseRowStyle);
            var blankRowTrigger2 = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Row[RowState]"),
                Value = "BlankLine"
            };
            blankRowTrigger2.Setters.Add(new Setter(DataGridRow.BackgroundProperty,
                new System.Windows.DynamicResourceExtension("BlankLineBrush")));
            rowStyle.Triggers.Add(blankRowTrigger2);
            var footerRowTrigger2 = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Row[RowState]"),
                Value = "FooterLine"
            };
            footerRowTrigger2.Setters.Add(new Setter(DataGridRow.BackgroundProperty,
                new System.Windows.DynamicResourceExtension("FooterLineBrush")));
            rowStyle.Triggers.Add(footerRowTrigger2);
            ScheduleDataGrid.RowStyle = rowStyle;

            // First add checkbox column
            var checkCol = CreateCheckBoxColumn();
            ScheduleDataGrid.Columns.Add(checkCol);
            
            // Then add schedule data columns (skip RowState, ElementId, Count, IsSelected)
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "Count", "TypeName" };
            foreach(System.Data.DataColumn col in viewTable.Columns)
            {
                if (skipColumns.Contains(col.ColumnName)) continue;
                
                // Skip columns that end with " (1)", " (2)", etc. - these are duplicates
                if (System.Text.RegularExpressions.Regex.IsMatch(col.ColumnName, @"\s\(\d+\)$")) continue;
                
                var textCol = new DataGridTextColumn
                {
                    Header = col.ColumnName,
                    Binding = new System.Windows.Data.Binding(col.ColumnName),
                    CellStyle = cellStyle,
                    EditingElementStyle = (Style)FindResource("DataGridEditingStyle"),
                    IsReadOnly = false
                };
                ScheduleDataGrid.Columns.Add(textCol);
            }
            
            // Finally add Count column at the end
            if (viewTable.Columns.Contains("Count"))
            {
                var countCol = new DataGridTextColumn
                {
                    Header = "Count",
                    Binding = new System.Windows.Data.Binding("Count"),
                    CellStyle = cellStyle,
                    IsReadOnly = true
                };
                ScheduleDataGrid.Columns.Add(countCol);
            }
            
            ScheduleDataGrid.ItemsSource = new System.Data.DataView(viewTable);
            
            // Subscribe to cell editing event (unsubscribe first to avoid duplicates)
            ScheduleDataGrid.CellEditEnding -= ScheduleDataGrid_CellEditEnding;
            ScheduleDataGrid.CellEditEnding += ScheduleDataGrid_CellEditEnding;
            ScheduleDataGrid.BeginningEdit -= ScheduleDataGrid_BeginningEdit;
            ScheduleDataGrid.BeginningEdit += ScheduleDataGrid_BeginningEdit;
        }

        private void Itemize_Checked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = true;
                RefreshScheduleView(true);
                ApplyCurrentSortLogic();
                ApplyFilterLogic();
            }
        }

        private void Itemize_Unchecked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = false;
                // Calling ApplyCurrentSortLogic will see the 'false' setting and trigger RefreshScheduleView(false) internally
                // Then it will proceed to apply SortDescriptions.
                ApplyCurrentSortLogic();
                ApplyFilterLogic();
            }
        }

        private async void ScheduleDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            
            // Get the edited cell data
            var row = e.Row.Item as System.Data.DataRowView;
            if (row == null || _currentScheduleData == null) return;

            string columnName = e.Column.Header?.ToString();
            if (string.IsNullOrEmpty(columnName)) return;

            // Skip non-parameter columns
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "TypeName", "Count" };
            if (skipColumns.Contains(columnName)) return;

            // Get the new value from the editing element
            var editingElement = e.EditingElement as System.Windows.Controls.TextBox;
            if (editingElement == null) return;

            string newValue = editingElement.Text;
            string oldValue = row[columnName]?.ToString() ?? "";

            // If value hasn't changed, skip
            if (newValue == oldValue) return;

            // Get element ID from the row
            if (!row.Row.Table.Columns.Contains("ElementId")) return;
            string elementIdStr = row["ElementId"]?.ToString();
            if (string.IsNullOrEmpty(elementIdStr)) return;

            // Find the parameter ID for this column
            int columnIndex = _currentScheduleData.Columns.IndexOf(columnName);
            if (columnIndex < 0 || columnIndex >= _currentScheduleData.ParameterIds.Count) return;

            ElementId parameterId = _currentScheduleData.ParameterIds[columnIndex];
            
            // Check if it's a type parameter
            bool isTypeParameter = _currentScheduleData.IsTypeParameter.ContainsKey(columnName) && 
                                   _currentScheduleData.IsTypeParameter[columnName];

            // Show confirmation for type parameters
            if (isTypeParameter)
            {
                // Store values for the confirmation callback
                var tempElementIdStr = elementIdStr;
                var tempParameterId = parameterId;
                var tempNewValue = newValue;
                var tempOldValue = oldValue;
                var tempRow = row;
                var tempColumnName = columnName;

                ShowConfirmationPopup(
                    "Type Parameter Warning",
                    "This is a TYPE parameter. Changing it will affect ALL elements of this type.\n\nDo you want to proceed?",
                    () => PerformParameterUpdate(tempElementIdStr, tempParameterId, tempNewValue, tempOldValue, tempRow, tempColumnName),
                    () => 
                    {
                        // Cancelled - revert the value in the UI
                        // Since CellEditEnding happens after commit, the DataTable has the new value.
                        // We need to set it back to the old value.
                        tempRow.Row[tempColumnName] = tempOldValue;
                    });
                return;
            }

            // For instance parameters, update immediately
            PerformParameterUpdate(elementIdStr, parameterId, newValue, oldValue, row, columnName);
        }

        private void PerformParameterUpdate(string elementIdStr, ElementId parameterId, string newValue,
                                                  string oldValue, System.Data.DataRowView row, string columnName)
        {
            if (_isManualMode)
            {
                // Queue the change instead of sending to Revit
                var existing = _pendingParameterChanges.FirstOrDefault(p =>
                    p.ElementIdStr == elementIdStr &&
                    p.ParameterId.Value == parameterId.Value &&
                    p.ColumnName == columnName);

                if (existing != null)
                {
                    existing.NewValue = newValue;
                    // If user reverted to original, remove the pending change
                    if (existing.NewValue == existing.OldValue)
                    {
                        _pendingParameterChanges.Remove(existing);
                        ClearCellHighlight(row, columnName);
                        return;
                    }
                }
                else
                {
                    int rowIndex = row.Row.Table.Rows.IndexOf(row.Row);
                    _pendingParameterChanges.Add(new PendingParameterChange
                    {
                        ElementIdStr = elementIdStr,
                        ParameterId = parameterId,
                        NewValue = newValue,
                        OldValue = oldValue,
                        ColumnName = columnName,
                        RowIndex = rowIndex
                    });
                }

                HighlightPendingCell(row, columnName);
                return;
            }

            // Auto mode: update the parameter value via external event immediately
            _parameterValueUpdateHandler.IsBatchMode = false;
            _parameterValueUpdateHandler.ElementIdStr = elementIdStr;
            _parameterValueUpdateHandler.ParameterIdStr = parameterId.Value.ToString();
            _parameterValueUpdateHandler.NewValue = newValue;

            _parameterValueUpdateExternalEvent.Raise();

            // UI Update is now handled by OnParameterValueUpdateFinished
        }

        private void OnDuplicationFinished(int success, int fail, string errorMsg, List<ElementId> newSheetIds)
        {
            Dispatcher.Invoke(() =>
            {
                // Update pending creation sheets with new IDs
                if (newSheetIds != null && newSheetIds.Count > 0)
                {
                    var pendingCreations = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                    for (int i = 0; i < Math.Min(pendingCreations.Count, newSheetIds.Count); i++)
                    {
                        pendingCreations[i].Id = newSheetIds[i];
                        pendingCreations[i].State = SheetItemState.ExistingInRevit;
                        pendingCreations[i].OriginalSheetNumber = pendingCreations[i].SheetNumber;
                        pendingCreations[i].OriginalName = pendingCreations[i].Name;
                    }
                }

                // Reload data to show all sheets
                if (_uiApplication != null)
                {
                    LoadData(_uiApplication.ActiveUIDocument.Document);
                }

                UpdateButtonStates();

                // Show Result Popup
                if (fail > 0)
                {
                    ShowPopup("Duplication Report", $"Success: {success}\nFailures: {fail}\nLast Error: {errorMsg}");
                }
                else
                {
                    ShowPopup("Success", $"Successfully created {success} item(s).");
                }
            });
        }

        #endregion

        #region Actions

        private void AddDuplicates_Click(object sender, RoutedEventArgs e)
        {
            // Validate selection
            var selectedCount = Sheets.Count(x => x.IsSelected);
            if (selectedCount == 0)
            {
                ShowPopup("No Items Selected", "Please select at least one item to duplicate.");
                return;
            }

            OptCopies.Text = "1";
            OptKeepViews.IsChecked = false;
            OptKeepLegends.IsChecked = false;
            OptKeepSchedules.IsChecked = false;
            OptCopyRevisions.IsChecked = true;
            OptCopyParams.IsChecked = true;

            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void Popup_MouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true; // Prevent click from bubbling to background
        }

        private void OptionsPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Close on background click? User choice, but safe to allow cancelling
            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void DuplicateCancel_Click(object sender, RoutedEventArgs e)
        {
            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void DuplicateConfirm_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Validate and parse options
                if (!int.TryParse(OptCopies.Text, out int copies) || copies < 1)
                {
                    ShowPopup("Invalid Input", "Please enter a valid number of copies (1 or more).");
                    return;
                }

                // 2. Build DuplicateOptions object
                var options = new DuplicateOptions
                {
                    NumberOfCopies = copies,
                    DuplicateMode = OptKeepViews.IsChecked == true
                        ? ExternalEvents.SheetDuplicateMode.WithViews
                        : ExternalEvents.SheetDuplicateMode.WithSheetDetailing,
                    KeepLegends = OptKeepLegends.IsChecked == true,
                    KeepSchedules = OptKeepSchedules.IsChecked == true,
                    CopyRevisions = OptCopyRevisions.IsChecked == true,
                    CopyParameters = OptCopyParams.IsChecked == true
                };

                // 3. Create pending SheetItems
                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                foreach (var sourceSheet in selectedSheets)
                {
                    for (int i = 0; i < copies; i++)
                    {
                        var pendingSheet = new SheetItem(
                            sheetNumber: $"{sourceSheet.SheetNumber} - Copy {i + 1}",
                            name: $"{sourceSheet.Name} - Copy {i + 1}",
                            sourceSheetId: sourceSheet.Id,
                            options: options
                        );

                        // Subscribe to property changes
                        pendingSheet.PropertyChanged += OnItemPropertyChanged;

                        // Add to collections
                        Sheets.Add(pendingSheet);

                        // Check if it matches current search filter
                        var searchText = ScheduleSearchBox?.Text?.ToLowerInvariant() ?? "";
                        if (string.IsNullOrEmpty(searchText) ||
                            pendingSheet.Name.ToLowerInvariant().Contains(searchText) ||
                            pendingSheet.SheetNumber.ToLowerInvariant().Contains(searchText))
                        {
                            FilteredSheets.Add(pendingSheet);
                        }
                    }
                }

                // 4. Update button states
                UpdateButtonStates();

                // 5. Close Popup
                OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private SortingWindow _sortingWindow;
        private FilterWindow _filterWindow;

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            if (_filterWindow == null || !_filterWindow.IsLoaded)
            {
                _filterWindow = new FilterWindow(this);
                _filterWindow.Owner = this;
                _filterWindow.Show();
            }
            else
            {
                _filterWindow.Activate();
                if (_filterWindow.WindowState == WindowState.Minimized)
                    _filterWindow.WindowState = WindowState.Normal;
            }
        }

        private void Sort_Click(object sender, RoutedEventArgs e)
        {


            if (_sortingWindow == null || !_sortingWindow.IsLoaded)
            {
                _sortingWindow = new SortingWindow(this);
                _sortingWindow.Owner = this;
                _sortingWindow.Show();
            }
            else
            {
                _sortingWindow.Activate();
                if (_sortingWindow.WindowState == WindowState.Minimized)
                    _sortingWindow.WindowState = WindowState.Normal;
            }
        }



        internal void ApplyCurrentSortLogicInternal()
        {
            ApplyCurrentSortLogic();
        }

        internal List<string> GetUniqueValuesForColumn(string columnName)
        {
            if (_rawScheduleData == null || !_rawScheduleData.Columns.Contains(columnName))
                return new List<string>();

            return _rawScheduleData.AsEnumerable()
                .Select(r => r[columnName]?.ToString() ?? "")
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct()
                .OrderBy(v => v)
                .ToList();
        }

        internal void ApplyFilterLogic()
        {
            if (ScheduleDataGrid.ItemsSource is System.Data.DataView dataView)
            {
                var activeFilters = FilterCriteria
                    .Where(f => !string.IsNullOrEmpty(f.SelectedColumn) && f.SelectedColumn != "(none)")
                    .ToList();

                if (activeFilters.Count == 0)
                {
                    dataView.RowFilter = string.Empty;
                    return;
                }

                var parts = new List<string>();
                foreach (var f in activeFilters)
                {
                    string col = f.SelectedColumn.Replace("]", "]]");
                    string val = (f.Value ?? "").Replace("'", "''");

                    string expr = null;
                    switch (f.SelectedCondition)
                    {
                        case "equals":
                            expr = $"CONVERT([{col}], System.String) = '{val}'";
                            break;
                        case "does not equal":
                            expr = $"CONVERT([{col}], System.String) <> '{val}'";
                            break;
                        case "is greater than":
                            expr = $"[{col}] > '{val}'";
                            break;
                        case "is greater than or equal to":
                            expr = $"[{col}] >= '{val}'";
                            break;
                        case "is less than":
                            expr = $"[{col}] < '{val}'";
                            break;
                        case "is less than or equal to":
                            expr = $"[{col}] <= '{val}'";
                            break;
                        case "contains":
                            expr = $"CONVERT([{col}], System.String) LIKE '%{val}%'";
                            break;
                        case "does not contain":
                            expr = $"CONVERT([{col}], System.String) NOT LIKE '%{val}%'";
                            break;
                        case "begins with":
                            expr = $"CONVERT([{col}], System.String) LIKE '{val}%'";
                            break;
                        case "does not begin with":
                            expr = $"CONVERT([{col}], System.String) NOT LIKE '{val}%'";
                            break;
                        case "ends with":
                            expr = $"CONVERT([{col}], System.String) LIKE '%{val}'";
                            break;
                        case "does not end with":
                            expr = $"CONVERT([{col}], System.String) NOT LIKE '%{val}'";
                            break;
                        case "has a value":
                            expr = $"CONVERT([{col}], System.String) <> '' AND [{col}] IS NOT NULL";
                            break;
                        case "has no value":
                            expr = $"(CONVERT([{col}], System.String) = '' OR [{col}] IS NULL)";
                            break;
                    }

                    if (expr != null) parts.Add($"({expr})");
                }

                // Always let separator/footer rows through so they're never hidden by active filters
                dataView.RowFilter = parts.Count > 0
                    ? $"(RowState = 'BlankLine') OR (RowState = 'FooterLine') OR ({string.Join(" AND ", parts)})"
                    : string.Empty;
            }
        }

        private void ApplyCurrentSortLogic()
        {
            if (ScheduleDataGrid.ItemsSource == null) return;

            bool isNonItemized = _currentScheduleData != null &&
                                 _scheduleItemizeSettings.ContainsKey(_currentScheduleData.ScheduleId) &&
                                 _scheduleItemizeSettings[_currentScheduleData.ScheduleId] == false;

            // When blank/footer lines are active the sort order is baked into the DataTable —
            // we must rebuild rather than apply SortDescriptions on the existing view.
            if (HasAnySpecialRows())
            {
                RefreshScheduleView(!isNonItemized);
                return;
            }

            // Non-itemized (grouped) mode must also rebuild to regroup
            if (isNonItemized)
            {
                RefreshScheduleView(false);
            }

            System.ComponentModel.ICollectionView view = System.Windows.Data.CollectionViewSource.GetDefaultView(ScheduleDataGrid.ItemsSource);
            view.SortDescriptions.Clear();

            foreach (var sortItem in SortCriteria)
            {
                if (string.IsNullOrEmpty(sortItem.SelectedColumn) || sortItem.SelectedColumn == "(none)") continue;
                
                string propertyName = sortItem.SelectedColumn;
                if (propertyName == "Sheet Number") propertyName = "SheetNumber";
                else if (propertyName == "Sheet Name") propertyName = "Name";
                
                view.SortDescriptions.Add(new System.ComponentModel.SortDescription(
                    propertyName, 
                    sortItem.IsAscending ? System.ComponentModel.ListSortDirection.Ascending : System.ComponentModel.ListSortDirection.Descending
                ));
            }
        }

        private bool HasAnyBlankLineSorts() =>
            SortCriteria.Any(s => s.ShowBlankLine &&
                                  !string.IsNullOrEmpty(s.SelectedColumn) &&
                                  s.SelectedColumn != "(none)");

        private bool HasAnyFooterSorts() =>
            SortCriteria.Any(s => s.ShowFooter &&
                                  !string.IsNullOrEmpty(s.SelectedColumn) &&
                                  s.SelectedColumn != "(none)");

        private bool HasAnySpecialRows() => HasAnyBlankLineSorts() || HasAnyFooterSorts();

        private System.Data.DataTable BuildSortedTableWithBlanks(bool itemize)
        {
            var activeSorts = SortCriteria
                .Where(s => !string.IsNullOrEmpty(s.SelectedColumn) && s.SelectedColumn != "(none)")
                .ToList();

            var blankLineSorts = activeSorts
                .Where(s => s.ShowBlankLine && _rawScheduleData.Columns.Contains(s.SelectedColumn))
                .ToList();

            var footerSorts = activeSorts
                .Where(s => s.ShowFooter && _rawScheduleData.Columns.Contains(s.SelectedColumn))
                .ToList();

            var blankLineCols = blankLineSorts.Select(s => s.SelectedColumn).ToList();

            // Sort raw data via DataView (respects column types)
            var sortedView = new System.Data.DataView(_rawScheduleData);
            var sortParts = activeSorts
                .Where(s => _rawScheduleData.Columns.Contains(s.SelectedColumn))
                .Select(s => $"[{s.SelectedColumn.Replace("]", "]]")}] {(s.IsAscending ? "ASC" : "DESC")}");
            string sortExpr = string.Join(", ", sortParts);
            if (!string.IsNullOrEmpty(sortExpr))
                sortedView.Sort = sortExpr;

            var result = _rawScheduleData.Clone();

            if (!itemize)
            {
                // Grouped mode: merge rows sharing the same composite sort key.
                // Blank separator: inserted only when blank-line columns change (between groups).
                // Footer row: inserted only when footer column changes (accumulates count across sub-groups).
                var groupCols = activeSorts
                    .Select(s => s.SelectedColumn)
                    .Where(c => _rawScheduleData.Columns.Contains(c))
                    .ToList();

                string prevGroupKey      = null;
                string prevBlankKey      = null;
                string prevFooterKey     = null;
                int    footerAccumCount  = 0;
                System.Data.DataRow footerFirstDataRow = null;
                var    groupRows         = new List<System.Data.DataRow>();

                void FlushFooter()
                {
                    if (footerSorts.Count == 0 || prevFooterKey == null || footerAccumCount == 0) return;
                    var fr = result.NewRow();
                    fr["IsSelected"] = false;
                    fr["RowState"] = "FooterLine";
                    AddFooterContent(fr, footerSorts[0], footerFirstDataRow, footerAccumCount);
                    result.Rows.Add(fr);
                    footerAccumCount = 0;
                    footerFirstDataRow = null;
                }

                void FlushGroup()
                {
                    if (groupRows.Count == 0) return;

                    string blankKey = blankLineCols.Count > 0
                        ? string.Join("||", blankLineCols.Select(c =>
                              groupRows[0].Table.Columns.Contains(c)
                                  ? groupRows[0][c]?.ToString() ?? ""
                                  : ""))
                        : null;

                    string footerKey = footerSorts.Count > 0 && groupRows[0].Table.Columns.Contains(footerSorts[0].SelectedColumn)
                        ? groupRows[0][footerSorts[0].SelectedColumn]?.ToString() ?? ""
                        : null;

                    // Footer key changed → emit footer for previous footer group first
                    if (footerKey != null && prevFooterKey != null && footerKey != prevFooterKey)
                        FlushFooter();

                    // Blank key changed → emit blank separator
                    if (prevBlankKey != null && blankKey != prevBlankKey)
                    {
                        var blank = result.NewRow();
                        blank["IsSelected"] = false;
                        blank["RowState"] = "BlankLine";
                        result.Rows.Add(blank);
                    }
                    prevBlankKey = blankKey;

                    var newRow = result.NewRow();
                    newRow.ItemArray = (object[])groupRows[0].ItemArray.Clone();
                    newRow["RowState"] = "Unchanged";
                    newRow["Count"] = groupRows.Count;
                    if (result.Columns.Contains("ElementId"))
                        newRow["ElementId"] = string.Join(",",
                            groupRows.Select(r => r["ElementId"]?.ToString())
                                     .Where(s => !string.IsNullOrEmpty(s)));

                    ApplyVaries(groupRows, newRow, groupCols);
                    result.Rows.Add(newRow);

                    // Accumulate into footer group
                    if (footerKey != null)
                    {
                        if (footerFirstDataRow == null) footerFirstDataRow = groupRows[0];
                        footerAccumCount += groupRows.Count;
                        prevFooterKey = footerKey;
                    }

                    groupRows.Clear();
                }

                foreach (System.Data.DataRowView drv in sortedView)
                {
                    string key = groupCols.Count > 0
                        ? string.Join("||", groupCols.Select(c => drv[c]?.ToString() ?? ""))
                        : "";
                    if (prevGroupKey != null && key != prevGroupKey) FlushGroup();
                    prevGroupKey = key;
                    groupRows.Add(drv.Row);
                }
                FlushGroup();
                FlushFooter(); // emit footer for the last footer group
            }
            else
            {
                // Itemized mode: one row per data row.
                // Footer and blank rows are emitted when their respective tracked columns change.
                string prevBlankKey     = null;
                string prevFooterKey    = null;
                int    footerAccumCount = 0;
                System.Data.DataRow footerFirstDataRow = null;

                void FlushFooter()
                {
                    if (footerSorts.Count == 0 || prevFooterKey == null || footerAccumCount == 0) return;
                    var fr = result.NewRow();
                    fr["IsSelected"] = false;
                    fr["RowState"] = "FooterLine";
                    AddFooterContent(fr, footerSorts[0], footerFirstDataRow, footerAccumCount);
                    result.Rows.Add(fr);
                    footerAccumCount = 0;
                    footerFirstDataRow = null;
                }

                foreach (System.Data.DataRowView drv in sortedView)
                {
                    string blankKey = blankLineCols.Count > 0
                        ? string.Join("||", blankLineCols.Select(c => drv[c]?.ToString() ?? ""))
                        : null;
                    string footerKey = footerSorts.Count > 0 && drv.DataView.Table.Columns.Contains(footerSorts[0].SelectedColumn)
                        ? drv[footerSorts[0].SelectedColumn]?.ToString() ?? ""
                        : null;

                    // Footer key changed → emit footer for previous group
                    if (footerKey != null && prevFooterKey != null && footerKey != prevFooterKey)
                        FlushFooter();

                    // Blank key changed → emit blank separator
                    if (blankKey != null && prevBlankKey != null && blankKey != prevBlankKey)
                    {
                        var blank = result.NewRow();
                        blank["IsSelected"] = false;
                        blank["RowState"] = "BlankLine";
                        result.Rows.Add(blank);
                    }
                    prevBlankKey = blankKey;
                    prevFooterKey = footerKey;

                    // Accumulate into footer group
                    if (footerKey != null)
                    {
                        if (footerFirstDataRow == null) footerFirstDataRow = drv.Row;
                        footerAccumCount++;
                    }

                    var dataRow = result.NewRow();
                    dataRow.ItemArray = (object[])drv.Row.ItemArray.Clone();
                    result.Rows.Add(dataRow);
                }

                FlushFooter(); // emit footer for the last group
            }

            return result;
        }

        private void AddFooterContent(System.Data.DataRow footerRow, SortItem footerSort,
                                      System.Data.DataRow firstDataRow, int count)
        {
            string opt = footerSort.FooterOption ?? "Title, count, and totals";
            bool showTitle = opt == "Title, count, and totals" || opt == "Title and totals";
            bool showCount = opt == "Title, count, and totals" || opt == "Count and totals";

            var parts = new System.Collections.Generic.List<string>();

            if (showTitle && firstDataRow.Table.Columns.Contains(footerSort.SelectedColumn))
            {
                string titleVal = firstDataRow[footerSort.SelectedColumn]?.ToString();
                if (!string.IsNullOrEmpty(titleVal))
                    parts.Add(titleVal);
            }

            if (showCount)
                parts.Add(count.ToString());

            if (parts.Count == 0) return;

            string displayText = string.Join(", ", parts);

            // Place text in the first visible data column (leftmost column shown in the grid)
            var skipCols = new HashSet<string> { "IsSelected", "RowState", "ElementId", "Count", "TypeName" };
            foreach (System.Data.DataColumn col in footerRow.Table.Columns)
            {
                if (skipCols.Contains(col.ColumnName)) continue;
                if (System.Text.RegularExpressions.Regex.IsMatch(col.ColumnName, @"\s\(\d+\)$")) continue;
                footerRow[col.ColumnName] = displayText;
                break;
            }
        }

        /// <summary>
        /// For any column that isn't a sort/group key column and has differing values across the
        /// grouped rows, replaces the value in <paramref name="newRow"/> with "&lt;Varies&gt;".
        /// </summary>
        private void ApplyVaries(IList<System.Data.DataRow> rows, System.Data.DataRow newRow,
                                 IList<string> groupKeyCols)
        {
            if (rows.Count <= 1) return;

            var skipCols = new HashSet<string> { "IsSelected", "RowState", "ElementId", "Count", "TypeName" };

            foreach (System.Data.DataColumn col in newRow.Table.Columns)
            {
                if (skipCols.Contains(col.ColumnName)) continue;
                if (groupKeyCols != null && groupKeyCols.Contains(col.ColumnName)) continue;
                if (col.DataType != typeof(string)) continue; // only string columns shown as text

                string first = rows[0][col.ColumnName]?.ToString() ?? "";
                bool allSame = rows.All(r => (r[col.ColumnName]?.ToString() ?? "") == first);
                if (!allSame)
                    newRow[col.ColumnName] = "<Varies>";
            }
        }

        private void ScheduleDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            var row = e.Row.Item as System.Data.DataRowView;
            if (row != null)
            {
                var state = row["RowState"]?.ToString();
                if (state == "BlankLine" || state == "FooterLine")
                    e.Cancel = true;
            }
        }

/*
        private void SortCancel_Click(object sender, RoutedEventArgs e)
        {
            // Revert changes (if any were made live without backup? No, we bound directly)
            // Wait, if we bound directly, changes are LIVE.
            // We need to restore backup.
            
            // But we didn't store backup in a field? We did in Sort_Click but it was local.
            // Logic was flawed or reliant on local variable capture? No, WPF ensures modal? No, it was a Popup.
            // Actually the original Sort_Click logic was weird, it cleared then added.
            
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }
*/

        #region Persistence

        // ---------------------------------------------------------------
        // DTOs (shared by both save paths)
        // ---------------------------------------------------------------

        private long GetIdValue(ElementId id)
        {
            return id.Value;
        }

        public class SavedScheduleSort
        {
            public long ScheduleId { get; set; }
            public List<SortItem> Items { get; set; }
            public bool ItemizeEveryInstance { get; set; } = true;
        }

        public class SavedScheduleFilter
        {
            public long ScheduleId { get; set; }
            public List<Models.FilterItem> Items { get; set; }
        }

        // ---------------------------------------------------------------
        // CommitFilterSettings — in-memory only, called on Apply/Clear
        // ---------------------------------------------------------------

        internal void CommitFilterSettings()
        {
            if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
            {
                var list = new ObservableCollection<Models.FilterItem>();
                foreach (var item in FilterCriteria) list.Add(item.Clone());
                _scheduleFilterSettings[_currentScheduleData.ScheduleId] = list;
            }
        }

        // ---------------------------------------------------------------
        // Save — serialise both dicts and raise external event
        // ---------------------------------------------------------------

        /// <summary>
        /// Flush current schedule's sort criteria to the dictionary, then save
        /// all settings to Extensible Storage via ExternalEvent.
        /// Call from any "Apply" commit point (sort/filter windows) as well as on close.
        /// </summary>
        internal void CommitSortSettingsToStorage()
        {
            if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
            {
                var list = new ObservableCollection<SortItem>();
                foreach (var item in SortCriteria) list.Add(item.Clone());
                _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
            }
            SaveSettingsToStorage();
        }

        internal void SaveSettingsToStorage()
        {
            try
            {
                // Flush current schedule's sort into the dict before saving
                if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
                {
                    var list = new ObservableCollection<SortItem>();
                    foreach (var item in SortCriteria) list.Add(item.Clone());
                    _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
                }

                // Build sort JSON
                var sortDtos = new System.Collections.Generic.List<SavedScheduleSort>();
                foreach (var kvp in _scheduleSortSettings)
                {
                    bool itemize = !_scheduleItemizeSettings.TryGetValue(kvp.Key, out bool iv) || iv;
                    sortDtos.Add(new SavedScheduleSort
                    {
                        ScheduleId = GetIdValue(kvp.Key),
                        Items = kvp.Value.ToList(),
                        ItemizeEveryInstance = itemize
                    });
                }
                string sortJson = Newtonsoft.Json.JsonConvert.SerializeObject(sortDtos, Newtonsoft.Json.Formatting.Indented);

                // Build filter JSON
                var filterDtos = new System.Collections.Generic.List<SavedScheduleFilter>();
                foreach (var kvp in _scheduleFilterSettings)
                {
                    filterDtos.Add(new SavedScheduleFilter
                    {
                        ScheduleId = GetIdValue(kvp.Key),
                        Items = kvp.Value.ToList()
                    });
                }
                string filterJson = Newtonsoft.Json.JsonConvert.SerializeObject(filterDtos, Newtonsoft.Json.Formatting.Indented);

                // Raise external event to write to Extensible Storage on Revit's thread
                _saveSettingsHandler.SortSettingsJson   = sortJson;
                _saveSettingsHandler.FilterSettingsJson = filterJson;
                _saveSettingsExternalEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Persistence] SaveSettingsToStorage error: {ex.Message}");
            }
        }

        // ---------------------------------------------------------------
        // Load — read from Extensible Storage (no transaction needed)
        // ---------------------------------------------------------------

        private void LoadSettingsFromStorage(Document doc)
        {
            try
            {
                string username = _uiApplication.Application.Username;
                var (sortJson, filterJson) = Services.ExtensibleStorageService.LoadSettings(doc, username);

                if (!string.IsNullOrWhiteSpace(sortJson))
                {
                    var dtos = Newtonsoft.Json.JsonConvert.DeserializeObject<System.Collections.Generic.List<SavedScheduleSort>>(sortJson);
                    if (dtos != null)
                    {
                        _scheduleSortSettings.Clear();
                        foreach (var dto in dtos)
                        {
                            ElementId eid = new ElementId((long)dto.ScheduleId);
                            _scheduleSortSettings[eid] = new ObservableCollection<SortItem>(dto.Items);
                            _scheduleItemizeSettings[eid] = dto.ItemizeEveryInstance;
                        }
                    }
                }

                if (!string.IsNullOrWhiteSpace(filterJson))
                {
                    var dtos = Newtonsoft.Json.JsonConvert.DeserializeObject<System.Collections.Generic.List<SavedScheduleFilter>>(filterJson);
                    if (dtos != null)
                    {
                        _scheduleFilterSettings.Clear();
                        foreach (var dto in dtos)
                        {
                            ElementId eid = new ElementId((long)dto.ScheduleId);
                            _scheduleFilterSettings[eid] = new ObservableCollection<Models.FilterItem>(dto.Items);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Persistence] LoadSettingsFromStorage error: {ex.Message}");
            }
        }

        #endregion

        private void AddSortLevel_Click(object sender, RoutedEventArgs e)
        {
            SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
        }


/*
        private bool _isSortDragging = false;
        private System.Windows.Point _sortDragStart;

        private void SortHeader_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                var element = sender as IInputElement;
                _isSortDragging = true;
                _sortDragStart = e.GetPosition(this);
                element.CaptureMouse();
                e.Handled = true;
            }
        }
        
        private void SortPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void SortHeader_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isSortDragging && sender is IInputElement element)
            {
                var current = e.GetPosition(this);
                var diff = current - _sortDragStart;
                
                if (SortPopupTransform != null)
                {
                    SortPopupTransform.X += diff.X;
                    SortPopupTransform.Y += diff.Y;
                }
                
                _sortDragStart = current;
            }
        }

        private void SortHeader_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSortDragging)
            {
                _isSortDragging = false;
                (sender as IInputElement)?.ReleaseMouseCapture();
            }
        }
*/



        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Allow clicks on interactive controls or DataGrid rows to function normally
            DependencyObject obj = e.OriginalSource as DependencyObject;
            while (obj != null)
            {
                // If we clicked a control that should handle its own events, or the DataGridRow itself, don't deselect
                if (obj is System.Windows.Controls.Button || 
                    obj is System.Windows.Controls.TextBox || 
                    obj is System.Windows.Controls.CheckBox || 
                    obj is System.Windows.Controls.ComboBox || 
                    obj is System.Windows.Controls.Primitives.ScrollBar || 
                    obj is System.Windows.Controls.Primitives.DataGridColumnHeader || 
                    obj is System.Windows.Controls.DataGridRow || 
                    obj is System.Windows.Controls.Primitives.Thumb)
                {
                    return;
                }

                if (obj is System.Windows.ContentElement contentElement)
                {
                    obj = System.Windows.LogicalTreeHelper.GetParent(contentElement);
                    continue; // Skip rest of loop for this iteration
                }

                if (obj is System.Windows.Media.Visual || obj is System.Windows.Media.Media3D.Visual3D)
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }
                else
                {
                    obj = null; // Stop traversal if not a visual
                }
            }

            // If we are here, we clicked empty space (background, borders, etc.) -> Deselect All
            if (ScheduleDataGrid != null)
            {
                ScheduleDataGrid.UnselectAll();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Validate no sheet number conflicts
                if (!ValidateAllItems())
                {
                    return; // Error message already shown
                }

                // 2. Separate pending creations from pending edits
                var sheetsToCreate = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                var sheetsToEdit = Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList();

                bool hasWork = sheetsToCreate.Count > 0 || sheetsToEdit.Count > 0;
                if (!hasWork)
                {
                    return;
                }

                // 3. Trigger creation handler if needed
                if (sheetsToCreate.Count > 0)
                {
                    _handler.PendingSheetData = sheetsToCreate;
                    _externalEvent.Raise();
                }

                // 4. Trigger edit handler if needed
                if (sheetsToEdit.Count > 0)
                {
                    _editHandler.SheetsToEdit = sheetsToEdit;
                    _editExternalEvent.Raise();
                }

                // Buttons removed - changes are applied immediately
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void DiscardPending_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var pendingCount = Sheets.Count(s => s.HasUnsavedChanges);
                if (pendingCount == 0) return;

                ShowConfirmPopup("Confirm Discard", $"Discard {pendingCount} pending change(s)?", () =>
                {
                    // Remove pending creations
                    var toRemove = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                    foreach (var sheet in toRemove)
                    {
                        sheet.PropertyChanged -= OnItemPropertyChanged; // Unsubscribe
                        Sheets.Remove(sheet);
                        FilteredSheets.Remove(sheet);
                    }

                    // Revert pending edits
                    foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList())
                    {
                        // Reset properties to original? 
                        // Simplified: Just clear dirty flag if that was the only change
                        // Real revert needs original values which we might not have stored easily here
                        // For now, just reset the state
                        sheet.State = SheetItemState.ExistingInRevit;
                    }

                    UpdateButtonStates();
                    ShowPopup("Success", "Discarded all pending changes.");
                });
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void HighlightInModel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<ElementId> elementIds = new List<ElementId>();

                // Check referencing ItemsSource to determine mode
                if (ScheduleDataGrid.ItemsSource is System.Data.DataView dataView)
                {
                    // Schedule Mode (DataTable)
                    foreach (System.Data.DataRowView row in dataView)
                    {
                        // Check IsSelected column
                        if (row.Row.Table.Columns.Contains("IsSelected") && 
                            row["IsSelected"] != DBNull.Value && 
                            Convert.ToBoolean(row["IsSelected"]))
                        {
                            if (row.Row.Table.Columns.Contains("ElementId"))
                            {
                                string idStr = row["ElementId"]?.ToString();
                                
                                // Handle grouped rows (comma-separated IDs)
                                if (!string.IsNullOrEmpty(idStr))
                                {
                                    var parts = idStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach(var part in parts)
                                    {
                                        if (long.TryParse(part.Trim(), out long idLong))
                                        {
#if NET8_0_OR_GREATER
                                            elementIds.Add(new ElementId(idLong));
#else
                                            elementIds.Add(new ElementId(idLong));
#endif
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Sheet Mode (SheetItem collection)
                    // Use ticked sheets (Checkbox = IsSelected)
                    elementIds = Sheets
                        .Where(s => s.IsSelected)
                        .Select(s => s.Id)
                        .ToList();
                }

                if (elementIds.Count == 0)
                {
                    ShowPopup("Selection Required", "Please tick at least one item to highlight.");
                    return;
                }

                // Filter valid IDs
                var validIds = elementIds
                    .Where(id => id != null && id != ElementId.InvalidElementId)
                    .ToList();

                if (validIds.Count == 0)
                {
                    ShowPopup("Error", "Selected items do not have valid Revit Element IDs.");
                    return;
                }

                // Raise External Event
                _highlightInModelHandler.ElementIds = validIds;
                _highlightInModelExternalEvent.Raise();
            }
            catch (Exception ex)
            {
                ShowPopup("Error", $"Failed to highlight elements: {ex.Message}");
            }
        }

        private void OnItemPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (sender is SheetItem sheet && (e.PropertyName == "Name" || e.PropertyName == "SheetNumber"))
            {
                // Mark as edited if it's an existing sheet
                if (sheet.State == SheetItemState.ExistingInRevit &&
                    (sheet.SheetNumber != sheet.OriginalSheetNumber || sheet.Name != sheet.OriginalName))
                {
                    sheet.State = SheetItemState.PendingEdit;
                }

                // If the sheet was reverted to original values, reset state
                if (sheet.State == SheetItemState.PendingEdit &&
                    sheet.SheetNumber == sheet.OriginalSheetNumber && sheet.Name == sheet.OriginalName)
                {
                    sheet.State = SheetItemState.ExistingInRevit;
                }

                ValidateItemNumber(sheet);
                UpdateButtonStates();

                // In Auto mode, immediately push sheet edits to Revit
                if (!_isManualMode && sheet.State == SheetItemState.PendingEdit && !sheet.HasNumberConflict)
                {
                    _editHandler.SheetsToEdit = new List<SheetItem> { sheet };
                    _editExternalEvent.Raise();
                }
            }
        }

        private void OnEditFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Update states to ExistingInRevit
                foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit))
                {
                    sheet.State = SheetItemState.ExistingInRevit;
                    sheet.OriginalSheetNumber = sheet.SheetNumber;
                    sheet.OriginalName = sheet.Name;
                }

                UpdateButtonStates();

                // Show results
                if (fail > 0)
                {
                    ShowPopup("Edit Report", $"Success: {success}\nFailures: {fail}\nError: {errorMsg}");
                }
                else
                {
                    ShowPopup("Success", $"Successfully updated {success} item(s).");
                }
            });
        }

        private void ValidateItemNumber(SheetItem sheet)
        {
            var duplicates = Sheets.Where(s =>
                s != sheet &&
                s.SheetNumber == sheet.SheetNumber &&
                (s.State == SheetItemState.ExistingInRevit || s.State == SheetItemState.PendingEdit)
            ).ToList();

            sheet.HasNumberConflict = duplicates.Any();
            sheet.ValidationError = duplicates.Any()
                ? $"Number '{sheet.SheetNumber}' already exists"
                : null;
        }

        private bool ValidateAllItems()
        {
            bool hasErrors = false;

            foreach (var sheet in Sheets.Where(s => s.HasUnsavedChanges))
            {
                ValidateItemNumber(sheet);
                if (sheet.HasNumberConflict)
                {
                    hasErrors = true;
                }
            }

            if (hasErrors)
            {
                ShowPopup("Validation Error", "Please fix duplicate numbers before applying.");
                return false;
            }

            return true;
        }

        private void UpdateButtonStates()
        {
            // Buttons removed - this method is kept in case it's referenced elsewhere
        }

        private RenameWindow _renameWindow;

        private void Rename_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Determine if we're working with schedule data or sheets
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                bool isScheduleMode = selectedItem != null && selectedItem.Schedule != null && _currentScheduleData != null;

                if (isScheduleMode)
                {
                    // Get selected rows from schedule DataView
                    var view = ScheduleDataGrid.ItemsSource as System.Data.DataView;
                    if (view == null)
                    {
                        ShowPopup("Error", "No schedule data available.");
                        return;
                    }

                    var selectedRows = new List<System.Data.DataRowView>();
                    foreach (System.Data.DataRowView row in view)
                    {
                        var isSelectedValue = row["IsSelected"];
                        bool isSelected = false;
                        
                        if (isSelectedValue is bool b)
                        {
                            isSelected = b;
                        }
                        else if (isSelectedValue != null && isSelectedValue != DBNull.Value)
                        {
                            isSelected = Convert.ToBoolean(isSelectedValue);
                        }
                        
                        if (isSelected)
                        {
                            selectedRows.Add(row);
                        }
                    }

                    if (selectedRows.Count == 0)
                    {
                        ShowPopup("No Rows Selected", "Please tick at least one row to rename.");
                        return;
                    }

                    // Open RenameWindow in schedule mode
                    _renameWindow = new RenameWindow(this, selectedRows, _currentScheduleData);
                    _renameWindow.OnScheduleRenameApply += OnScheduleRenameApply;
                    _renameWindow.ShowDialog();
                }
                else
                {
                    // Sheet mode
                    var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                    if (selectedSheets.Count == 0)
                    {
                        ShowPopup("No Items Selected", "Please select at least one item to rename.");
                        return;
                    }

                    // Open RenameWindow in sheet mode
                    _renameWindow = new RenameWindow(this, selectedSheets);
                    _renameWindow.OnItemRenameApply += OnItemRenameApply;
                    _renameWindow.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void OnScheduleRenameApply(List<ScheduleRenameItem> items)
        {
            _parameterRenameHandler.RenameItems = items;
            _parameterRenameExternalEvent.Raise();
        }

        private void OnItemRenameApply(List<RenamePreviewItem> items, string parameterName)
        {
            bool isSheetNumber = parameterName == "Sheet Number";

            foreach (var item in items)
            {
                if (isSheetNumber)
                {
                    item.Sheet.SheetNumber = item.New;
                }
                else
                {
                    item.Sheet.Name = item.New;
                }
            }
        }



        private void RenameParameter_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateRenamePreview();
        }

        private void RenameText_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateRenamePreview();
        }

        private void UpdateRenamePreview()
        {
            try
            {
                RenamePreviewItems.Clear();

                string findText = RenameFindText?.Text ?? "";
                string replaceText = RenameReplaceText?.Text ?? "";
                string prefix = RenamePrefixText?.Text ?? "";
                string suffix = RenameSuffixText?.Text ?? "";

                var selectedOption = RenameParameter?.SelectedItem as RenameParameterOption;
                if (selectedOption == null) return;

                // This popup is only used for sheet mode now (schedule mode uses RenameWindow)
                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                if (selectedSheets.Count == 0) return;

                bool isSheetNumber = selectedOption.Name == "Sheet Number";

                foreach (var sheet in selectedSheets)
                {
                    string original = isSheetNumber ? sheet.SheetNumber : sheet.Name;
                    string newValue = ApplyRenameTransform(original, findText, replaceText, prefix, suffix);

                    var previewItem = new RenamePreviewItem(sheet, original)
                    {
                        New = newValue
                    };

                    RenamePreviewItems.Add(previewItem);
                }

                RenamePreviewDataGrid.ItemsSource = RenamePreviewItems;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating rename preview: {ex.Message}");
            }
        }


        /// <summary>
        /// Applies find/replace, prefix, and suffix transformations to a value.
        /// </summary>
        private string ApplyRenameTransform(string original, string find, string replace, string prefix, string suffix)
        {
            string newValue = original ?? "";

            // Apply find/replace
            if (!string.IsNullOrEmpty(find))
            {
                newValue = newValue.Replace(find, replace);
            }

            // Apply prefix
            if (!string.IsNullOrEmpty(prefix))
            {
                newValue = prefix + newValue;
            }

            // Apply suffix
            if (!string.IsNullOrEmpty(suffix))
            {
                newValue = newValue + suffix;
            }

            return newValue;
        }


        private void RenameApply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedOption = RenameParameter?.SelectedItem as RenameParameterOption;

                // This popup is only used for sheet mode now (schedule mode uses RenameWindow)
                bool isSheetNumber = selectedOption?.Name == "Sheet Number";

                foreach (var previewItem in RenamePreviewItems)
                {
                    if (previewItem.Original != previewItem.New)
                    {
                        if (isSheetNumber)
                        {
                            previewItem.Sheet.SheetNumber = previewItem.New;
                        }
                        else
                        {
                            previewItem.Sheet.Name = previewItem.New;
                        }
                    }
                }

                // Close popup
                RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }


        /// <summary>
        /// Executes the schedule parameter rename via ExternalEvent.
        /// </summary>
        private void ExecuteScheduleRename(List<ScheduleRenameItem> items)
        {
            _parameterRenameHandler.RenameItems = items;
            _parameterRenameExternalEvent.Raise();
            RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        /// <summary>
        /// Callback when the parameter rename operation completes.
        /// </summary>
        private void OnParameterRenameFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Reload schedule data to show updated values
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                if (selectedItem?.Schedule != null)
                {
                    LoadScheduleData(selectedItem.Schedule);

                    // Restore itemize setting and refresh view
                    bool itemize = true;
                    if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                    {
                        itemize = _scheduleItemizeSettings[selectedItem.Id];
                    }
                    RefreshScheduleView(itemize);
                    ApplyCurrentSortLogic();
                }

                // Show result
                if (fail > 0)
                {
                    ShowPopup("Rename Report", $"Success: {success}\nFailures: {fail}\nError: {errorMsg}");
                }
                else if (success > 0)
                {
                    ShowPopup("Success", $"Successfully renamed {success} parameter value(s).");
                }
            });
        }

        private void OnParameterValueUpdateFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Clear manual mode pending state after apply
                _pendingParameterChanges.Clear();
                ClearAllCellHighlights();

                // Reload schedule data to show updated values from Revit
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                if (selectedItem?.Schedule != null)
                {
                    LoadScheduleData(selectedItem.Schedule);

                    // Restore itemize setting and refresh view
                    bool itemize = true;
                    if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                    {
                        itemize = _scheduleItemizeSettings[selectedItem.Id];
                    }
                    RefreshScheduleView(itemize);
                    ApplyCurrentSortLogic();
                }

                // Show error if failed
                if (fail > 0 || !string.IsNullOrEmpty(errorMsg))
                {
                    string msg = string.IsNullOrEmpty(errorMsg) ? "Operation failed." : errorMsg;
                    if (fail > 0) msg += $"\nFailures: {fail}";
                    if (success > 0) msg += $"\nSuccess: {success}";
                    
                    ShowPopup("Update Report", msg);
                }
                // No popup for pure success to avoid annoying the user on every cell edit
            });
        }


        private void RenameCancel_Click(object sender, RoutedEventArgs e)
        {
            RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void Parameters_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem == null || selectedItem.Schedule == null)
            {
                ShowPopup("No Schedule Selected", "Please select a schedule first.");
                return;
            }

            _parameterLoadHandler.ScheduleId = selectedItem.Id;
            _parameterLoadExternalEvent.Raise();
        }

        private void OnParameterDataLoaded(List<ParameterItem> available, List<ParameterItem> scheduled, string categoryName)
        {
            Dispatcher.Invoke(() =>
            {
                var win = new ParametersWindow(available, scheduled, categoryName);
                win.Owner = this;
                win.OnApply += OnParametersApply;
                win.ShowDialog();
            });
        }

        private void OnParametersApply(List<ParameterItem> newFields)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem == null) return;

            _scheduleFieldsHandler.ScheduleId = selectedItem.Id;
            _scheduleFieldsHandler.NewFields = newFields;
            _scheduleFieldsExternalEvent.Raise();
        }

        private void OnScheduleFieldsUpdateFinished(int count, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    ShowPopup("Error Updating Fields", errorMsg);
                }
                else
                {
                    // Refresh data
                    var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                    if (selectedItem?.Schedule != null)
                    {
                        LoadScheduleData(selectedItem.Schedule);
                        
                        // Restore itemize setting and refresh view
                        bool itemize = true;
                        if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                        {
                            itemize = _scheduleItemizeSettings[selectedItem.Id];
                        }
                        RefreshScheduleView(itemize);
                        ApplyCurrentSortLogic();
                    }
                    ShowPopup("Success", "Schedule fields updated successfully.");
                }
            });
        }

        private void ParametersClose_Click(object sender, RoutedEventArgs e)
        {
            ParametersPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void ParametersPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        private void RenamePopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Optionally close on background click
            // RenamePopupOverlay.Visibility = Visibility.Collapsed;
        }

        private void ScheduleDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // Clear cell selection
                if (ScheduleDataGrid != null && ScheduleDataGrid.SelectedCells.Count > 0)
                {
                    ScheduleDataGrid.SelectedCells.Clear();
                    ScheduleDataGrid.CurrentCell = new DataGridCellInfo();
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Space)
            {
                var dataGrid = sender as DataGrid;
                if (dataGrid?.SelectedItems != null && dataGrid.SelectedItems.Count > 0)
                {
                    // Check if items are SheetItem (Duplicate Sheets mode)
                    if (dataGrid.SelectedItems[0] is SheetItem)
                    {
                        e.Handled = true;
                        
                        var selectedSheets = dataGrid.SelectedItems.Cast<SheetItem>().ToList();
                        bool newState = !selectedSheets.First().IsSelected;
                        
                        foreach (var sheet in selectedSheets)
                        {
                            sheet.IsSelected = newState;
                        }
                        
                        dataGrid.Items.Refresh();
                    }
                }
            }
        }

        private void FillHandle_DragDelta(object sender, DragDeltaEventArgs e)
        {
            var thumb = sender as Thumb;
            if (thumb != null)
            {
                // Find the target cell under the mouse
                var point = Mouse.GetPosition(ScheduleDataGrid);
                var hitResult = VisualTreeHelper.HitTest(ScheduleDataGrid, point);
                if (hitResult == null) return;
                
                DataGridCell targetCell = FindVisualParent<DataGridCell>(hitResult.VisualHit);
                if (targetCell == null) return;

                // Stop if we are over the same cell as the start
                // Or checking if selection actually needs change could be optimization
                
                // Get Anchor (Current Cell)
                var anchorInfo = ScheduleDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                var anchorItem = anchorInfo.Item;
                var anchorCol = anchorInfo.Column;
                
                // Resolve Indices
                int anchorRowIdx = ScheduleDataGrid.Items.IndexOf(anchorItem);
                int anchorColIdx = anchorCol.DisplayIndex;
                
                int targetRowIdx = ScheduleDataGrid.Items.IndexOf(targetCell.DataContext);
                int targetColIdx = targetCell.Column.DisplayIndex;
                
                if (anchorRowIdx < 0 || targetRowIdx < 0) return;
                
                // Determine Range
                int minRow = Math.Min(anchorRowIdx, targetRowIdx);
                int maxRow = Math.Max(anchorRowIdx, targetRowIdx);
                int minCol = Math.Min(anchorColIdx, targetColIdx);
                int maxCol = Math.Max(anchorColIdx, targetColIdx);
                
                ScheduleDataGrid.SelectedCells.Clear();
                
                // Select Range
                for (int r = minRow; r <= maxRow; r++)
                {
                    var item = ScheduleDataGrid.Items[r];
                    for (int c = minCol; c <= maxCol; c++)
                    {
                        var col = ScheduleDataGrid.Columns[c];
                        ScheduleDataGrid.SelectedCells.Add(new DataGridCellInfo(item, col));
                    }
                }
                
                UpdateSelectionAdorner(); // Force update during drag
            }
        }

        private enum SmartFillMode { Copy, Series }

        private void FillHandle_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                PerformAutoFill(SmartFillMode.Copy);
            }
            else
            {
                PerformAutoFill(SmartFillMode.Series);
            }
        }

        private void PerformAutoFill(SmartFillMode mode)
        {
            try
            {
                if (ScheduleDataGrid.SelectedCells.Count < 2) return;

                // 1. Get Anchor Value (CurrentCell)
                var anchorInfo = ScheduleDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                var anchorRow = anchorInfo.Item as System.Data.DataRowView;
                if (anchorRow == null || _currentScheduleData == null) return;

                // Safely get column name or identify if it's the CheckBox column
                string colName = "";
                if (anchorInfo.Column.Header is CheckBox)
                {
                    colName = "IsSelected";
                }
                else
                {
                    colName = anchorInfo.Column.Header?.ToString();
                }

                if (string.IsNullOrEmpty(colName)) return;

                // Special handling for IsSelected (Checkbox)
                if (colName == "IsSelected")
                {
                    PerformCheckboxAutoFill();
                    return;
                }

                // Check if Anchor passes validation (e.g. not ReadOnly)
                // Actually, Anchor value is valid. We need to check Targets.
                
                string sourceValue = anchorRow[colName]?.ToString() ?? "";
                
                int anchorRowIndex = ScheduleDataGrid.Items.IndexOf(anchorRow);

                // 2. Prepare Updates
                var updates = new List<ParameterUpdateInfo>();
                bool hasTypeParameters = false;
                var affectedTypeParams = new HashSet<string>();

                foreach (var cellInfo in ScheduleDataGrid.SelectedCells)
                {
                    // Skip if it's the anchor itself (Reference equality check on Item and Column)
                    if (cellInfo.Item == anchorInfo.Item && cellInfo.Column == anchorInfo.Column) continue;

                    var targetRow = cellInfo.Item as System.Data.DataRowView;
                    if (targetRow == null) continue;

                    var targetCol = cellInfo.Column;
                    string targetColName = targetCol.Header?.ToString();
                    
                    // We only auto-fill within the SAME column usually?
                    // Excel allows filling across columns if dragging corner?
                    // Typically dragging corner fills the selected range.
                    // If I select a 2x2 range, and drag...
                    // But here we are dragging the Grip of the Anchor.
                    // The selection expands.
                    // Typically we fill with the Anchor's value into ALL selected cells.
                    // But usually you only fill into same-column cells unless standard Copy/Paste.
                    // Let's assume SAME COLUMN as Anchor for safety?
                    // Or follow Selection?
                    // If I drag right, I copy to right.
                    // If I drag down, I copy down.
                    // Let's allow all selected cells.
                    
                    if (string.IsNullOrEmpty(targetColName)) continue;
                    
                    // Check if Column is ReadOnly
                    if (targetCol.IsReadOnly) continue;

                    // Get IDs
                    if (!targetRow.Row.Table.Columns.Contains("ElementId")) continue;
                    string elementIdStr = targetRow["ElementId"]?.ToString();
                    if (string.IsNullOrEmpty(elementIdStr)) continue;

                    // Find Parameter ID
                    int colIndex = _currentScheduleData.Columns.IndexOf(targetColName);
                    if (colIndex < 0 || colIndex >= _currentScheduleData.ParameterIds.Count) continue;
                    ElementId paramId = _currentScheduleData.ParameterIds[colIndex];

                    // Check Type Parameter
                    bool isType = _currentScheduleData.IsTypeParameter.ContainsKey(targetColName) && 
                                  _currentScheduleData.IsTypeParameter[targetColName];

                    if (isType) 
                    {
                        hasTypeParameters = true;
                        affectedTypeParams.Add(targetColName);
                    }

                    // Determine New Value
                    string newValue;
                    if (mode == SmartFillMode.Series)
                    {
                        int targetRowIndex = ScheduleDataGrid.Items.IndexOf(targetRow);
                        int offset = targetRowIndex - anchorRowIndex;
                        newValue = GetSequentialValue(sourceValue, offset);
                    }
                    else
                    {
                        newValue = sourceValue;
                    }

                    // Check if value actually changes (Optimization)
                    string currentVal = targetRow[targetColName]?.ToString() ?? "";
                    if (currentVal == newValue) continue;

                    updates.Add(new ParameterUpdateInfo
                    {
                        ElementIdStr = elementIdStr,
                        ParameterId = paramId,
                        NewValue = newValue,
                        ColumnName = targetColName,
                        IsTypeParameter = isType,
                        Row = targetRow
                    });
                }

                if (updates.Count == 0) return;

                // 3. Confirm Type Parameters
                if (hasTypeParameters)
                {
                    string paramNames = string.Join(", ", affectedTypeParams);
                    ShowConfirmationPopup(
                        "Batch Update Type Parameters",
                        $"You are about to update Type Parameters ({paramNames}).\nThis will affect ALL elements of the corresponding types.\n\nProceed with Auto-Fill?",
                        () => ExecuteBatchUpdates(updates)
                    );
                    return; // Wait for user confirmation — ExecuteBatchUpdates is called inside the callback
                }
                if (updates.Count > 0)
                {
                    ExecuteBatchUpdates(updates);
                }
            }
            catch (Exception ex)
            {
                ShowPopup("AutoFill Error", ex.Message);
            }
        }

        private void PerformCheckboxAutoFill()
        {
            try
            {
                var anchorInfo = ScheduleDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                bool anchorValue = false;

                // determining anchor value
                if (anchorInfo.Item is SheetItem sheet)
                {
                    anchorValue = sheet.IsSelected;
                }
                else if (anchorInfo.Item is System.Data.DataRowView row)
                {
                    if (row.Row.Table.Columns.Contains("IsSelected") && row["IsSelected"] != DBNull.Value)
                    {
                        anchorValue = Convert.ToBoolean(row["IsSelected"]);
                    }
                }

                int updatedCount = 0;

                foreach (var cellInfo in ScheduleDataGrid.SelectedCells)
                {
                    // Skip if it's the anchor itself
                    if (cellInfo.Item == anchorInfo.Item && cellInfo.Column == anchorInfo.Column) continue;

                    // Apply value
                    if (cellInfo.Item is SheetItem targetSheet)
                    {
                        if (targetSheet.IsSelected != anchorValue)
                        {
                            targetSheet.IsSelected = anchorValue;
                            updatedCount++;
                        }
                    }
                    else if (cellInfo.Item is System.Data.DataRowView targetRow)
                    {
                        if (targetRow.Row.Table.Columns.Contains("IsSelected"))
                        {
                            bool currentValue = targetRow["IsSelected"] != DBNull.Value && Convert.ToBoolean(targetRow["IsSelected"]);
                            if (currentValue != anchorValue)
                            {
                                targetRow["IsSelected"] = anchorValue;
                                updatedCount++;
                            }
                        }
                    }
                }

                // If items are in DataTable mode, we might need to notify UI or just rely on Binding.
                // Normally DataRowView changes reflect in UI if bound properly.
            }
            catch (Exception ex)
            {
                ShowPopup("Selection Fill Error", ex.Message);
            }
        }
        
        private void UpdateSelectionAdorner()
        {
            if (SelectionCanvas == null) return;

            if (ScheduleDataGrid.SelectedCells.Count == 0)
            {
                SelectionBox.Visibility = System.Windows.Visibility.Collapsed;
                FillHandle.Visibility = System.Windows.Visibility.Collapsed;
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
                return;
            }

            // Calculate Union of Visible Cells
            System.Windows.Rect unionRect = System.Windows.Rect.Empty;
            bool hasVisibleCells = false;

            foreach (var cellInfo in ScheduleDataGrid.SelectedCells)
            {
                var row = ScheduleDataGrid.ItemContainerGenerator.ContainerFromItem(cellInfo.Item) as DataGridRow;
                if (row == null) continue; // Row not loaded/visible

                var col = cellInfo.Column;
                var cellContent = col.GetCellContent(row);
                if (cellContent == null) continue; // Cell not loaded

                var cell = cellContent.Parent as FrameworkElement;
                if (cell == null) continue;

                // Create Rect for this cell
                System.Windows.Point p = cell.TranslatePoint(new System.Windows.Point(0, 0), SelectionCanvas);
                System.Windows.Rect cellRect = new System.Windows.Rect(p, new System.Windows.Size(cell.ActualWidth, cell.ActualHeight));

                if (unionRect == System.Windows.Rect.Empty)
                    unionRect = cellRect;
                else
                    unionRect.Union(cellRect);

                hasVisibleCells = true;
            }

            if (!hasVisibleCells)
            {
                SelectionBox.Visibility = System.Windows.Visibility.Collapsed;
                FillHandle.Visibility = System.Windows.Visibility.Collapsed;
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
                return;
            }

            // Update UI
            SelectionBox.Width = unionRect.Width;
            SelectionBox.Height = unionRect.Height;
            Canvas.SetLeft(SelectionBox, unionRect.Left);
            Canvas.SetTop(SelectionBox, unionRect.Top);
            SelectionBox.Visibility = System.Windows.Visibility.Visible;

            // Fill Handle (Bottom Right)
            Canvas.SetLeft(FillHandle, unionRect.Right - 6); 
            Canvas.SetTop(FillHandle, unionRect.Bottom - 6);
            FillHandle.Visibility = System.Windows.Visibility.Visible;
            
            // Copy Indicator
            if (IsCopyMode)
            {
                FillIndicator.Visibility = System.Windows.Visibility.Visible;
                // Position at Top Right of the square (FillHandle)
                // FillHandle is 6x6, placed at (Right-6, Bottom-6).
                // So FillHandle Top-Right is (Right, Bottom-6).
                // We place the "+" so its bottom-left is roughly there?
                // Or center it?
                // User said: "top right of the square in the corner" and "2x its size"
                
                // Let's place it slightly offset to look "top right"
                Canvas.SetLeft(FillIndicator, unionRect.Right);
                Canvas.SetTop(FillIndicator, unionRect.Bottom - 20); // Moved up to sit above/corner
            }
            else
            {
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
            }
        }



        private string GetSequentialValue(string original, int offset)
        {
            if (string.IsNullOrEmpty(original)) return original;

            // Find the last sequence of digits
            var matches = System.Text.RegularExpressions.Regex.Matches(original, @"\d+");
            if (matches.Count > 0)
            {
                var lastMatch = matches[matches.Count - 1];
                long number;
                if (long.TryParse(lastMatch.Value, out number))
                {
                    long newNumber = number + offset;
                    
                    // Format preservation: if original was "001", try to keep "002" length
                    // If "1", keep "2". 
                    // Use zero padding if original had it.
                    string format = new string('0', lastMatch.Length);
                    // Check if actually zero-padded
                    bool isZeroPadded = lastMatch.Value.StartsWith("0") && lastMatch.Value.Length > 1;
                    
                    string newNumStr = isZeroPadded ? newNumber.ToString(format) : newNumber.ToString();
                    
                    string prefix = original.Substring(0, lastMatch.Index);
                    string suffix = original.Substring(lastMatch.Index + lastMatch.Length);
                    
                    return prefix + newNumStr + suffix;
                }
            }
            
            return original;
        }

        private void ExecuteBatchUpdates(List<ParameterUpdateInfo> updates)
        {
            if (updates == null || updates.Count == 0) return;

            if (_isManualMode)
            {
                // Queue all updates as pending changes instead of executing immediately
                foreach (var update in updates)
                {
                    int rowIndex = -1;
                    if (update.Row != null)
                    {
                        rowIndex = update.Row.Row.Table.Rows.IndexOf(update.Row.Row);
                    }

                    var existing = _pendingParameterChanges.FirstOrDefault(p =>
                        p.ElementIdStr == update.ElementIdStr &&
                        p.ParameterId.Value == update.ParameterId.Value &&
                        p.ColumnName == update.ColumnName);

                    if (existing != null)
                    {
                        existing.NewValue = update.NewValue;
                    }
                    else
                    {
                        string oldValue = "";
                        if (update.Row != null && update.Row.Row.Table.Columns.Contains(update.ColumnName))
                        {
                            oldValue = update.Row[update.ColumnName]?.ToString() ?? "";
                        }

                        _pendingParameterChanges.Add(new PendingParameterChange
                        {
                            ElementIdStr = update.ElementIdStr,
                            ParameterId = update.ParameterId,
                            NewValue = update.NewValue,
                            OldValue = oldValue,
                            ColumnName = update.ColumnName,
                            RowIndex = rowIndex
                        });
                    }

                    if (update.Row != null)
                    {
                        HighlightPendingCell(update.Row, update.ColumnName);
                    }
                }
                return;
            }

            var batchData = new List<ParameterBatchData>();
            foreach (var update in updates)
            {
                batchData.Add(new ParameterBatchData
                {
                    ElementIdStr = update.ElementIdStr,
                    ParameterId = update.ParameterId,
                    Value = update.NewValue
                });
            }

            _parameterValueUpdateHandler.IsBatchMode = true;
            _parameterValueUpdateHandler.BatchData = batchData;
            _parameterValueUpdateExternalEvent.Raise();
        }

        private class ParameterUpdateInfo
        {
            public string ElementIdStr { get; set; }
            public ElementId ParameterId { get; set; }
            public string NewValue { get; set; }
            public string ColumnName { get; set; }
            public bool IsTypeParameter { get; set; }
            public System.Data.DataRowView Row { get; set; }
        }

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            var parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent) return parent;
            return FindVisualParent<T>(parentObject);
        }

        private void DataGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var dep = e.OriginalSource as DependencyObject;
            while (dep != null && !(dep is DataGridCell) && !(dep is DataGridColumnHeader))
            {
                if (dep is CheckBox)
                {
                    e.Handled = true;
                    var checkbox = dep as CheckBox;
                    checkbox.IsChecked = !checkbox.IsChecked;
                    return;
                }
                dep = VisualTreeHelper.GetParent(dep);
            }
        }



        private void ScheduleDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectionAdorner();
        }

        private void ScheduleDataGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var scrollViewer = GetScrollViewer(ScheduleDataGrid);
            if (scrollViewer == null) return;

            // Handle horizontal scrolling with Shift + MouseWheel OR horizontal wheel
            if (Keyboard.Modifiers == ModifierKeys.Shift || e.Delta == 0)
            {
                e.Handled = true;
                
                // For horizontal wheel, Delta is 0 but we need to check the actual MouseDevice
                // In WPF, horizontal scrolling is typically handled via MouseWheel with Shift modifier
                // However, if we detect Shift is pressed, we definitely want horizontal scroll
                if (Keyboard.Modifiers == ModifierKeys.Shift)
                {
                    if (e.Delta > 0)
                        scrollViewer.LineLeft();
                    else
                        scrollViewer.LineRight();
                }
            }
        }


        private static ScrollViewer GetScrollViewer(DependencyObject depObj)
        {
            if (depObj is ScrollViewer) return depObj as ScrollViewer;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);
                var result = GetScrollViewer(child);
                if (result != null) return result;
            }
            return null;
        }

        private void ScheduleSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = (sender as System.Windows.Controls.TextBox)?.Text?.ToLowerInvariant() ?? "";

            // Schedule mode — filter the DataView via RowFilter
            if (_currentScheduleData != null && ScheduleDataGrid.ItemsSource is System.Data.DataView dataView)
            {
                if (string.IsNullOrEmpty(searchText))
                {
                    dataView.RowFilter = string.Empty;
                }
                else
                {
                    var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "TypeName", "Count" };
                    var filterableCols = dataView.Table.Columns
                        .Cast<System.Data.DataColumn>()
                        .Where(c => !skipColumns.Contains(c.ColumnName) && c.DataType == typeof(string))
                        .Select(c => $"CONVERT([{c.ColumnName}], System.String) LIKE '%{searchText.Replace("'", "''")}%'")
                        .ToList();

                    // Always let blank separator rows through
                    dataView.RowFilter = filterableCols.Count > 0
                        ? $"(RowState = 'BlankLine') OR (RowState = 'FooterLine') OR ({string.Join(" OR ", filterableCols)})"
                        : string.Empty;
                }
                return;
            }

            // Sheet mode — filter FilteredSheets
            FilteredSheets.Clear();
            foreach (var sheet in Sheets.Where(s => 
                string.IsNullOrEmpty(searchText) || 
                s.Name.ToLowerInvariant().Contains(searchText) ||
                s.SheetNumber.ToLowerInvariant().Contains(searchText)))
            {
                FilteredSheets.Add(sheet);
            }
        }

        private void ClearScheduleSearch_Click(object sender, RoutedEventArgs e)
        {
            ScheduleSearchBox.Clear();
        }

        private void RowCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is SheetItem clickedItem)
            {
                e.Handled = true;
                bool newState = !(checkBox.IsChecked ?? false);
                checkBox.IsChecked = newState;

                if (ScheduleDataGrid.SelectedItems.Contains(clickedItem))
                {
                    foreach (SheetItem item in ScheduleDataGrid.SelectedItems)
                    {
                        if (item != clickedItem)
                        {
                            item.IsSelected = newState;
                        }
                    }
                }
            }
        }

        private void SelectAll_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in FilteredSheets)
            {
                sheet.IsSelected = true;
            }
        }

        private void SelectAll_Unchecked(object sender, RoutedEventArgs e)
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
                    ? "pack://application:,,,/ProSchedules;component/UI/Themes/DarkTheme.xaml" 
                    : "pack://application:,,,/ProSchedules;component/UI/Themes/LightTheme.xaml", UriKind.Absolute);
                
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
            }

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

        #region Manual Update Mode

        private void ManualModeToggle_Click(object sender, RoutedEventArgs e)
        {
            bool wasManualMode = _isManualMode;
            _isManualMode = ManualModeToggle.IsChecked == true;
            UpdateManualModeUI();
            SaveManualModeState();

            // If switching from Manual back to Auto with pending changes, prompt
            if (wasManualMode && !_isManualMode && HasPendingManualChanges())
            {
                ShowConfirmationPopup("Switch to Auto Mode",
                    "You have pending changes. Apply them before switching to Auto mode?",
                    () => { ApplyAllPendingChanges(); },
                    () => { CancelAllPendingChanges(); });
            }
        }

        private void UpdateManualModeUI()
        {
            if (_isManualMode)
            {
                ManualApplyButton.IsEnabled = true;
                ManualApplyButton.Opacity = 1.0;
                ManualCancelButton.IsEnabled = true;
                ManualCancelButton.Opacity = 1.0;
            }
            else
            {
                ManualApplyButton.IsEnabled = false;
                ManualApplyButton.Opacity = 0.3;
                ManualCancelButton.IsEnabled = false;
                ManualCancelButton.Opacity = 0.3;
            }
        }

        private void LoadManualModeState()
        {
            try
            {
                var config = LoadConfig();
                if (TryGetBool(config, "IsManualMode", out var isManual))
                {
                    _isManualMode = isManual;
                }
            }
            catch (Exception)
            {
            }

            if (ManualModeToggle != null)
            {
                ManualModeToggle.IsChecked = _isManualMode;
            }
            UpdateManualModeUI();
        }

        private void SaveManualModeState()
        {
            try
            {
                var config = LoadConfig();
                config["IsManualMode"] = _isManualMode;
                SaveConfig(config);
            }
            catch (Exception)
            {
            }
        }

        private bool HasPendingManualChanges()
        {
            return _pendingParameterChanges.Count > 0 ||
                   Sheets.Any(s => s.HasUnsavedChanges);
        }

        private void ManualApply_Click(object sender, RoutedEventArgs e)
        {
            ApplyAllPendingChanges();
        }

        private void ApplyAllPendingChanges()
        {
            // 1. Apply pending schedule parameter changes
            if (_pendingParameterChanges.Count > 0)
            {
                var batchData = _pendingParameterChanges.Select(p => new ExternalEvents.ParameterBatchData
                {
                    ElementIdStr = p.ElementIdStr,
                    ParameterId = p.ParameterId,
                    Value = p.NewValue
                }).ToList();

                _parameterValueUpdateHandler.IsBatchMode = true;
                _parameterValueUpdateHandler.BatchData = batchData;
                _parameterValueUpdateExternalEvent.Raise();
                // _pendingParameterChanges will be cleared in OnParameterValueUpdateFinished
            }

            // 2. Apply pending sheet edits
            var sheetsToEdit = Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList();
            if (sheetsToEdit.Count > 0)
            {
                _editHandler.SheetsToEdit = sheetsToEdit;
                _editExternalEvent.Raise();
            }

            // 3. Apply pending sheet creations
            var sheetsToCreate = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
            if (sheetsToCreate.Count > 0)
            {
                _handler.PendingSheetData = sheetsToCreate;
                _externalEvent.Raise();
            }
        }

        private void ManualCancel_Click(object sender, RoutedEventArgs e)
        {
            if (!HasPendingManualChanges()) return;

            ShowConfirmationPopup("Confirm Cancel", "Discard all pending changes?",
                () => { CancelAllPendingChanges(); });
        }

        private void CancelAllPendingChanges()
        {
            // 1. Revert schedule parameter changes in the DataTable UI
            var view = ScheduleDataGrid.ItemsSource as System.Data.DataView;
            if (view != null)
            {
                foreach (var change in _pendingParameterChanges)
                {
                    if (change.RowIndex >= 0 && change.RowIndex < view.Table.Rows.Count)
                    {
                        var row = view.Table.Rows[change.RowIndex];
                        if (view.Table.Columns.Contains(change.ColumnName))
                        {
                            row[change.ColumnName] = change.OldValue;
                        }
                    }
                }
            }
            _pendingParameterChanges.Clear();
            ClearAllCellHighlights();

            // 2. Revert pending sheet edits
            foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList())
            {
                sheet.SheetNumber = sheet.OriginalSheetNumber;
                sheet.Name = sheet.OriginalName;
                sheet.State = SheetItemState.ExistingInRevit;
            }

            // 3. Remove pending sheet creations
            var toRemove = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
            foreach (var sheet in toRemove)
            {
                sheet.PropertyChanged -= OnItemPropertyChanged;
                Sheets.Remove(sheet);
                FilteredSheets.Remove(sheet);
            }

            UpdateButtonStates();
        }

        private void HighlightPendingCell(System.Data.DataRowView row, string columnName)
        {
            int rowIndex = row.Row.Table.Rows.IndexOf(row.Row);
            _highlightedCells.Add($"{rowIndex}:{columnName}");
            ApplyCellHighlights();
        }

        private void ClearCellHighlight(System.Data.DataRowView row, string columnName)
        {
            int rowIndex = row.Row.Table.Rows.IndexOf(row.Row);
            _highlightedCells.Remove($"{rowIndex}:{columnName}");
            ApplyCellHighlights();
        }

        private void ClearAllCellHighlights()
        {
            _highlightedCells.Clear();
            ApplyCellHighlights();
        }

        private void ApplyCellHighlights()
        {
            // Force the DataGrid to re-render rows so LoadingRow fires
            if (ScheduleDataGrid.ItemsSource is System.Data.DataView)
            {
                ScheduleDataGrid.Items.Refresh();
                UpdateSelectionAdorner();
            }
        }

        private void ScheduleDataGrid_LoadingRow_ManualMode(object sender, DataGridRowEventArgs e)
        {
            if (!_isManualMode || _highlightedCells.Count == 0)
            {
                // Clear any previously applied highlights when not in manual mode
                ClearRowCellHighlights(e.Row);
                return;
            }

            var row = e.Row.Item as System.Data.DataRowView;
            if (row == null) return;

            int rowIndex = row.Row.Table.Rows.IndexOf(row.Row);

            // We need to defer this until after the row is rendered
            e.Row.Loaded -= Row_ApplyHighlights;
            e.Row.Loaded += Row_ApplyHighlights;
        }

        private void Row_ApplyHighlights(object sender, RoutedEventArgs e)
        {
            var dataGridRow = sender as DataGridRow;
            if (dataGridRow == null) return;
            dataGridRow.Loaded -= Row_ApplyHighlights;

            var row = dataGridRow.Item as System.Data.DataRowView;
            if (row == null) return;

            int rowIndex = row.Row.Table.Rows.IndexOf(row.Row);

            foreach (var column in ScheduleDataGrid.Columns)
            {
                string colName = column.Header?.ToString();
                if (colName == null) continue;

                var cellContent = column.GetCellContent(dataGridRow);
                if (cellContent == null) continue;
                var cell = cellContent.Parent as DataGridCell;
                if (cell == null) continue;

                if (_isManualMode && _highlightedCells.Contains($"{rowIndex}:{colName}"))
                {
                    cell.Background = (Brush)FindResource("PendingChangeBrush");
                }
                else
                {
                    cell.ClearValue(DataGridCell.BackgroundProperty);
                }
            }
        }

        private void ClearRowCellHighlights(DataGridRow dataGridRow)
        {
            if (dataGridRow == null) return;

            foreach (var column in ScheduleDataGrid.Columns)
            {
                var cellContent = column.GetCellContent(dataGridRow);
                if (cellContent == null) continue;
                var cell = cellContent.Parent as DataGridCell;
                if (cell == null) continue;
                cell.ClearValue(DataGridCell.BackgroundProperty);
            }
        }

        #endregion

        #region Custom Popup



        public void ShowPopup(string title, string message, Action onCloseAction = null)
        {
            PopupTitle.Text = title;
            PopupMessage.Text = message;
            _onPopupClose = onCloseAction;
            MainContentGrid.IsEnabled = false;
            PopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ClosePopup()
        {
            MainContentGrid.IsEnabled = true;
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
        }



        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                IsCopyMode = true;
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                IsCopyMode = false;
            }
        }
        
        public void ShowConfirmationPopup(string title, string message, Action onConfirmAction, Action onCancelAction = null)
        {
            ConfirmationTitle.Text = title;
            ConfirmationMessage.Text = message;
            _onConfirmAction = onConfirmAction;
            _onCancelAction = onCancelAction;
            MainContentGrid.IsEnabled = false;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ConfirmationOK_Click(object sender, RoutedEventArgs e)
        {
            MainContentGrid.IsEnabled = true;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onConfirmAction != null)
            {
                var action = _onConfirmAction;
                _onConfirmAction = null;
                _onCancelAction = null;
                action.Invoke();
            }
        }

        private void ConfirmationCancel_Click(object sender, RoutedEventArgs e)
        {
            MainContentGrid.IsEnabled = true;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onCancelAction != null)
            {
                var action = _onCancelAction;
                _onCancelAction = null;
                _onConfirmAction = null;
                action.Invoke();
            }
            else
            {
                _onConfirmAction = null;
            }
        }

        private void ShowConfirmPopup(string title, string message, Action onConfirmAction, string confirmButtonText = "Discard")
        {
            ConfirmPopupTitle.Text = title;
            ConfirmPopupMessage.Text = message;
            _onConfirmAction = onConfirmAction;
            ConfirmActionButton.Content = confirmButtonText;
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }


        private void CloseConfirmPopup()
        {
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            _onConfirmAction = null;
        }

        private void ConfirmDiscard_Click(object sender, RoutedEventArgs e)
        {
            var action = _onConfirmAction;
            CloseConfirmPopup();
            action?.Invoke();
        }

        private void ConfirmCancel_Click(object sender, RoutedEventArgs e)
        {
            CloseConfirmPopup();
        }

        private void ConfirmPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        #endregion

        #region Window chrome / resize handlers

        private void TitleBar_Loaded(object sender, RoutedEventArgs e)
        {
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


        #region Window Startup

        private void DeferWindowShow()
        {
            Opacity = 0;
            Loaded += ProSchedulesWindow_Loaded;
        }

        private void ProSchedulesWindow_Loaded(object sender, RoutedEventArgs e)
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

                // Migrate old config keys from DuplicateSheetsWindow.* to ProSchedulesWindow.*
                MigrateConfigKey(config, "DuplicateSheetsWindow.Left", WindowLeftKey);
                MigrateConfigKey(config, "DuplicateSheetsWindow.Top", WindowTopKey);
                MigrateConfigKey(config, "DuplicateSheetsWindow.Width", WindowWidthKey);
                MigrateConfigKey(config, "DuplicateSheetsWindow.Height", WindowHeightKey);

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

        /// <summary>
        /// Migrates a config key from an old name to a new name, removing the old key.
        /// </summary>
        private void MigrateConfigKey(Dictionary<string, object> config, string oldKey, string newKey)
        {
            if (!config.ContainsKey(newKey) && config.ContainsKey(oldKey))
            {
                config[newKey] = config[oldKey];
                config.Remove(oldKey);
                try { SaveConfig(config); } catch { }
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

        private void SaveScheduleName(string scheduleName)
        {
            try
            {
                string folder = GetProjectSettingsFolder();
                string file = System.IO.Path.Combine(folder, "last_schedule.txt");
                System.IO.File.WriteAllText(file, scheduleName ?? "");
                _lastSelectedScheduleName = scheduleName;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving schedule name: {ex.Message}");
            }
        }

        private string GetSavedScheduleName()
        {
            try
            {
                string folder = GetProjectSettingsFolder();
                string file = System.IO.Path.Combine(folder, "last_schedule.txt");
                if (System.IO.File.Exists(file))
                {
                    string scheduleName = System.IO.File.ReadAllText(file);
                    _lastSelectedScheduleName = scheduleName;
                    return scheduleName;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading schedule name: {ex.Message}");
            }

            return null;
        }


        #endregion

        #region Helpers

        private string GetProjectSettingsFolder()
        {
            try
            {
                string docTitle = "Default";
                if (_uiApplication?.ActiveUIDocument?.Document != null)
                {
                    docTitle = _uiApplication.ActiveUIDocument.Document.Title;
                    // Remove extension if present
                    if (docTitle.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                    {
                        docTitle = docTitle.Substring(0, docTitle.Length - 4);
                    }
                }

                // Use Roaming AppData for user settings: %APPDATA%\RK Tools\ProSchedules\{ProjectName}
                string folder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "RK Tools", "ProSchedules", docTitle);

                if (!System.IO.Directory.Exists(folder))
                {
                    System.IO.Directory.CreateDirectory(folder);
                }
                
                return folder;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting settings folder: {ex.Message}");
                // Fallback
                return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RK Tools", "ProSchedules", "Default");
            }
        }

        #endregion

    }
}
