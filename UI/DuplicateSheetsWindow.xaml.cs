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
        public bool ShowFooter { get; set; }
        public bool ShowBlankLine { get; set; }

        public SortItem Clone()
        {
            return new SortItem
            {
                SelectedColumn = this.SelectedColumn,
                IsAscending = this.IsAscending,
                ShowHeader = this.ShowHeader,
                ShowFooter = this.ShowFooter,
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

    public partial class DuplicateSheetsWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\ProSchedules\config.json";
        private const string WindowLeftKey = "DuplicateSheetsWindow.Left";
        private const string WindowTopKey = "DuplicateSheetsWindow.Top";
        private const string WindowWidthKey = "DuplicateSheetsWindow.Width";
        private const string WindowHeightKey = "DuplicateSheetsWindow.Height";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Revit state / UI state

        private List<SheetItem> _allSheets;
        private Action _onPopupClose;
        private Action _onConfirmAction;
        private Action _onDeleteConfirmAction;
        private ExternalEvent _externalEvent;
        private ExternalEvents.SheetDuplicationHandler _handler;
        private ExternalEvent _editExternalEvent;
        private ExternalEvents.SheetEditHandler _editHandler;
        private ExternalEvent _deleteExternalEvent;
        private ExternalEvents.SheetDeleteHandler _deleteHandler;
        private List<SheetItem> _pendingDeleteItems = new List<SheetItem>();
        private int _pendingDeleteLocalCount;

        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isDataLoaded;
        private UIApplication _uiApplication;
        private RevitService _revitService;

        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<SheetItem> FilteredSheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<RenamePreviewItem> RenamePreviewItems { get; set; } = new ObservableCollection<RenamePreviewItem>();
        public ObservableCollection<SortItem> SortCriteria { get; set; } = new ObservableCollection<SortItem>();
        public ObservableCollection<string> AvailableSortColumns { get; set; } = new ObservableCollection<string>();

        private Commands.ParameterUpdateHandler _paramHandler;
        private ExternalEvent _paramExternalEvent;
        private ProSchedules.Models.ScheduleData _currentScheduleData;
        private Dictionary<ElementId, bool> _scheduleItemizeSettings = new Dictionary<ElementId, bool>();
        private System.Data.DataTable _rawScheduleData;
        private Dictionary<ElementId, ObservableCollection<SortItem>> _scheduleSortSettings = new Dictionary<ElementId, ObservableCollection<SortItem>>();

        #endregion

        #region Ctor / Init

        public DuplicateSheetsWindow(UIApplication app)
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

            // Create delete handler
            _deleteHandler = new ExternalEvents.SheetDeleteHandler();
            _deleteHandler.OnDeleteFinished += OnDeleteFinished;
            _deleteExternalEvent = ExternalEvent.Create(_deleteHandler);

            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            DeferWindowShow();

            // Window infrastructure
            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;

            // Window-level mouse hooks for resizing
            MouseMove += Window_MouseMove;
            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            // Theme
            LoadThemeState();
            LoadWindowState();
            DataContext = this;

            LoadData(app.ActiveUIDocument.Document);
            
            // Load persistent settings
            LoadSortSettings();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveSortSettings();
            SaveWindowState();
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
                sheetItem.PropertyChanged += OnSheetPropertyChanged;
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
            comboItems.Add(new ScheduleOption { Name = "All Sheets (Default)", Id = ElementId.InvalidElementId, Schedule = null });
            foreach(var s in schedules)
            {
                comboItems.Add(new ScheduleOption { Name = s.Name, Id = s.Id, Schedule = s });
            }
            SchedulesComboBox.ItemsSource = comboItems;
            SchedulesComboBox.SelectedIndex = 0;

            UpdateButtonStates();
            _isDataLoaded = true;
            TryShowWindow();
        }

        private void SchedulesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            
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
                    // Auto-apply immediately loading
                    ApplyCurrentSortLogic();
                }
            }
            else
            {
                RestoreSheetView();
            }
        }

        private void RestoreSheetView()
        {
            SheetsDataGrid.ItemsSource = null;
            SheetsDataGrid.Columns.Clear();
            
            SheetsDataGrid.ItemsSource = FilteredSheets;
            SheetsDataGrid.AutoGenerateColumns = false;
            
            InitializeSheetColumns();
        }

        private void InitializeSheetColumns()
        {
            SheetsDataGrid.Columns.Clear();
            
            var checkBoxColumn = CreateCheckBoxColumn();
            SheetsDataGrid.Columns.Add(checkBoxColumn);
            
            var numberCol = new DataGridTextColumn
            {
                Header = "Sheet Number",
                Binding = new System.Windows.Data.Binding("SheetNumber"),
                Width = new DataGridLength(150)
            };
            SheetsDataGrid.Columns.Add(numberCol);
            
            var nameCol = new DataGridTextColumn
            {
                Header = "Sheet Name",
                Binding = new System.Windows.Data.Binding("Name"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            };
            SheetsDataGrid.Columns.Add(nameCol);
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

            var baseStyle = FindResource("CustomDataGridCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center));
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center));

            // Create template for checkbox cells
            var cellTemplate = new DataTemplate();
            var checkBoxFactory = new FrameworkElementFactory(typeof(CheckBox));
            checkBoxFactory.SetValue(CheckBox.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.StyleProperty, FindResource("CustomCheckBoxStyle"));
            checkBoxFactory.SetBinding(CheckBox.IsCheckedProperty, new System.Windows.Data.Binding("IsSelected") { Mode = System.Windows.Data.BindingMode.TwoWay });
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
            // If multiple rows are selected, apply the checkbox state to all selected rows
            if (SheetsDataGrid.SelectedItems.Count > 1)
            {
                var checkBox = sender as CheckBox;
                if (checkBox == null) return;

                bool isChecked = checkBox.IsChecked == true;
                
                var view = SheetsDataGrid.ItemsSource;
                if (view is System.Data.DataView dataView)
                {
                    foreach (var selectedItem in SheetsDataGrid.SelectedItems)
                    {
                        if (selectedItem is System.Data.DataRowView rowView)
                        {
                            rowView["IsSelected"] = isChecked;
                        }
                    }
                }
                else if (view is ObservableCollection<SheetItem> sheets)
                {
                    foreach (var selectedItem in SheetsDataGrid.SelectedItems)
                    {
                        if (selectedItem is SheetItem sheet)
                        {
                            sheet.IsSelected = isChecked;
                        }
                    }
                }
            }
        }

        private void HeaderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var view = SheetsDataGrid.ItemsSource;
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
            var view = SheetsDataGrid.ItemsSource;
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
                
                if (!_scheduleItemizeSettings.ContainsKey(schedule.Id))
                {
                    _scheduleItemizeSettings[schedule.Id] = true;
                }
                bool isItemized = _scheduleItemizeSettings[schedule.Id];
                
                ItemizeCheckBox.IsChecked = isItemized;
                ItemizeCheckBox.Visibility = System.Windows.Visibility.Visible;

                var dt = new System.Data.DataTable();
                dt.Columns.Add("IsSelected", typeof(bool)).DefaultValue = false;
                dt.Columns.Add("RowState", typeof(string)).DefaultValue = "Unchanged";
                dt.Columns.Add("Count", typeof(int));

                // Detect column types
                for(int i = 0; i < data.Columns.Count; i++)
                {
                    string safeName = data.Columns[i];
                    int dupIdx = 1;
                    while(dt.Columns.Contains(safeName))
                    {
                        safeName = $"{data.Columns[i]} ({dupIdx++})";
                    }

                    // Check if column is numeric
                    bool isNumeric = true;
                    bool hasValue = false;
                    foreach(var r in data.Rows)
                    {
                        string val = r[i];
                        if (string.IsNullOrWhiteSpace(val)) continue;
                        hasValue = true;
                        if (!double.TryParse(val, out _))
                        {
                            isNumeric = false;
                            break;
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
                RefreshScheduleView(isItemized);
            }
            catch (Exception ex)
            {
                ShowPopup("Error Loading Schedule", ex.Message);
            }
        }

        private void RefreshScheduleView(bool itemize)
        {
            System.Data.DataTable viewTable = _rawScheduleData;
            
            if (!itemize && viewTable != null)
            {
                viewTable = viewTable.Clone();
                var grouped = _rawScheduleData.AsEnumerable()
                    .GroupBy(r => r["TypeName"]?.ToString() ?? "");
                
                foreach(var grp in grouped)
                {
                    var firstRow = grp.First();
                    var newRow = viewTable.NewRow();
                    newRow.ItemArray = firstRow.ItemArray;
                    newRow["Count"] = grp.Count();
                    viewTable.Rows.Add(newRow);
                }
            }
            
            SheetsDataGrid.ItemsSource = null;
            SheetsDataGrid.Columns.Clear();
            SheetsDataGrid.AutoGenerateColumns = false;
            
            if (viewTable == null) return;
            
            var baseStyle = FindResource("CustomDataGridCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            var cellTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "Pending"
            };
            cellTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(Colors.Yellow) { Opacity = 0.5 }));
            cellStyle.Triggers.Add(cellTrigger);
            
            // First add checkbox column
            var checkCol = CreateCheckBoxColumn();
            SheetsDataGrid.Columns.Add(checkCol);
            
            // Then add schedule data columns (skip RowState, ElementId, Count, IsSelected)
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "Count" };
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
                    IsReadOnly = false
                };
                SheetsDataGrid.Columns.Add(textCol);
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
                SheetsDataGrid.Columns.Add(countCol);
            }
            
            SheetsDataGrid.ItemsSource = new System.Data.DataView(viewTable);
        }

        private void Itemize_Checked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = true;
                RefreshScheduleView(true);
            }
        }

        private void Itemize_Unchecked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = false;
                RefreshScheduleView(false);
            }
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
                    ShowPopup("Success", $"Successfully created {success} sheet(s).");
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
                ShowPopup("No Sheets Selected", "Please select at least one sheet to duplicate.");
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
                        pendingSheet.PropertyChanged += OnSheetPropertyChanged;

                        // Add to collections
                        Sheets.Add(pendingSheet);

                        // Check if it matches current search filter
                        var searchText = SheetSearchBox?.Text?.ToLowerInvariant() ?? "";
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

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            // Protect current criteria from corruption during column clear
            var backup = new List<SortItem>();
            foreach(var item in SortCriteria) backup.Add(item);
            SortCriteria.Clear();

            // Always refresh available columns from current DataGrid state
            AvailableSortColumns.Clear();
            AvailableSortColumns.Add("(none)");
            if (SheetsDataGrid.Columns.Count > 0)
            {
                foreach (var col in SheetsDataGrid.Columns)
                {
                    if (col.Header is string header && !string.IsNullOrEmpty(header) 
                        && header != "Count" && header != "Sheet Number" && header != "Sheet Name") 
                    {
                         AvailableSortColumns.Add(header);
                    }
                     // Keep Sheet Number/Name if present
                     else if (col.Header is string h3 && (h3 == "Sheet Number" || h3 == "Sheet Name"))
                     {
                         AvailableSortColumns.Add(h3);
                     }
                }
            }
            
            // Restore items
            foreach(var item in backup) SortCriteria.Add(item);
            
            // Ensure at least one blank sort item if empty
            if (SortCriteria.Count == 0)
            {
                SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
            }

            SortPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void SortApply_Click(object sender, RoutedEventArgs e)
        {
            ApplyCurrentSortLogic();
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void ApplyCurrentSortLogic()
        {
            if (SheetsDataGrid.ItemsSource == null) return;
            
            System.ComponentModel.ICollectionView view = System.Windows.Data.CollectionViewSource.GetDefaultView(SheetsDataGrid.ItemsSource);
            view.SortDescriptions.Clear();

            foreach (var sortItem in SortCriteria)
            {
                if (string.IsNullOrEmpty(sortItem.SelectedColumn) || sortItem.SelectedColumn == "(none)") continue;
                
                // Map display name to binding path if necessary
                string propertyName = sortItem.SelectedColumn;
                
                // If using DataTable, property name is column name
                // If using SheetItem, property maps: 
                if (propertyName == "Sheet Number") propertyName = "SheetNumber"; // SheetItem property
                else if (propertyName == "Sheet Name") propertyName = "Name"; // SheetItem property
                
                view.SortDescriptions.Add(new System.ComponentModel.SortDescription(
                    propertyName, 
                    sortItem.IsAscending ? System.ComponentModel.ListSortDirection.Ascending : System.ComponentModel.ListSortDirection.Descending
                ));
            }
        }

        private void SortCancel_Click(object sender, RoutedEventArgs e)
        {
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        #region Persistence

        private void SaveSortSettings()
        {
            try
            {
                // 1. Commit current schedule settings to dictionary before saving
                if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
                {
                    var list = new ObservableCollection<SortItem>();
                    foreach(var item in SortCriteria) list.Add(item.Clone());
                    _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
                }

                var dtos = new List<SavedScheduleSort>();
                foreach(var kvp in _scheduleSortSettings)
                {
                    // Use robust ID extraction (Value or IntegerValue)
                    long idVal = GetIdValue(kvp.Key);
                    dtos.Add(new SavedScheduleSort 
                    { 
                        ScheduleId = idVal, 
                        Items = kvp.Value.ToList() 
                    });
                }

                string folder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "RK Tools", "ProSchedules");
                if (!System.IO.Directory.Exists(folder)) System.IO.Directory.CreateDirectory(folder);

                string file = System.IO.Path.Combine(folder, "sort_settings.json");
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(dtos, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(file, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private void LoadSortSettings()
        {
            try
            {
                string file = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "RK Tools", "ProSchedules", "sort_settings.json");
                if (System.IO.File.Exists(file))
                {
                    string json = System.IO.File.ReadAllText(file);
                    var dtos = Newtonsoft.Json.JsonConvert.DeserializeObject<List<SavedScheduleSort>>(json);
                    
                    if (dtos != null)
                    {
                        _scheduleSortSettings.Clear();
                        
                        foreach(var dto in dtos)
                        {
                            // Reconstruct ElementId
                            // Using the constructor available in older/newer API via helper or #if logic?
                            // Just use reflection or try generic constructor if possible.
                            // Actually, ElementId(long) exists in 2024+. ElementId(int) in older.
                            // Since we target multiple frameworks, let's try to map safely?
                            // Or just use the constructor that accepts long?
                            // Warnings earlier said ElementId(int) is deprecated, use ElementId(long).
                            
                            ElementId eid;
                            #if NET8_0_OR_GREATER
                                eid = new ElementId((long)dto.ScheduleId);
                            #else
                                eid = new ElementId((int)dto.ScheduleId);
                            #endif
                            
                            _scheduleSortSettings[eid] = new ObservableCollection<SortItem>(dto.Items);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }
        }

        private long GetIdValue(ElementId id)
        {
            #if NET8_0_OR_GREATER
                return id.Value;
            #else
                return id.IntegerValue;
            #endif
        }
        
        public class SavedScheduleSort
        {
            public long ScheduleId { get; set; }
            public List<SortItem> Items { get; set; }
        }
        
        #endregion

        private void AddSortLevel_Click(object sender, RoutedEventArgs e)
        {
            SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
        }


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
            }
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

        private void RemoveSortItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is SortItem item)
            {
                SortCriteria.Remove(item);
            }
        }

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

                obj = VisualTreeHelper.GetParent(obj);
            }

            // If we are here, we clicked empty space (background, borders, etc.) -> Deselect All
            if (SheetsDataGrid != null)
            {
                SheetsDataGrid.UnselectAll();
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
                if (!ValidateAllSheets())
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

                // 5. Disable buttons during operation
                ApplyButton.IsEnabled = false;
                DiscardButton.IsEnabled = false;
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
                        sheet.PropertyChanged -= OnSheetPropertyChanged; // Unsubscribe
                        Sheets.Remove(sheet);
                        FilteredSheets.Remove(sheet);
                    }

                    // Revert pending edits
                    foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList())
                    {
                        sheet.SheetNumber = sheet.OriginalSheetNumber;
                        sheet.Name = sheet.OriginalName;
                        sheet.State = SheetItemState.ExistingInRevit;
                    }

                    UpdateButtonStates();
                });
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void OnSheetPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (sender is SheetItem sheet && (e.PropertyName == "Name" || e.PropertyName == "SheetNumber"))
            {
                // Mark as edited if it's an existing sheet
                if (sheet.State == SheetItemState.ExistingInRevit &&
                    (sheet.SheetNumber != sheet.OriginalSheetNumber || sheet.Name != sheet.OriginalName))
                {
                    sheet.State = SheetItemState.PendingEdit;
                }

                ValidateSheetNumber(sheet);
                UpdateButtonStates();
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
                    ShowPopup("Success", $"Successfully updated {success} sheet(s).");
                }
            });
        }

        private void OnDeleteFinished(int success, int fail, string errorMsg, List<ElementId> deletedIds)
        {
            Dispatcher.Invoke(() =>
            {
                if (deletedIds != null && deletedIds.Count > 0)
                {
                    var deletedSet = new HashSet<ElementId>(deletedIds);
                    foreach (var sheet in _pendingDeleteItems.Where(s => deletedSet.Contains(s.Id)).ToList())
                    {
                        sheet.PropertyChanged -= OnSheetPropertyChanged;
                        Sheets.Remove(sheet);
                        FilteredSheets.Remove(sheet);
                    }
                }

                int totalSuccess = success + _pendingDeleteLocalCount;
                _pendingDeleteItems.Clear();
                _pendingDeleteLocalCount = 0;

                UpdateButtonStates();

                if (fail > 0)
                {
                    ShowPopup("Delete Report", $"Deleted: {totalSuccess}\nFailed: {fail}\nLast Error: {errorMsg}");
                }
                else
                {
                    ShowPopup("Success", $"Deleted {totalSuccess} sheet(s).");
                }
            });
        }

        private void ValidateSheetNumber(SheetItem sheet)
        {
            var duplicates = Sheets.Where(s =>
                s != sheet &&
                s.SheetNumber == sheet.SheetNumber &&
                (s.State == SheetItemState.ExistingInRevit || s.State == SheetItemState.PendingEdit)
            ).ToList();

            sheet.HasNumberConflict = duplicates.Any();
            sheet.ValidationError = duplicates.Any()
                ? $"Sheet number '{sheet.SheetNumber}' already exists"
                : null;
        }

        private bool ValidateAllSheets()
        {
            bool hasErrors = false;

            foreach (var sheet in Sheets.Where(s => s.HasUnsavedChanges))
            {
                ValidateSheetNumber(sheet);
                if (sheet.HasNumberConflict)
                {
                    hasErrors = true;
                }
            }

            if (hasErrors)
            {
                ShowPopup("Validation Error", "Please fix duplicate sheet numbers before applying.");
                return false;
            }

            return true;
        }

        private void UpdateButtonStates()
        {
            var pendingCount = Sheets.Count(s => s.HasUnsavedChanges);
            var hasConflicts = Sheets.Any(s => s.HasNumberConflict);

            ApplyButton.IsEnabled = pendingCount > 0 && !hasConflicts;
            DiscardButton.IsEnabled = pendingCount > 0;
        }

        private void Rename_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Get selected sheets
                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                if (selectedSheets.Count == 0)
                {
                    ShowPopup("No Sheets Selected", "Please select at least one sheet to rename.");
                    return;
                }

                // Reset rename controls
                RenameParameter.SelectedIndex = 0;
                RenameFindText.Clear();
                RenameReplaceText.Clear();
                RenamePrefixText.Clear();
                RenameSuffixText.Clear();

                // Populate preview
                UpdateRenamePreview();

                // Show popup
                RenamePopupOverlay.Visibility = System.Windows.Visibility.Visible;
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
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

                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                if (selectedSheets.Count == 0) return;

                // Determine which parameter to rename
                bool isSheetNumber = RenameParameter?.SelectedIndex == 0;

                string findText = RenameFindText?.Text ?? "";
                string replaceText = RenameReplaceText?.Text ?? "";
                string prefix = RenamePrefixText?.Text ?? "";
                string suffix = RenameSuffixText?.Text ?? "";

                foreach (var sheet in selectedSheets)
                {
                    string original = isSheetNumber ? sheet.SheetNumber : sheet.Name;
                    string newValue = original;

                    // Apply find/replace
                    if (!string.IsNullOrEmpty(findText))
                    {
                        newValue = newValue.Replace(findText, replaceText);
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

                    var previewItem = new RenamePreviewItem(sheet, original)
                    {
                        New = newValue
                    };

                    RenamePreviewItems.Add(previewItem);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating rename preview: {ex.Message}");
            }
        }

        private void RenameApply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isSheetNumber = RenameParameter.SelectedIndex == 0;

                // Apply changes to the main DataGrid
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

        private void RenameCancel_Click(object sender, RoutedEventArgs e)
        {
            RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void Parameters_Click(object sender, RoutedEventArgs e)
        {
            ParametersPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ParametersClose_Click(object sender, RoutedEventArgs e)
        {
            ParametersPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void ParametersPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        private void ShowDeleteConfirmPopup(string title, string message, Action onConfirmAction)
        {
            DeleteConfirmPopupTitle.Text = title;
            DeleteConfirmPopupMessage.Text = message;
            _onDeleteConfirmAction = onConfirmAction;
            DeleteConfirmPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void CloseDeleteConfirmPopup()
        {
            DeleteConfirmPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            _onDeleteConfirmAction = null;
        }

        private void DeleteConfirmOk_Click(object sender, RoutedEventArgs e)
        {
            var action = _onDeleteConfirmAction;
            CloseDeleteConfirmPopup();
            action?.Invoke();
        }

        private void DeleteConfirmCancel_Click(object sender, RoutedEventArgs e)
        {
            CloseDeleteConfirmPopup();
        }

        private void DeleteConfirmPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedSheets = Sheets.Where(s => s.IsSelected).ToList();
                if (selectedSheets.Count == 0)
                {
                    ShowPopup("No Sheets Selected", "Please select at least one sheet to delete.");
                    return;
                }

                string message = "This will permanently delete the selected sheets from the project.\nThis action cannot be undone.\n\nContinue?";
                ShowDeleteConfirmPopup("Delete Sheets", message, () => ExecuteDelete(selectedSheets));
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void ExecuteDelete(List<SheetItem> selectedSheets)
        {
            var pendingItems = selectedSheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
            foreach (var sheet in pendingItems)
            {
                sheet.PropertyChanged -= OnSheetPropertyChanged;
                Sheets.Remove(sheet);
                FilteredSheets.Remove(sheet);
            }

            _pendingDeleteLocalCount = pendingItems.Count;

            var existingItems = selectedSheets
                .Where(s => s.State == SheetItemState.ExistingInRevit || s.State == SheetItemState.PendingEdit)
                .ToList();

            if (existingItems.Count == 0)
            {
                UpdateButtonStates();
                ShowPopup("Success", $"Deleted {_pendingDeleteLocalCount} sheet(s).");
                _pendingDeleteLocalCount = 0;
                return;
            }

            _pendingDeleteItems = existingItems;
            _deleteHandler.SheetIdsToDelete = existingItems
                .Select(s => s.Id)
                .Where(id => id != null && id != ElementId.InvalidElementId)
                .ToList();

            _deleteExternalEvent.Raise();
        }

        private void RenamePopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Optionally close on background click
            // RenamePopupOverlay.Visibility = Visibility.Collapsed;
        }

        private void SheetsDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                var dataGrid = sender as DataGrid;
                if (dataGrid?.SelectedItems != null && dataGrid.SelectedItems.Count > 0)
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

        private void ClearSheetSearch_Click(object sender, RoutedEventArgs e)
        {
            SheetSearchBox.Clear();
        }

        private void SheetCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is SheetItem clickedItem)
            {
                e.Handled = true;
                bool newState = !(checkBox.IsChecked ?? false);
                checkBox.IsChecked = newState;

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

        #region Custom Popup



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
        }

        private void ShowConfirmPopup(string title, string message, Action onConfirmAction)
        {
            ConfirmPopupTitle.Text = title;
            ConfirmPopupMessage.Text = message;
            _onConfirmAction = onConfirmAction;
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
            Loaded += DuplicateSheetsWindow_Loaded;
        }

        private void DuplicateSheetsWindow_Loaded(object sender, RoutedEventArgs e)
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
