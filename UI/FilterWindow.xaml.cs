using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ProSchedules.Models;

namespace ProSchedules.UI
{
    public partial class FilterWindow : Window
    {
        private DuplicateSheetsWindow _parent;
        private List<FilterItem> _checkpoint;

        public ObservableCollection<FilterItem> FilterCriteria { get; private set; }
        public ObservableCollection<string> AvailableFilterColumns { get; private set; }

        public FilterWindow(DuplicateSheetsWindow parent)
        {
            InitializeComponent();
            _parent = parent;

            FilterCriteria = parent.FilterCriteria;
            AvailableFilterColumns = parent.AvailableFilterColumns;

            DataContext = this;

            foreach (var item in FilterCriteria)
                item.ColumnChanged += OnFilterColumnChanged;

            FilterCriteria.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                    foreach (FilterItem item in e.NewItems)
                        item.ColumnChanged += OnFilterColumnChanged;
            };

            Closing += (s, e) => RestoreCheckpoint();

            CreateCheckpoint();
        }

        private void OnFilterColumnChanged(FilterItem item)
        {
            item.AvailableValues.Clear();
            if (string.IsNullOrEmpty(item.SelectedColumn) || item.SelectedColumn == "(none)") return;

            var values = _parent.GetUniqueValuesForColumn(item.SelectedColumn);
            foreach (var v in values)
                item.AvailableValues.Add(v);
        }

        private void CreateCheckpoint()
        {
            _checkpoint = new List<FilterItem>();
            foreach (var item in FilterCriteria) _checkpoint.Add(item.Clone());
        }

        private void RestoreCheckpoint()
        {
            FilterCriteria.Clear();
            foreach (var item in _checkpoint) FilterCriteria.Add(item.Clone());
        }

        private void AddFilterRule_Click(object sender, RoutedEventArgs e)
        {
            FilterCriteria.Add(new FilterItem());
        }

        private void RemoveFilterRule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is FilterItem item)
            {
                FilterCriteria.Remove(item);
            }
        }

        private void FilterApply_Click(object sender, RoutedEventArgs e)
        {
            _parent.ApplyFilterLogic();
            _parent.CommitFilterSettings();
            _parent.SaveSettingsToStorage();
            CreateCheckpoint();
        }

        private void FilterClear_Click(object sender, RoutedEventArgs e)
        {
            FilterCriteria.Clear();
            _parent.ApplyFilterLogic();
            _parent.CommitFilterSettings();
            _parent.SaveSettingsToStorage();
            CreateCheckpoint();
        }

        private void FilterCancel_Click(object sender, RoutedEventArgs e)
        {
            RestoreCheckpoint();
            Close();
        }
    }
}
