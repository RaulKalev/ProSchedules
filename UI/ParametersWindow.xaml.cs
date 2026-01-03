using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using ProSchedules.Models;

namespace ProSchedules.UI
{


    public partial class ParametersWindow : Window
    {
        public ObservableCollection<ParameterItem> AvailableParams { get; set; } = new ObservableCollection<ParameterItem>();
        public ObservableCollection<ParameterItem> ScheduledParams { get; set; } = new ObservableCollection<ParameterItem>();

        public event Action<List<ParameterItem>> OnApply;

        public ParametersWindow(List<ParameterItem> available, List<ParameterItem> scheduled, string categoryName)
        {
            InitializeComponent();
            DataContext = this;
            Title = $"Schedule Parameters - {categoryName}";

            foreach (var p in available) AvailableParams.Add(p);
            foreach (var p in scheduled) ScheduledParams.Add(p);

            AvailableParamsList.ItemsSource = AvailableParams;
            ScheduledParamsList.ItemsSource = ScheduledParams;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            var selected = AvailableParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            foreach (var item in selected)
            {
                AvailableParams.Remove(item);
                ScheduledParams.Add(item);
            }
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            foreach (var item in selected)
            {
                ScheduledParams.Remove(item);
                AvailableParams.Add(item);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            if (selected.Count == 0) return;

            // Sort by index to move them properly
            var sortedSelected = selected.OrderBy(x => ScheduledParams.IndexOf(x)).ToList();

            foreach (var item in sortedSelected)
            {
                int index = ScheduledParams.IndexOf(item);
                if (index > 0)
                {
                    ScheduledParams.Move(index, index - 1);
                }
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            if (selected.Count == 0) return;

            // Sort reverse to move from bottom up
            var sortedSelected = selected.OrderByDescending(x => ScheduledParams.IndexOf(x)).ToList();

            foreach (var item in sortedSelected)
            {
                int index = ScheduledParams.IndexOf(item);
                if (index < ScheduledParams.Count - 1)
                {
                    ScheduledParams.Move(index, index + 1);
                }
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            OnApply?.Invoke(ScheduledParams.ToList());
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
