using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using ProSchedules.UI;

namespace ProSchedules.UI
{
    public partial class SortingWindow : Window
    {
        private ProSchedulesWindow _parent;
        private List<SortItem> _checkpoint;

        // Drag-and-drop state
        private SortItem _dragSource;
        private Point _dragStartPoint;
        private bool _isDragInProgress;

        public ObservableCollection<SortItem> SortCriteria { get; private set; }
        public ObservableCollection<string> AvailableSortColumns { get; private set; }

        public SortingWindow(ProSchedulesWindow parent)
        {
            InitializeComponent();
            _parent = parent;
            
            // Link to parent collections
            SortCriteria = parent.SortCriteria;
            AvailableSortColumns = parent.AvailableSortColumns;
            
            DataContext = this;



            // Create initial checkpoint
            CreateCheckpoint();
            
            // Handle window move via TitleBar drag (handled by TitleBar control usually, but checking ProSchedulesWindow logic)
            // TitleBar control usually has logic. If not, we can add standard drag logic.
            // ProSchedulesWindow is WindowStyle=None, so it likely handles dragging.
        }

        private void CreateCheckpoint()
        {
            _checkpoint = new List<SortItem>();
            foreach (var item in SortCriteria) _checkpoint.Add(item.Clone());
        }

        private void RestoreCheckpoint()
        {
            SortCriteria.Clear();
            foreach (var item in _checkpoint) SortCriteria.Add(item.Clone());
        }

        private void AddSortLevel_Click(object sender, RoutedEventArgs e)
        {
            SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
        }

        private void RemoveSortItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is SortItem item)
            {
                SortCriteria.Remove(item);
            }
        }

        private void SortApply_Click(object sender, RoutedEventArgs e)
        {
            _parent.ApplyCurrentSortLogicInternal();
            _parent.CommitSortSettingsToStorage();
            CreateCheckpoint();
        }

        private void SortCancel_Click(object sender, RoutedEventArgs e)
        {
            RestoreCheckpoint();
            Close();
        }

        // ── Drag-and-drop reordering ────────────────────────────────────────────

        private void DragHandle_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragSource = GetSortItemFromSender(sender);
            _dragStartPoint = e.GetPosition(null);
        }

        private void DragHandle_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _dragSource == null || _isDragInProgress) return;

            var pos = e.GetPosition(null);
            if (Math.Abs(pos.X - _dragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(pos.Y - _dragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance) return;

            _isDragInProgress = true;
            DragDrop.DoDragDrop((DependencyObject)sender, _dragSource, DragDropEffects.Move);
            _isDragInProgress = false;
            _dragSource = null;
        }

        private static SortItem GetSortItemFromSender(object sender)
        {
            var fe = sender as FrameworkElement;
            while (fe != null)
            {
                if (fe.DataContext is SortItem item) return item;
                fe = VisualTreeHelper.GetParent(fe) as FrameworkElement;
            }
            return null;
        }

        private void SortCard_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(typeof(SortItem))
                ? DragDropEffects.Move
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void SortCard_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(SortItem))) return;
            var source = e.Data.GetData(typeof(SortItem)) as SortItem;
            var target = (sender as FrameworkElement)?.DataContext as SortItem;
            if (source == null || target == null || ReferenceEquals(source, target)) return;

            int sourceIdx = SortCriteria.IndexOf(source);
            int targetIdx = SortCriteria.IndexOf(target);
            if (sourceIdx < 0 || targetIdx < 0) return;

            SortCriteria.Move(sourceIdx, targetIdx);
            e.Handled = true;
        }
    }
}
