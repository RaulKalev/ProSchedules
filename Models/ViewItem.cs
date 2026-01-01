using Autodesk.Revit.DB;
using System.ComponentModel;

namespace PlaceViews.Models
{
    public class ViewItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string ViewType { get; set; }
        public ElementId Id { get; set; }
        
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public ViewItem(View view)
        {
            Name = view.Name;
            ViewType = view.ViewType.ToString();
            Id = view.Id;
            IsSelected = false;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
