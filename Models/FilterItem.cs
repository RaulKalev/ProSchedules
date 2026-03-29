using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ProSchedules.Models
{
    public class FilterItem : INotifyPropertyChanged
    {
        private string _selectedColumn = "(none)";
        private string _selectedCondition = "equals";
        private string _value = "";

        public string SelectedColumn
        {
            get => _selectedColumn;
            set
            {
                _selectedColumn = value;
                OnPropertyChanged(nameof(SelectedColumn));
                ColumnChanged?.Invoke(this);
            }
        }

        public string SelectedCondition
        {
            get => _selectedCondition;
            set { _selectedCondition = value; OnPropertyChanged(nameof(SelectedCondition)); }
        }

        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(nameof(Value)); }
        }

        public ObservableCollection<string> AvailableValues { get; set; } = new ObservableCollection<string>();

        public event System.Action<FilterItem> ColumnChanged;

        public static readonly string[] Conditions = new[]
        {
            "equals",
            "does not equal",
            "is greater than",
            "is greater than or equal to",
            "is less than",
            "is less than or equal to",
            "contains",
            "does not contain",
            "begins with",
            "does not begin with",
            "ends with",
            "does not end with",
            "has a value",
            "has no value"
        };

        public FilterItem Clone()
        {
            var clone = new FilterItem
            {
                SelectedColumn = this.SelectedColumn,
                SelectedCondition = this.SelectedCondition,
                Value = this.Value
            };
            foreach (var v in AvailableValues)
                clone.AvailableValues.Add(v);
            return clone;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
