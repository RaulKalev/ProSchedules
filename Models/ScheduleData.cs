using System.Collections.Generic;

namespace ProSchedules.Models
{
    public class ScheduleData
    {
        public List<string> Columns { get; set; } = new List<string>();
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
        public Dictionary<string, bool> IsTypeParameter { get; set; } = new Dictionary<string, bool>();
    }
}
