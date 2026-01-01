using PlaceViews.ExternalEvents;

namespace PlaceViews.Models
{
    public class DuplicateOptions
    {
        public int NumberOfCopies { get; set; }
        public SheetDuplicateMode DuplicateMode { get; set; }
        public bool KeepLegends { get; set; }
        public bool KeepSchedules { get; set; }
        public bool CopyRevisions { get; set; }
        public bool CopyParameters { get; set; }
    }
}
