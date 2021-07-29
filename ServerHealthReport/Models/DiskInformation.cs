namespace ServerHealthReport.Models
{
    public class DiskInformation
    {
        public string DiskDriverLetter { get; set; }
        public string DriveName { get; set; }
        public decimal TotalSize { get; set; }
        public decimal Used { get; set; }
        public decimal Free { get; set; }
        public decimal FreePercentage { get; set; }
    }
}
