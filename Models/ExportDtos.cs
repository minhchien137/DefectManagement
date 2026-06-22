namespace DefectManagement.Models
{
    public class ExportDefectDto
    {
        public string? DateFrom   { get; set; }
        public string? DateTo     { get; set; }
        public string? WorkOrder  { get; set; }
        public string? Operation  { get; set; }
        public string? DefectCode { get; set; }
        public string? ItemCode   { get; set; }
        public ChartImagesDto? ChartImages { get; set; }
    }

    public class ChartImagesDto
    {
        public string? ByOperation { get; set; }
        public string? Pareto      { get; set; }
        public string? Top10       { get; set; }
        public string? Trend       { get; set; }
    }
}
