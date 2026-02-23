namespace DefectManagement.Models
{
    public class DefectReportByOperation
    {
        public string Operation { get; set; }
        public int TotalRecords { get; set; }
        public int TotalQtyNG { get; set; }
        public double Percentage { get; set; }
        public List<DefectReportByDefectName> DefectDetails { get; set; }
    }

    public class DefectReportByDefectName
    {
        public string DefectCode { get; set; }
        public string DefectName { get; set; }
        public int Count { get; set; }
        public int TotalQtyNG { get; set; }
        public double Percentage { get; set; }
    }

    public class DefectDailyChartData
    {
        public List<string> Dates { get; set; }
        public List<int> QtyNGValues { get; set; }
        public List<int> CountValues { get; set; }
        public List<double> MovingAvg7Days { get; set; }
    }
}