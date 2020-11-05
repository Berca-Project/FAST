using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldIMLModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long? LocationID { get; set; }
        public string Type { get; set; }
        public decimal? Value { get; set; }
    }
}