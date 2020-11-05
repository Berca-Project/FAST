using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldTargetModel : BaseModel
    {
        public long? LocationID { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }
        public decimal? Target { get; set; }
    }
}