using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldMCDietModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public Nullable<double> MCFlake { get; set; }
        public Nullable<double> MCKrosok { get; set; }
        public Nullable<double> CVIB0069 { get; set; }
        public Nullable<double> CSFR0022 { get; set; }
        public Nullable<double> DSCL0034 { get; set; }
        public Nullable<double> CVIB0070 { get; set; }
        public Nullable<double> RV0054 { get; set; }
        public Nullable<double> DM { get; set; }
        public Nullable<double> Flake { get; set; }
        public Nullable<double> MCPacking { get; set; }
    }
}