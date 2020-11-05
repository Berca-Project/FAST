using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldDietModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public DateTime Date { get; set; }
        public string Blend { get; set; }
        public long Ops { get; set; }
        public Nullable<double> CVIB0069 { get; set; }
        public Nullable<double> CSFR0022 { get; set; }
        public Nullable<double> DSCL0034 { get; set; }
        public Nullable<double> CVIB0070 { get; set; }
        public Nullable<double> RV0054 { get; set; }
        public Nullable<double> Input { get; set; }
        public Nullable<double> Output { get; set; }
        public Nullable<double> WetYield { get; set; }
        public Nullable<double> DryYield { get; set; }
        public Nullable<double> Casing { get; set; }
        public Nullable<double> DryInput { get; set; }
        public Nullable<double> DryCasing { get; set; }
        public Nullable<double> DryWaste { get; set; }
        public Nullable<double> WetWaste { get; set; }
        public Nullable<double> DustWaste { get; set; }
        public Nullable<double> HotDustWaste { get; set; }
        public Nullable<double> CGNSolar { get; set; }
        public Nullable<double> STEAM { get; set; }
        public Nullable<double> Awal { get; set; }
        public Nullable<double> Terima { get; set; }
        public Nullable<double> Akhir { get; set; }
        public Nullable<double> Pakai { get; set; }
        public Nullable<double> RateCO2 { get; set; }
    }
}