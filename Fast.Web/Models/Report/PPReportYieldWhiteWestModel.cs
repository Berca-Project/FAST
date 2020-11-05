using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldWhiteWestModel : BaseModel
    {
        public DateTime EndTime { get; set; }
        public string SAPID { get; set; }
        public string BatchIdent { get; set; }
        public string BlendCode { get; set; }
        public Nullable<double> ProducedQTY { get; set; }
        public Nullable<double> Tobacco { get; set; }
        public Nullable<double> TotalStems { get; set; }
        public Nullable<double> TotalExpandedTobacco { get; set; }
        public Nullable<double> TotalSmallLamina { get; set; }
        public Nullable<double> TotalRipperShorts { get; set; }
        public Nullable<double> WetYield { get; set; }
        public Nullable<double> WetTarget { get; set; }
        public Nullable<double> DryTobacco { get; set; }
        public Nullable<double> DryISCRES { get; set; }
        public Nullable<double> DryET { get; set; }
        public Nullable<double> DrySL { get; set; }
        public Nullable<double> DryRS { get; set; }
        public Nullable<double> FinalOV { get; set; }
        public Nullable<double> InvoiceOV { get; set; }
        public Nullable<double> DMBS { get; set; }
        public Nullable<double> DMBT { get; set; }
        public Nullable<double> DMBC { get; set; }
        public Nullable<double> DMAC { get; set; }
        public Nullable<double> BS { get; set; }
        public Nullable<double> BT { get; set; }
        public Nullable<double> BC { get; set; }
        public Nullable<double> AC { get; set; }
        public Nullable<double> Packing { get; set; }
        public Nullable<double> DryYield { get; set; }
    }
}