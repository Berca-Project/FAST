using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldKretekWestModel : BaseModel
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
        public Nullable<double> TotalCloves { get; set; }
        public Nullable<double> TotalOffspec { get; set; }
        public Nullable<double> CSF { get; set; }
        public Nullable<double> WetYield { get; set; }
        public Nullable<double> WetTarget { get; set; }

        public Nullable<double> DryTobacco { get; set; }
        public Nullable<double> DryISCRES { get; set; }
        public Nullable<double> DryRTC { get; set; }
        public Nullable<double> DryET { get; set; }
        public Nullable<double> DryCLOVE { get; set; }
        public Nullable<double> DryCSF { get; set; }
        public Nullable<double> DryOfSpec { get; set; }
        public Nullable<double> FinalOV { get; set; }
        public Nullable<double> InvoiceOV { get; set; }
        public Nullable<double> DMBC { get; set; }
        public Nullable<double> TotalBrightCasing { get; set; }
        public Nullable<double> TotalBurleySpray { get; set; }
        public Nullable<double> TotalAfterCut { get; set; }
        public Nullable<double> Packing { get; set; }
        public Nullable<double> DryYield { get; set; }
        public Nullable<double> DMAC { get; set; }
        public Nullable<double> DryCasing { get; set; }
        public Nullable<double> DryAC { get; set; }
    }
}