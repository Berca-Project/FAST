using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldKretekSAPModel : BaseModel
    {
        public string WeekYear { get; set; }
        //public int Year { get; set; }
        //public int Week { get; set; }
        public string Material { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialDescription { get; set; }
        public string BaseUnitofMeasure { get; set; }
        public string Category { get; set; }
        public string MaterialGroupDesc { get; set; }
        public Nullable<double> GoodsReceiptQty { get; set; }
        public Nullable<double> GoodIssueQty { get; set; }
        public Nullable<double> WIPQuantity { get; set; }
        public Nullable<double> StockTakeQuantity { get; set; }
        public Nullable<double> ScrapQty { get; set; }
        public Nullable<double> NonBOMQty { get; set; }
        public Nullable<double> TotTKGWIP { get; set; }
        public Nullable<double> YieldPercent { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
    }
}