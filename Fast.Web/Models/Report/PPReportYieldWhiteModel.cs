using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldWhiteModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public string Blend { get; set; }
        public Nullable<double> InfeedMaterialWet { get; set; }
        public Nullable<double> InfeedMaterialDry { get; set; }
        public Nullable<double> SumInputMaterialDry { get; set; }
        public Nullable<double> CutFiller { get; set; }
        public Nullable<double> RS_Addback { get; set; }
        public Nullable<double> CFDry { get; set; }
        public Nullable<double> RS_AddDry { get; set; }
        public Nullable<double> CFDryExclude { get; set; }
        public Nullable<double> RS_AddbackWet { get; set; }
        public Nullable<double> AvgFinalOV { get; set; }
        public Nullable<double> RS_AddbackPercen { get; set; }
        public Nullable<double> DryYield { get; set; }
        public Nullable<double> WetYield { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
    }
}