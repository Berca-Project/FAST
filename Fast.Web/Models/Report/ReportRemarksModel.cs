using System;

namespace Fast.Web.Models.Report
{
    public class ReportRemarksModel : BaseModel
    {
        public string RSA { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public long LocationID { get; set; }
        public string Focus { get; set; }
        public string ActionPlan { get; set; }
        public long UserID { get; set; }
    }
}