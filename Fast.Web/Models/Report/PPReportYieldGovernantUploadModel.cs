using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldGovernantUploadModel : BaseModel
    {
        public string WeekStart { get; set; }
        public string WeekEnd { get; set; }
        public string Type { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
    }
}