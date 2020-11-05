using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldUploadModel : BaseModel
    {
        public string WeekYear { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
    }
}