using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportDowntimeModel : BaseModel
    {
        public long PK { get; set; }
        public double BatchID { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string FilterLine { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public double Duration { get; set; }
        public string LogUser { get; set; }
        public string SubmitHMI { get; set; }
        public DateTime SubmitDateTime { get; set; }
        public string LineProd { get; set; }
        public string PlanUnplan { get; set; }
        public string Category { get; set; }
        public string UpDownStream { get; set; }
        public string Issue { get; set; }
        public string Description { get; set; }
        public string Remark { get; set; }
        public string Status { get; set; }
        public string Flag { get; set; }
        public string UserEmployeeID { get; set; }
        public string ApprovedBy { get; set; }
        public string ApproverEmployeeID { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
    }
}