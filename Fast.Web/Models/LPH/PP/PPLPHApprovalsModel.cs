using System;

namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHApprovalsModel : BaseModel
    {
        public long LPHSubmissionID { get; set; }
        public long UserID { get; set; }
		public string User { get; set; }
		public long LocationID { get; set; }
		public string Location { get; set; }
		public string Status { get; set; }
        public string Notes { get; set; }
        public long ApproverID { get; set; }
		public string Approver { get; set; }
		public string Shift { get; set; }
        public DateTime Date { get; set; }
        public string ApproverName { get; set; }
		public bool IsNeedMyApproval { get; set; }
		public string LPHType { get; set; }
        public string Role { get; set; }
        public string UserEmployeeID { get; set; }
        public string ApproverEmployeeID { get; set; }

        public string ApproverRole { get; set; }
        public string ApproverJobTitle { get; set; }

        public string UserFullName { get; set; }
        public string ApproverFullName { get; set; }

    }
}