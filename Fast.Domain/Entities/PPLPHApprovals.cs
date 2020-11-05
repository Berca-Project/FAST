using System;

namespace Fast.Domain.Entities
{
	public class PPLPHApprovals : BaseEntity
    {
        public long LPHSubmissionID { get; set; }
        public long UserID { get; set; }
		public long LocationID { get; set; }
		public string Status { get; set; }
        public string Notes { get; set; }
        public long ApproverID { get; set; }
        public string Shift { get; set; }
        public DateTime Date { get; set; }
        public string ApproverEmployeeID { get; set; }
        public string UserEmployeeID { get; set; }


        public string ApproverRole { get; set; }
        public string ApproverJobTitle { get; set; }
        public string UserFullName { get; set; }
        public string ApproverFullName { get; set; }

    }
}
