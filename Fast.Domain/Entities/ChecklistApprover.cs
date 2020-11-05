using System;
using System.Collections.Generic;

namespace Fast.Domain.Entities
{
	public class ChecklistApprover : BaseEntity
	{
        public long ChecklistID { get; set; }
        public string ADGroup { get; set; }
        public string EmployeeID { get; set; }
        public string Approve { get; set; }
        public string Revise { get; set; }
        public string Edit { get; set; }
        public string Reject { get; set; }
        public int Tier { get; set; }
    }
}
