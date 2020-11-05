using System;
using System.Collections.Generic;

namespace Fast.Domain.Entities
{
	public class ChecklistApproval : BaseEntity
	{
        public long ChecklistSubmitID { get; set; }
        public long ChecklistApproverID { get; set; }
        public string Role { get; set; }
        public string EmployeeID { get; set; }
        public string Status { get; set; }
        public string Comments { get; set; }
    }
}
