using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistApproverModel : BaseModel
    {
        public long ChecklistID { get; set; }
        public string ADGroup { get; set; }
        public string EmployeeID { get; set; }
        public string Approve { get; set; }
        public string Revise { get; set; }
        public string Edit { get; set; }
        public string Reject { get; set; }
        public int Tier { get; set; }
        public List<string> EmployeeIDs { get; set; }
        public ChecklistApprovalModel Approval { get; set; }
    }
}