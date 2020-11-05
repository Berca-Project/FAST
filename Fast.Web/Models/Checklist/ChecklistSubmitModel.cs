using System;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistSubmitModel : BaseModel
    {
        public long ChecklistID { get; set; }
        public DateTime CompleteDate { get; set; }
        public long UserID { get; set; }
        public string Shift { get; set; }
        public DateTime date { get; set; }
        public string Status { get; set; }
        public string Comment { get; set; }
        public ChecklistModel Checklist { get; set; }
        public EmployeeModel Submiter { get; set; }
        public UserModel User { get; set; }
        public bool IsEditable { get; set; }
        public bool IsEdited { get; set; }
        public bool IsComplete { get; set; }

        public bool isApprover { get; set; }
        public long Location { get; set; }

        public int ContentOk { get; set; }
        public int ContentAll { get; set; }
    }
}