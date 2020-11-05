using System;
using System.Web;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistValueHistoryModel : BaseModel
    {
        public long ChecklistSubmitID { get; set; }
        public long ChecklistValueID { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long UserID { get; set; }
    }
}