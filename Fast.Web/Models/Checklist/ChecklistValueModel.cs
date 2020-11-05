using System;
using System.Web;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistValueModel : BaseModel
    {
        public long ChecklistSubmitID { get; set; }
        public long ChecklistComponentID { get; set; }
        public HttpPostedFileBase ValueFile { get; set; }
        public string Value { get; set; }
        public string ValueType { get; set; }
        public DateTime? ValueDate { get; set; }
        public int OrderNum { get; set; }
        public int ColumnNum { get; set; }
    }
}