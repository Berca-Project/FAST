using System.Web;
using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistComponentModel : BaseModel
    {
        public long ChecklistID { get; set; }
        public string Segment { get; set; }
        public string ComponentType { get; set; }
        public string ComponentName { get; set; }
        public string AdditionalValue { get; set; }
        public HttpPostedFileBase ComponentFile { get; set; }
        public int OrderNum { get; set; }
        public int ColumnNum { get; set; }
        public bool IsRequired { get; set; }

    }
}