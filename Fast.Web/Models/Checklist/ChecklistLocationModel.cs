using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ChecklistLocationModel : BaseModel
    {
        public long ChecklistID { get; set; }
        public long LocationID { get; set; }
    }
}