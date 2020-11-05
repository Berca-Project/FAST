using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.LPH
{
    public class LPHValueHistoriesModel : BaseModel
    {
        public long LPHValuesID { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long? UserID { get; set; }
    }
}