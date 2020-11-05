using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class InputOVModel : BaseModel
    {
        public string WasteCategory { get; set; }
        public string Week { get; set; }
        public string OV { get; set; }
        public long LocationID { get; set; }
    }
}