using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHValuesModel : BaseModel
    {
        public long LPHComponentID { get; set; }
        public string Value { get; set; }
        public DateTime ValueDate { get; set; }
        public string ValueType { get; set; }
        public long? SubmissionID { get; set; }
        public long UserID { get; set; }
        public DateTime Date { get; set; }
        public DateTime CompleteDate { get; set; }
        public long LocationID { get; set; }
    }
}