using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models
{
    public class ReportModel
    {
        public long LocationID { get; set; }
        public long ProdCenterID { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Week { get; set; }
    }
}