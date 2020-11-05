using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class InputTargetModel : BaseModel
    {
        public string KPI { get; set; }
        public decimal Value { get; set; }
        public string Month { get; set; }
        public string Version { get; set; }
        public string ProductionCenter { get; set; }
        public string Country { get; set; }
        public long CountryID { get; set; }
        public long LocationID { get; set; }
        public long ProdCenterID { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
    }
}