using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class ExtractDataModel
    {
        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string LPHType { get; set; }

        public string location1 { get; set; }
        public string location2 { get; set; }
        public string location3 { get; set; }
        public string Status { get; set; }
    }
}