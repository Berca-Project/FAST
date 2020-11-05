using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class ReportModel
    {
        public ReportModel()
        {           
            Brands = new List<string>();
        }
        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public List<string> Brands { get; set; }
        public DateTime DateFilter { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Machine { get; set; }

    }
}