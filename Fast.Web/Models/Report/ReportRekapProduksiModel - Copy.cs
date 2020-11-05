using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class RekapProduksiRPH1Model
    {
        public List<WppModel> wppList { get; set; }
        public List<WppSimpleModel> wppSimpleList { get; set; }
        public List<RekapProduksiDataLPHModel> DataLPH { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }

    public class RekapProduksiDataLPHModel
    {
        public DateTime Date { get; set; }
        public string Location { get; set; }
        public string Brand { get; set; }
        public string Machine { get; set; }
        public double Actual { get; set; }
    }

    public class ProdCenterSubsModel
    {
        public long LocationID { get; set; }
        public List<long> Subs { get; set; }
    }
}