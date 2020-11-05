using System;
using System.Collections.Generic;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class PPReportYieldGovernanceModel : BaseModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public Dictionary<string, double> Items { get; set; }
        //public double DryYield { get; set; }
        //public double WetYield { get; set; }
        //public double DryWaste { get; set; }
        //public double WetWaste { get; set; }
        //public double DustWaste { get; set; }
        //public double InfeedMassLoad { get; set; }
        //public double Unaccountable { get; set; }
    }
}