using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class DailyModel
    {
        public string LinkUp { get; set; }
        public DateTime DateValue { get; set; }
        public string Shift { get; set; }
        public decimal MTBFValue { get; set; }
        public decimal CPQIValue { get; set; }
        public decimal VQIValue { get; set; }
        public decimal WorkingValue { get; set; }
        public decimal UptimeValue { get; set; }
        public decimal STRSValue { get; set; }
        public decimal ProdVolumeValue { get; set; }
        public decimal CRRValue { get; set; }        
    }
    public class TargetModel
    {
        public string KPI { get; set; }
        public string Version { get; set; }
        public decimal ValueTarget { get; set; }
    }

    public class UptimeModel
    {
        public string Shift { get; set; }
        public decimal UptimeValue { get; set; }
        public string UptimeFocus { get; set; }
        public string UptimeActPlan { get; set; }
    }
    public class WTDModel
    {
        public string MTBF_WTD { get; set; }
        public string CPQI_WTD { get; set; }
        public string VQI_WTD { get; set; }
        public string Working_WTD { get; set; }
        public string Uptime_WTD { get; set; }
        public string STRS_WTD { get; set; }
        public string ProdVolume_WTD { get; set; }
        public string CRR_WTD { get; set; }
    }
    public class MTDModel
    {
        public string MTBF_MTD { get; set; }
        public string CPQI_MTD { get; set; }
        public string VQI_MTD { get; set; }
        public string Working_MTD { get; set; }
        public string Uptime_MTD { get; set; }
        public string STRS_MTD { get; set; }
        public string ProdVolume_MTD { get; set; }
        public string CRR_MTD { get; set; }
    }
    public class MTGModel
    {
        public string MTBF_MTG { get; set; }
        public string CPQI_MTG { get; set; }
        public string VQI_MTG { get; set; }
        public string Working_MTG { get; set; }
        public string Uptime_MTG { get; set; }
        public string STRS_MTG { get; set; }
        public string ProdVolume_MTG { get; set; }
        public string CRR_MTG { get; set; }
    }
}