using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class InputDailyModel : BaseModel
    {
        public long ProdCenterID { get; set; }
        public string Shift { get; set; }
        public string LinkUp { get; set; }
        public DateTime Date { get; set; }
        public string Week { get; set; }
        public decimal MTBFValue { get; set; }
        public string MTBFFocus { get; set; }
        public string MTBFActPlan { get; set; }
        public decimal CPQIValue { get; set; }
        public string CPQIFocus { get; set; }
        public string CPQIActPlan { get; set; }
        public decimal VQIValue { get; set; }
        public string VQIFocus { get; set; }
        public string VQIActPlan { get; set; }
        public decimal WorkingValue { get; set; }
        public string WorkingFocus { get; set; }
        public string WorkingActPlan { get; set; }
        public decimal UptimeValue { get; set; }
        public string UptimeFocus { get; set; }
        public string UptimeActPlan { get; set; }
        public decimal STRSValue { get; set; }
        public string STRSFocus { get; set; }
        public string STRSActPlan { get; set; }
        public decimal ProdVolumeValue { get; set; }
        public string ProdVolumeFocus { get; set; }
        public string ProdVolumeActPlan { get; set; }
        public decimal CRRValue { get; set; }
        public string CRRFocus { get; set; }
        public string CRRActPlan { get; set; }
        public long LocationID { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }
        public bool IsShift1 { get; set; }
        public bool IsShift2 { get; set; }
        public bool IsShift3 { get; set; }
        public bool IsAllShift { get; set; }
        public bool IsDaily { get; set; }
        public bool IsAllKPI { get; set; }
        public bool IsVQI { get; set; }
        public bool IsCPQI { get; set; }
        public bool IsWorkingTime { get; set; }
        public bool IsMTBF { get; set; }        
        public bool IsSTRS { get; set; }
        public bool IsCRR { get; set; }
        public bool IsProdVolume { get; set; }
        public bool IsUptime { get; set; }

        public string ProductionCenter { get; set; }
        public long CountryID { get; set; }        
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
        public string Focus { get; set; }
        public string ActionPlan { get; set; }
        

    }
}