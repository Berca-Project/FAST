using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class InputDaily: BaseEntity
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
      
        


    }
}
