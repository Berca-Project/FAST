using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class ReportRemarks: BaseEntity
    {
        public string RSA { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public long LocationID { get; set; }
        public string Focus { get; set; }
        public string ActionPlan { get; set; }
        public long UserID { get; set; }
       
    }
}
