using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class PPLPHSubmissions : BaseEntity
    {
        public long LPHID { get; set; }
        public string LPHHeader { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public int? SubShift { get; set; }
        public long UserID { get; set; }
        public long LocationID { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }

        public string Machine { get; set; }
        public string UserFullName { get; set; }
        public string Location { get; set; }
        public Boolean IsComplete { get; set; }

        public int? Flag { get; set; }

    }
}
