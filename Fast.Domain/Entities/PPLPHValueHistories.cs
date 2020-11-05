using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class PPLPHValueHistories : BaseEntity
    {
        public long LPHValuesID { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long? UserID { get; set; }

    }
}
