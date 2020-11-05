using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class InputTarget: BaseEntity
    {
        public string KPI { get; set; }
        public decimal Value { get; set; }
        public string Month { get; set; }
        public string Version { get; set; }
        public long ProdCenterID { get; set; }

    }
}
