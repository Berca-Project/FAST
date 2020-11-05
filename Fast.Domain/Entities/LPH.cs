using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class LPH : BaseEntity
    {
        public string MenuTitle { get; set; }
        public string Header { get; set; }
        public string Type { get; set; }
        public long? LocationID { get; set; }

    }
}
