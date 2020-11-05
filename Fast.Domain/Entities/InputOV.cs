using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class InputOV : BaseEntity
    {
        public string WasteCategory { get; set; }
        public string Week { get; set; }
        public string OV { get; set; }
        public long LocationID { get; set; }
    }
}
