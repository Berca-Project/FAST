using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
    public class LPHLocations : BaseEntity
    {
        public long LPHID { get; set; }
        public long LocationID { get; set; }

    }
}
