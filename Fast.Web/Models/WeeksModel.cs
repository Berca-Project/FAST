using System;

namespace Fast.Web.Models
{
    public class WeeksModel : BaseModel
    {        
        public int Year { get; set; }
        public int Week { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}