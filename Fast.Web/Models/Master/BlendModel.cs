using System;

namespace Fast.Web.Models
{
    public class BlendModel : BaseModel
    {
        public string Code { get; set; }
        public string Description { get; set; }
		public Nullable<long> LocationID { get; set; }
		public string Location { get; set; }
		public Nullable<long> DeptID { get; set; }
		public string Department { get; set; }
		public Nullable<long> PcID { get; set; }
		public string ProductionCenter { get; set; }
		public Nullable<long> CountryID { get; set; }
		public float OpsToKg { get; set; }
	}
}