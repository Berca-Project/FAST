using System;

namespace Fast.Domain.Entities
{
	public class Blend 
	{
		public long ID { get; set; }		
		public string Code { get; set; }
		public string Description { get; set; }
		public Nullable<long> LocationID { get; set; }
		public Nullable<long> DeptID { get; set; }
		public Nullable<long> PcID { get; set; }
		public bool IsActive { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<DateTime> ModifiedDate { get; set; }
        public Nullable<double> OpsToKg { get; set; }
	}
}
