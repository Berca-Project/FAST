using System;

namespace Fast.Domain.Entities
{
	public class Machine : BaseEntity
	{
		public string Code { get; set; }
		public string Items { get; set; }
		public string LegalEntity { get; set; }
		public string MachineBrand { get; set; }
		public long MachineTypeID { get; set; }
		public Nullable<long> LocationID { get; set; }
		public string MachineSN { get; set; }
		public string SubProcess { get; set; }
		public string SamID { get; set; }
		public string Notes { get; set; }
		public string LinkUp { get; set; }
		public string Location { get; set; }
		public string Cluster { get; set; }
		public Nullable<int> OrderNumber { get; set; }
		public bool IsActive { get; set; }
		public Nullable<decimal> DesignSpeed { get; set; }
		public Nullable<decimal> CellophanerSpeed { get; set; }
	}
}
