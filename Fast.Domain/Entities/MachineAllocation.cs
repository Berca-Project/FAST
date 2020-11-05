using System;

namespace Fast.Domain.Entities
{
	public class MachineAllocation : BaseEntity
	{
		public long MachineID { get; set; }
		public string MachineCode { get; set; }
		public string MachineCategory { get; set; }
        public decimal Value { get; set; }
    }
}
