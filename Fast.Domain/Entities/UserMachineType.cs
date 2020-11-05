using System;

namespace Fast.Domain.Entities
{
	public class UserMachineType : BaseEntity
	{
		public long UserID { get; set; }		
		public long MachineTypeID { get; set; }
	}
}
