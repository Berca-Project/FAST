using System;

namespace Fast.Domain.Entities
{
	public class UserMachine : BaseEntity
	{
		public long UserID { get; set; }		
		public long MachineID { get; set; }
	}
}
