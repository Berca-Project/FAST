using System;

namespace Fast.Domain.Entities
{
	public class UserRole : BaseEntity
	{
		public long UserID { get; set; }		
		public string RoleName { get; set; }
	}
}
