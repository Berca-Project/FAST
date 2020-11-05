using System;

namespace Fast.Domain.Entities
{
	public class AccessRight : BaseEntity
	{
		public string RoleName { get; set; }
		public long MenuID { get; set; }
		public long LocationID { get; set; }
		public Nullable<bool> Read { get; set; }
		public Nullable<bool> Write { get; set; }
		public Nullable<bool> Print { get; set; }
	}
}
