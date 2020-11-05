using System;

namespace Fast.Domain.Entities
{
	public class User : BaseEntity
	{
		public string UserName { get; set; }
		public string EmployeeID { get; set; }
		public long JobTitleID { get; set; }
		public Nullable<long> LocationID { get; set; }
		public Nullable<long> CanteenID { get; set; }
		public bool IsActive { get; set; }
		public bool IsOS { get; set; }
		public bool IsAdmin { get; set; }
		public bool IsFast { get; set; }

		public string Location { get; set; }
		public string SupervisorID { get; set; }
		public string SupervisorName { get; set; }
	}
}
