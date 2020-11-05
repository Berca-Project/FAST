using System;

namespace Fast.Domain.Entities
{
	public class BaseEntity
	{
		public long ID { get; set; }
		public bool IsDeleted { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<DateTime> ModifiedDate { get; set; }
	}
}
