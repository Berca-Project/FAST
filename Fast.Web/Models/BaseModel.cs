using System;

namespace Fast.Web.Models
{
	public abstract class BaseModel
	{
		public AccessRightDBModel Access { get; set; }
		public long ID { get; set; }
		public bool IsDeleted { get; set; }
		public bool IsActive { get; set; } 		
		public string ModifiedBy { get; set; }		
		public Nullable<DateTime> ModifiedDate { get; set; }
	}
}