using System;

namespace Fast.WebApi.Models
{
	public class BaseModel
	{
		public bool IsDeleted { get; set; }
		public string CreatedBy { get; set; }
		public Nullable<System.DateTime> CreatedDate { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<System.DateTime> ModifiedDate { get; set; }
	}
}