using System;
using System.ComponentModel.DataAnnotations;

namespace Fast.Domain.Entities
{
	public class Role
	{
		[Key]
		public string Name { get; set; }		
		public string Description { get; set; }
		public bool IsDeleted { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<DateTime> ModifiedDate { get; set; }
	}
}
