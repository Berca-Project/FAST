using System;

namespace Fast.Domain.Entities
{
	public class Location : BaseEntity
	{
		public string Code { get; set; }		
		public long ParentID { get; set; }
		public string ParentCode { get; set; }
    }
}
