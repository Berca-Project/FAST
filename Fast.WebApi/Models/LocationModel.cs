using System;

namespace Fast.WebApi.Models
{
	public class LocationModel : BaseModel
	{
		public long ID { get; set; }
		public string Name { get; set; }
		public string Type { get; set; }
		public Nullable<short> Level { get; set; }
		public Nullable<long> ParentID { get; set; }
	}
}