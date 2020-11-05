using System;

namespace Fast.Domain.Entities
{
	public class Menu : BaseEntity
	{
		public int DisplayOrder { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public string PageIcon { get; set; }
		public string PageSlug { get; set; }
		public string PageController { get; set; }
		public string PageAction { get; set; }
		public bool IsActive { get; set; }
		public bool IsParent { get; set; }
		public bool IsTopMenu { get; set; }
		public Nullable<long> ParentID { get; set; }
	}
}
