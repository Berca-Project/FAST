using System.Collections.Generic;

namespace Fast.Domain.Entities
{
	public class Checklist : BaseEntity
	{
		public Checklist()
		{
			Locations = new HashSet<ChecklistLocation>();
		}
        public string CreatorEmployeeID { get; set; }
        public string MenuTitle { get; set; }
        public string Header { get; set; }
        public string Icon { get; set; }
        public int ColumnHeader { get; set; }
        public int ColumnContent { get; set; }
        public int FrequencyAmount { get; set; }
        public int FrequencyDivider { get; set; }
        public string FrequencyUnit { get; set; }
        public virtual ICollection<ChecklistLocation> Locations { get; set; }
	}
}
