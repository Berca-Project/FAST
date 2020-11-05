using System;

namespace Fast.Domain.Entities
{
	public class ChecklistValueHistory : BaseEntity
	{
        public long ChecklistSubmitID { get; set; }
        public long ChecklistValueID { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long UserID { get; set; }
    }
}
