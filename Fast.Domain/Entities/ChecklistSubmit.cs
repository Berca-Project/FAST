using System;
using System.Collections.Generic;

namespace Fast.Domain.Entities
{
	public class ChecklistSubmit : BaseEntity
	{
        public long ChecklistID { get; set; }
        public DateTime CompleteDate { get; set; }
        public long UserID { get; set; }
        public string Shift { get; set; }
        public DateTime date { get; set; }
        public bool IsEdited { get; set; }
        public bool IsComplete { get; set; }
    }
}
