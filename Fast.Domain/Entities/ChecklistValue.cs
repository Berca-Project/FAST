using System;

namespace Fast.Domain.Entities
{
	public class ChecklistValue : BaseEntity
	{
        public long ChecklistSubmitID { get; set; }
        public long ChecklistComponentID { get; set; }
        public string Value { get; set; }
        public string ValueType { get; set; }
        public DateTime? ValueDate { get; set; }
        public int OrderNum { get; set; }
        public int ColumnNum { get; set; }
    }
}
