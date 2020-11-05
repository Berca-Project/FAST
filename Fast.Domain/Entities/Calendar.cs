using System;

namespace Fast.Domain.Entities
{
	public class Calendar : BaseEntity
	{
		public string Shift1 { get; set; }
		public string Shift2 { get; set; }
		public string Shift3 { get; set; }
		public long LocationID { get; set; }
        public string Location { get; set; }
        public long GroupTypeID { get; set; }
        public DateTime Date { get; set; }
	}
}
