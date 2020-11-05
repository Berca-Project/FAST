using System;

namespace Fast.Domain.Entities
{
	public class CalendarHoliday : BaseEntity
	{
		public string Description { get; set; }
		public string Color { get; set; }
		public long LocationID { get; set; }
        public long HolidayTypeID { get; set; }
        public DateTime Date { get; set; }
	}
}
