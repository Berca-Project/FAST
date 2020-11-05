using System;
using System.ComponentModel.DataAnnotations;

namespace Fast.Web.Models
{
	public class CalendarHolidayModel : BaseModel
	{
		public string Description { get; set; }
		public string Color { get; set; }
		public long LocationID { get; set; }
		public long HolidayTypeID { get; set; }
		public string HolidayType { get; set; }
		public string Location { get; set; }
		public DateTime Date { get; set; }
		public string DateStr { get; set; }
		public long CountryID { get; set; }
		public long ProdCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }
		public int Year { get; set; }

	}
}