using System;

namespace Fast.Domain.Entities
{
	public class MealRequest : BaseEntity
	{
		public string CostCenter { get; set; }
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public DateTime Date { get; set; }
		public string ProductionCenter { get; set; }
		public string Canteen { get; set; }
		public string EmployeeID { get; set; }
		public int TotalGuest { get; set; }
		public string GuestType { get; set; }
		public string Guest { get; set; }
		public string Company { get; set; }
		public string Purpose { get; set; }
		public string Department { get; set; }
		public string Shift { get; set; }
		public string Phone { get; set; }
	}
}

