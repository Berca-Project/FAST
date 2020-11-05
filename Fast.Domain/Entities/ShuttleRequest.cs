using System;

namespace Fast.Domain.Entities
{
	public class ShuttleRequest : BaseEntity
	{
		public string CostCenter { get; set; }
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public DateTime Date { get; set; }
		public TimeSpan Time { get; set; }
		public string ProductionCenter { get; set; }
		public string EmployeeID { get; set; }
		public int TotalPassengers { get; set; }
		public string GuestType { get; set; }
		public string LocationFrom { get; set; }
		public string LocationTo { get; set; }
		public string Purpose { get; set; }
		public string Department { get; set; }
		public string Phone { get; set; }
	}
}

