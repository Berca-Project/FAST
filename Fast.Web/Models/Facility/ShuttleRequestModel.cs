using System;

namespace Fast.Web.Models
{
	public class ShuttleRequestModel : BaseModel
	{
		public string CostCenter { get; set; }
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public DateTime Date { get; set; }
		public TimeSpan Time { get; set; }
		public string ProductionCenterID { get; set; }
		public string ProductionCenter { get; set; }
		public string EmployeeID { get; set; }
		public string EmployeeFullname { get; set; }
		public int TotalPassengers { get; set; }
		public string GuestType { get; set; }
		public string LocationFrom { get; set; }
		public string LocationTo { get; set; }
		public string Purpose { get; set; }
		public string Department { get; set; }
		public string Phone { get; set; }
		public string RequestType { get; set; }
		public int Qty { get; set; }
        public bool IsMPP { get; set; }
    }
}