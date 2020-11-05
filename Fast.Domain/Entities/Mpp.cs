using System;

namespace Fast.Domain.Entities
{
	public class Mpp : BaseEntity
	{			
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public int Year { get; set; }
		public int Week { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public string StatusMPP { get; set; }
        public string JobTitle { get; set; }
		public string EmployeeID { get; set; }
		public string EmployeeName { get; set; }
		public string EmployeeMachine { get; set; }
		public long WPPID { get; set; }		
		public long LocationID { get; set; }
		public string Location { get; set; }
		public string Remark { get; set; }
		public string GroupType { get; set; }
		public string GroupName { get; set; }
	}
}
