using System;

namespace Fast.Domain.Entities
{
	public class EmployeeOvertime
	{
		public long ID { get; set; }
		public string EmployeeID { get; set; }
		public string DepartmentDesc { get; set; }
		public string PositionDesc { get; set; }
		public string BasetownLocation { get; set; }
		public string CostCenter { get; set; }
		public DateTime Date { get; set; }
		public string ClockIn { get; set; }
		public string ClockOut { get; set; }
		public string ActualIn { get; set; }
		public string ActualOut { get; set; }
		public decimal Overtime { get; set; }
		public string OvertimeCategory { get; set; }
		public string Comments { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }
    }
}
