using System;

namespace Fast.Domain.Entities
{
    public class EmployeeLeave 
    {
        public long ID { get; set; }
        public string EmployeeID { get; set; }
        public string LeaveType { get; set; }
        public string StartDateHalfDay { get; set; }
        public string StartDatePagiSiang { get; set; }
		public string EndDateHalfDay { get; set; }
		public string EndDatePagiSiang { get; set; }
        public Nullable<DateTime> StartDate { get; set; }
        public Nullable<DateTime> EndDate { get; set; }
        public string AllDayHalf { get; set; }
        public string Comments { get; set; }
        public string Status { get; set; }
        public string EmployeeType { get; set; }
        public Nullable<DateTime> TimesheetLastModified { get; set; }
        public Nullable<long> LocationID { get; set; }
        public string Location { get; set; }

    }
}
