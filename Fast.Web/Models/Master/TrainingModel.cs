using System;

namespace Fast.Web.Models
{
	public class TrainingModel : BaseModel
	{		
		public string EmployeeID { get; set; }
		public string FullName { get; set; }
		public string Position { get; set; }
		public string Department { get; set; }
		public string BU { get; set; }
		public string BasetownLocation { get; set; }
		public string TrainingCategory { get; set; }
		public string TrainingTitle { get; set; }
		public Nullable<DateTime> StartDate { get; set; }
		public Nullable<DateTime> EndDate { get; set; }
		public string Trainer { get; set; }
		public string Score { get; set; }
		public string StatusTraining { get; set; }
		public Nullable<long> MachineTypeID { get; set; }
		public string MachineType { get; set; }
        public string TraineeList { get; set; }
    }
}