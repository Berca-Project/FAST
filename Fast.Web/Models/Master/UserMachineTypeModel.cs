namespace Fast.Web.Models
{
	public class UserMachineTypeModel : BaseModel
	{
		public long UserID { get; set; }
		public long MachineTypeID { get; set; }
		public long[] MachineTypeIDs { get; set; }
		public string UserName { get; set; }
		public string FullName { get; set; }
		public string EmployeeID { get; set; }
		public string PositionDesc { get; set; }
		public string MachineType { get; set; }
		public string Skills { get; set; }
        public string Competency { get; set; }
    }
}