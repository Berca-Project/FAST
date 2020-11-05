namespace Fast.Web.Models
{
	public class UserMachineModel : BaseModel
	{
		public long UserID { get; set; }
		public long MachineID { get; set; }
		public long[] MachineIDs { get; set; }
		public string Machine { get; set; }
		public string MachineList { get; set; }
		public string EmployeeID { get; set; }
		public EmployeeModel Employee { get; set; }
		public string Location { get; set; }
	}
}