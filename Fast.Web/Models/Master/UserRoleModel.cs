namespace Fast.Web.Models
{
	public class UserRoleModel : BaseModel
	{
		public long UserID { get; set; }
		public string RoleName { get; set; }
		public string[] RoleNames { get; set; }
		public string RoleList { get; set; }
		public string EmployeeID { get; set; }
		public string UserName { get; set; }
		public EmployeeModel Employee { get; set; }
		public long LocationID { get; set; }
		public string Location { get; set; }
	}
}