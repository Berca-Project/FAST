namespace Fast.Web.Models
{
	public class EmployeeModel : BaseModel
	{
		public string UserName { get; set; }
		public string Phone { get; set; }
		public string EmployeeID { get; set; }
		public string FullName { get; set; }
		public string PositionDesc { get; set; }
		public string DepartmentDesc { get; set; }
		public string HMSOrg3 { get; set; }
		public string BusinessUnit { get; set; }
		public string BaseTownLocation { get; set; }
		public string HomeTownLocation { get; set; }
		public string BaseTownCity { get; set; }
		public string ReportToID1 { get; set; }
		public string ReportToID2 { get; set; }
		public string Status { get; set; }
		public string Email { get; set; }
		public string EmployeeType { get; set; }
		public string GroupType { get; set; }
		public string GroupName { get; set; }
		public string Location { get; set; }
		public string CostCenter { get; set; }
		public string Canteen { get; set; }
        public string MachineReferenceList { get; set; }
        public string OS { get; set; }
    }

    public class EmployeeRoleModel
	{
		public string RoleName { get; set; }
		public EmployeeModel EmployeeModel { get; set; }
	}
}