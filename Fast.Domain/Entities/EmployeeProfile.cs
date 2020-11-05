using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Entities
{
	public class EmployeeProfile : BaseEntity
	{
		public string EmployeeID { get; set; }
		public string FullName { get; set; }
		public string PositionDesc { get; set; }
		public string DepartmentDesc { get; set; }
		public string HMSOrg3 { get; set; }
		public string BusinessUnit { get; set; }
		public string BaseTownLocation { get; set; }
		public string HomeTownLocation { get; set; }
		public string BaseTownCity { get; set; }
		public string Status { get; set; }
		public string Email { get; set; }
		public string EmployeeType { get; set; }
		public string GroupType { get; set; }
		public string GroupName { get; set; }
		public string ReportToID1 { get; set; }
		public string ReportToID2 { get; set; }
		public string CostCenter { get; set; }
	}
}
