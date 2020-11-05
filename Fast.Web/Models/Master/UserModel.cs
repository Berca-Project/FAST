using System;
using System.Collections.Generic;
using System.Web;

namespace Fast.Web.Models
{
    public class UserModel : BaseModel
    {
        public UserModel()
        {
            Employee = new EmployeeModel();
            Access = new AccessRightDBModel();
            RoleNames = new List<string>();
        }

        public string UserName { get; set; }
        public string Password { get; set; }
        public string Email { get; set; }
        public string EmployeeID { get; set; }
        public long JobTitleID { get; set; }
        public string JobTitle { get; set; }
        public string RoleName { get; set; }
        public List<string> RoleNames { get; set; }
        public long? LocationID { get; set; }
        public long? CanteenID { get; set; }
        public EmployeeModel Employee { get; set; }

        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }

        public string GroupType { get; set; }

        public string GroupName { get; set; }

        public bool IsOS { get; set; }
        public bool IsAdmin { get; set; }
        public bool IsFast { get; set; }
        public bool IsHasExtraRole { get; set; }
        public string FullName { get; set; }

        public string Location { get; set; }

        public string SupervisorName { get; set; }
        public string ManagerName { get; set; }
        public string SupervisorID { get; set; }
        public long SupervisorUserID { get; set; }
        public string SupervisorEmail { get; set; }
        public string SearchBy { get; set; }
        public string Canteen { get; set; }

        public string ComputerName { get; set; }
        public string IpAddress { get; set; }

    }

    public class ExtraRoleModel
    {
        public string Role { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime ModifiedDate { get; set; }
    }
}