using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models
{
	public class ActiveDirectoryModel
	{
		public string EmployeeID { get; set; }
		public string Department { get; set; }
		public string Company { get; set; }
		public string DisplayName { get; set; }
		public string Title { get; set; }
		public string Email { get; set; }
		public string Phone { get; set; }
		public string Manager { get; set; }
	}
}