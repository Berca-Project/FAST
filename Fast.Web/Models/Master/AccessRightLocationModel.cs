using System;
using System.Collections.Generic;

namespace Fast.Web.Models
{
	public class AccessRightLocationModel 
	{
		public long LocationID { get; set; }
		public List<AccessRightDBModel> AccessList { get; set; }
	}
}