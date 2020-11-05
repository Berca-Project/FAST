using Fast.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Utils
{
	public static class TreeNodeExtensions
	{
		public static AccessRightDBModel ToAccessRightModel(this MenuModel value, string roleName)
		{
			AccessRightDBModel arModel = new AccessRightDBModel
			{
				MenuID = value.ID,
				MenuName = value.Name,
				RoleName = roleName
			};

			return arModel;
		}
	}
}