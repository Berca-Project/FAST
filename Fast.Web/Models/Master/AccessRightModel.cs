using System;
using System.Collections.Generic;

namespace Fast.Web.Models
{
	public class AccessRightModel : BaseModel
	{
		public string RoleName { get; set; }
		public long MenuID { get; set; }
		public string MenuName { get; set; }
		public long SubMenuID { get; set; }
		public string SubMenuName { get; set; }
		public string Description { get; set; }
		public Nullable<bool> Read { get; set; }
		public Nullable<bool> Write { get; set; }
		public Nullable<bool> Print { get; set; }
		public bool IsTopMenu { get; set; }
	}

	public class AccessRightDBModel : BaseModel
	{
		public string RoleName { get; set; }
		public long MenuID { get; set; }
		public long LocationID { get; set; }
		public string MenuName { get; set; }
		public Nullable<bool> Read { get; set; }
		public Nullable<bool> Write { get; set; }
		public Nullable<bool> Print { get; set; }
		public string ReadName { get; set; }
		public string WriteName { get; set; }
		public string PrintName { get; set; }

		#region Helper
		public long ProductionCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }
		public bool hasRead { get { return Read.HasValue && Read.Value; } }
		public bool hasWrite { get { return Write.HasValue && Write.Value; } }
		public bool hasPrint { get { return Print.HasValue && Print.Value; } }

		public string hasReadStr { get { return hasRead.ToString().ToLower(); } }
		public string hasWriteStr { get { return hasWrite.ToString().ToLower(); } }
		public string hasPrintStr { get { return hasPrint.ToString().ToLower(); } }
		public bool IsAdmin { get; set; }
		#endregion
	}

	public class AccessRightTreeModel : AccessRightDBModel
	{
		public List<ParentAccessRightModel> Parents { get; set; }
		public AccessRightTreeModel()
		{
			Parents = new List<ParentAccessRightModel>();
		}
	}

	public class ParentAccessRightModel : AccessRightDBModel
	{
		public List<ChildAccessRightModel> Children { get; set; }
		public ParentAccessRightModel()
		{
			Children = new List<ChildAccessRightModel>();
		}
	}

	public class ChildAccessRightModel : AccessRightDBModel
	{
	}
}