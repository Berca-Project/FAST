using System;

namespace Fast.Web.Models
{
	public class LocationModel : BaseModel
	{
		public string Code { get; set; }
		public string Description { get; set; }
		public long ParentID { get; set; }
		public string ParentCode { get; set; }

		public string ProductionCenterCode { get; set; }
		public string DepartmentCode { get; set; }
		public string SubDepartmentCode { get; set; }
		public string CountryCode { get; set; }
		public string Country { get; set; }
		public string LocationParentCode { get; set; }
		public string LocationParent { get; set; }
		public string LocationCode { get; set; }
		public string Location { get; set; }
		public int Type { get; set; }

		public string LocationFullCode
		{
			get
			{
				if (string.IsNullOrEmpty(LocationParentCode))
					return CountryCode + "-" + LocationCode;
				else
					return CountryCode + "-" + LocationParentCode + "-" + LocationCode;
			}
		}

		public string ProdCenterFullCode
		{
			get
			{
				return ParentCode + "-" + Code;
			}
		}

		public string DepartementFullCode
		{
			get
			{
				return CountryCode + "-" + ParentCode + "-" + Code;
			}
		}

		public string SubDepartementFullCode
		{
			get
			{
				return CountryCode + "-" + ProductionCenterCode + "-" + ParentCode + "-" + Code;
			}
		}
	}
}