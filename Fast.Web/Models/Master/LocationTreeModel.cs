using System.Collections.Generic;

namespace Fast.Web.Models
{
	public class LocationTreeModel : LocationModel
	{				
		public List<ProductionCenterModel> ProductionCenters { get; set; }
		public LocationTreeModel()
		{
			ProductionCenters = new List<ProductionCenterModel>();
		}
	}

	public class ProductionCenterModel : ReferenceDetailModel
	{
		public long LocationID { get; set; }
		public int Index { get; set; }
		public List<DepartmentModel> Departments { get; set; }
		public ProductionCenterModel()
		{
			Departments = new List<DepartmentModel>();
		}
	}

	public class DepartmentModel : ReferenceDetailModel
	{
		public int Index { get; set; }
		public long LocationID { get; set; }
		public List<SubDepartmentModel> SubDepartments { get; set; }
		public DepartmentModel()
		{
			SubDepartments = new List<SubDepartmentModel>();
		}
	}

	public class SubDepartmentModel : ReferenceDetailModel
	{
		public int Index { get; set; }
		public long LocationID { get; set; }
	}
}