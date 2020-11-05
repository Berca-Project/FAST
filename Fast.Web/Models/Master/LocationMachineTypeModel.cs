namespace Fast.Web.Models
{
	public class LocationMachineTypeModel : BaseModel
	{
		public long LocationID { get; set; }
        public string Location { get; set; }
        public long MachineTypeID { get; set; }
        public string MachineType { get; set; }
        public long[] MachineTypeIDs { get; set; }

        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
        public long SubDepartmentID { get; set; }
    }
}