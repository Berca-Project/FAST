namespace Fast.Web.Models
{
	public class ManPowerModel : BaseModel
	{
		public long JobTitleID { get; set; }
		public string JobTitle { get; set; }
        public string RoleName { get; set; }
        public long LocationID { get; set; }
		public string Location { get; set; }
		public decimal Value { get; set; }
        public long CountryID { get; set; }
        public long ProdCenterID { get; set; }
        public long DepartmentID { get; set; }
    }
}