using System;

namespace Fast.Web.Models
{
	public class BrandModel : BaseModel
	{
		public string Code { get; set; }
		public string Description { get; set; }
		public Nullable<long> LocationID { get; set; }
		public string Location { get; set; }
		public Nullable<long> DeptID { get; set; }
		public string Department { get; set; }
		public Nullable<long> PcID { get; set; }
		public string ProductionCenter { get; set; }
		public Nullable<long> CountryID { get; set; }
		public Int32 PackToStick { get; set; }
		public Int32 SlofToPack { get; set; }
		public Int32 BoxToSlof { get; set; }
		public Nullable<double> BeratCigarette { get; set; }
		public Nullable<double> CTW { get; set; }
        public Nullable<double> CTF { get; set; }
		public string RSCode { get; set; }
	}
}