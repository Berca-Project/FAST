using System;
using System.Web;

namespace Fast.Web.Models
{
	public class WppModel : BaseModel
	{
		public string Location { get; set; }
		public string Brand { get; set; }
		public string Description { get; set; }
		public string Packer { get; set; }
		public string Maker { get; set; }
		public DateTime Date { get; set; }
		public string Activity { get; set; }
		public HttpPostedFileBase PostedFilename { get; set; }
		public int Week { get; set; }
		public int Year { get; set; }
		public string PONumber { get; set; }
		public string OPSNumber { get; set; }
		public string BatchSAP { get; set; }
		public decimal Shift1 { get; set; }
		public decimal Shift2 { get; set; }
		public decimal Shift3 { get; set; }
		public decimal Actual1 { get; set; }
		public decimal Actual2 { get; set; }
		public decimal Actual3 { get; set; }
		public long LinkUpID { get; set; }
		public long BrandID { get; set; }
		public long ActivityID { get; set; }
		public string Others { get; set; }

		public long CountryID { get; set; }
		public long ProdCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }
		public long LocationID { get; set; }
	}

	public class WppDBModel : BaseModel
	{
		public string Location { get; set; }
		public string Brand { get; set; }
		public string Description { get; set; }
		public string Packer { get; set; }
		public string Maker { get; set; }
		public DateTime Date { get; set; }
		public string Activity { get; set; }
		public string PONumber { get; set; }
		public string OPSNumber { get; set; }
		public string BatchSAP { get; set; }
		public decimal Shift1 { get; set; }
		public decimal Shift2 { get; set; }
		public decimal Shift3 { get; set; }
		public string Others { get; set; }
	}

	public class WppSimpleModel
	{
		public string Location { get; set; }
		public string Brand { get; set; }
		public string Description { get; set; }
		public string Packer { get; set; }
		public string Maker { get; set; }
		public long LocationID { get; set; }
	}
}