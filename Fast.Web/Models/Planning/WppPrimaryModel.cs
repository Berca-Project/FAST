using System;
using System.Web;

namespace Fast.Web.Models
{
	public class WppPrimaryModel : BaseModel
	{
		public string Blend { get; set; }
		public int VolPerOps { get; set; }
		public DateTime StartDate { get; set; }
		public long LocationID { get; set; }
		public string Location { get; set; }

		public int Monday { get; set; }
		public int Tuesday { get; set; }
		public int Wednesday { get; set; }
		public int Thursday { get; set; }
		public int Friday { get; set; }
		public int Saturday { get; set; }
		public int Sunday { get; set; }
		public HttpPostedFileBase PostedFilename { get; set; }
		public int Total { get; set; }

        public int Week { get; set; }
        public int Year { get; set; }
        public long CountryID { get; set; }
		public long ProdCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }
	}
}