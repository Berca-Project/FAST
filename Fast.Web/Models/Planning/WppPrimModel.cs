using System;
using System.Web;

namespace Fast.Web.Models
{
	public class WppPrimModel : BaseModel
	{
		public string Location { get; set; }
		public string Blend { get; set; }
		public string BatchLama { get; set; }
		public DateTime Date { get; set; }
		public string PONumber { get; set; }
		public string BatchSAP { get; set; }
		public string Others { get; set; }
		public long LocationID { get; set; }		
		public string Activity { get; set; }		
		public HttpPostedFileBase PostedFilename { get; set; }
		public int Week { get; set; }
		public int Year { get; set; }
		public int EndWeek { get; set; }
		public int EndYear { get; set; }
		public long LinkUpID { get; set; }
		public long BlendID { get; set; }
		public long ActivityID { get; set; }		        

		public long CountryID { get; set; }
		public long ProdCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }        
    }

	public class WppPrimDBModel : BaseModel
	{
		public string Location { get; set; }
		public string Blend { get; set; }
		public string BatchLama { get; set; }
		public DateTime Date { get; set; }
		public string PONumber { get; set; }
		public string BatchSAP { get; set; }
		public string Others { get; set; }
		public long LocationID { get; set; }
	}
}