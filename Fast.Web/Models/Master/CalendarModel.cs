using System;
using System.Web;

namespace Fast.Web.Models
{
    public class CalendarModel : BaseModel
    {
        public string Shift1 { get; set; }
        public string Shift2 { get; set; }
        public string Shift3 { get; set; }
        public long LocationID { get; set; }
		public string LocationType { get; set; }
		public long GroupTypeID { get; set; }
        public string GroupType { get; set; }
        public string Location { get; set; }
        public DateTime Date { get; set; }
        public HttpPostedFileBase PostedFilename { get; set; }

		public long CountryID { get; set; }
		public long ProdCenterID { get; set; }
		public long DepartmentID { get; set; }
		public long SubDepartmentID { get; set; }
        public int Year { get; set; }

		public long PcID { get; set; }
		public long DepID { get; set; }
		public long SubDepID { get; set; }

        public string Week { get; set; }
	}
}