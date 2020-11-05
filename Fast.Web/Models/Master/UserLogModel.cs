using System;

namespace Fast.Web.Models
{
	public class UserLogModel : BaseModel
	{		
		public int UserID { get; set; }
        public string UserName { get; set; }
        public string Level { get; set; }
		public string Message { get; set; }
		public string Logger { get; set; }
		public DateTime Timestamp { get; set; }
		public Nullable<DateTime> DeletedDate { get; set; }
		public string Exception { get; set; }
		public string Stacktrace { get; set; }
	}
}