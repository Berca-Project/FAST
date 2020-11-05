using System;

namespace Fast.Domain.Entities
{
	public class UserLog
	{
		public long ID { get; set; }
		public int UserID { get; set; }
        public string UserName { get; set; }
        public string Level { get; set; }
		public string Message { get; set; }
		public string Logger { get; set; }
		public DateTime Timestamp { get; set; }
		public string Exception { get; set; }
		public string Stacktrace { get; set; }		
	}
}
