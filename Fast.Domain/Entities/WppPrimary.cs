using System;

namespace Fast.Domain.Entities
{
	public class WppPrimary : BaseEntity
	{
		public string Blend { get; set; }
		public int VolPerOps { get; set; }
        public int Week { get; set; }
        public DateTime StartDate { get; set; }
		public int Monday { get; set; }
		public int Tuesday { get; set; }
		public int Wednesday { get; set; }
		public int Thursday { get; set; }
		public int Friday { get; set; }
		public int Saturday { get; set; }
		public int Sunday { get; set; }
		public long LocationID { get; set; }
		public string Location { get; set; }
	}
}
