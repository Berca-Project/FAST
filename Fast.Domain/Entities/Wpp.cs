using System;

namespace Fast.Domain.Entities
{
	public class Wpp : BaseEntity
	{
		public string Location { get; set; }
		public string Brand { get; set; }
		public string Description { get; set; }
		public string Packer { get; set; }
		public string Maker { get; set; }
		public DateTime Date { get; set; }
		public decimal Shift1 { get; set; }
		public decimal Shift2 { get; set; }
		public decimal Shift3 { get; set; }
		public decimal Actual1 { get; set; }
		public decimal Actual2 { get; set; }
		public decimal Actual3 { get; set; }
		public string PONumber { get; set; }
		public string OPSNumber { get; set; }
		public string BatchSAP { get; set; }
		public string Activity { get; set; }
		public string Others { get; set; }
		public long LocationID { get; set; }
	}
}
