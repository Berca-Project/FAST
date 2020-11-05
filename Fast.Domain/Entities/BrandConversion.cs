using System;

namespace Fast.Domain.Entities
{
	public class BrandConversion
	{
		public long ID { get; set; }
		public string BrandCode { get; set; }
		public decimal Value1 { get; set; }
		public string UOM1 { get; set; }
		public decimal Value2 { get; set; }
		public string UOM2 { get; set; }
		public string Notes { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<DateTime> ModifiedDate { get; set; }
	}
}
