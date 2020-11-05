using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIDust : BaseEntity
	{
		public string ProductionCenter { get; set; }
		public int Dust { get; set; }
		public int Winnower { get; set; }
		public int FloorSweeping { get; set; }
		public int RS { get; set; }
		public decimal GRSpec { get; set; }
	}
}
