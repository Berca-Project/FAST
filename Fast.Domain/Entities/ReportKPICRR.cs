using System;

namespace Fast.Domain.Entities
{
	public class ReportKPICRR : BaseEntity
	{
		public int Year { get; set; }
		public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public DateTime PostingDate { get; set; }
		public string Machine { get; set; }
		public string MachineType { get; set; }
		public int Shift { get; set; }
		public string OrderNumber { get; set; }
		public string POLeadMaterial { get; set; }
		public string RejectMaterial { get; set; }
		public string MaterialDescription { get; set; }
		public decimal Quantity { get; set; }
		public string BaseUnit { get; set; }
	}
}
