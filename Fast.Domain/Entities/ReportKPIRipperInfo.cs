using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIRipperInfo : BaseEntity
	{
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string Material { get; set; }
        public string OrderNum { get; set; }
        public DateTime ActualStartDate { get; set; }
        public string Description { get; set; }
        public decimal QtyIss { get; set; }
        public decimal QtyRec { get; set; }
        public decimal Yield { get; set; }
        public decimal ValIssued { get; set; }
        public decimal ValReceiv { get; set; }
        public decimal ValDiffer { get; set; }
    }
}
