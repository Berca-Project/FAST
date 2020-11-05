using System;

namespace Fast.Domain.Entities
{
	public class ReportKPITobaccoWeight : BaseEntity
	{
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public decimal Mean { get; set; }
        public decimal Stdev { get; set; }
        public decimal MeanMC { get; set; }
        public decimal StdevMC { get; set; }
    }
}
