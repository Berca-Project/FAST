using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIWorkHour : BaseEntity
	{
        public int Year { get; set; }
        public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public string Packer { get; set; }
        public decimal WorkHour { get; set; }
    }
}
