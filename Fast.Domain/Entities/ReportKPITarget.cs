using System;

namespace Fast.Domain.Entities
{
	public class ReportKPITarget : BaseEntity
	{
		public string ProductionCenter { get; set; }
		public string KPI { get; set; }
		public decimal TargetInternal { get; set; }
		public decimal TargetOB { get; set; }
	}
}
