using System;

namespace Fast.Domain.Entities
{
	public class ReportKPIStickPerPack : BaseEntity
	{
        public string ProductionCenter { get; set; }
        public string Packer { get; set; }
        public int Stick { get; set; }
    }
}
