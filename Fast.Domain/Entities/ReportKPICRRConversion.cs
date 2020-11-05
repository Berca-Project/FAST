using System;

namespace Fast.Domain.Entities
{
	public class ReportKPICRRConversion : BaseEntity
	{
		public int Year { get; set; }
		public int Week { get; set; }
        public string ProductionCenter { get; set; }
        public decimal DebuFinal { get; set; }
		public decimal DustHalusLM { get; set; }
		public decimal DustHalusLMnoRipper { get; set; }
		public decimal SaponTembakau { get; set; }
		public decimal SaponCigarette { get; set; }
		public decimal ClaimableInThStick { get; set; }
		public decimal AverageFinalOv { get; set; }
		public decimal C2Weight { get; set; }
    }
}
