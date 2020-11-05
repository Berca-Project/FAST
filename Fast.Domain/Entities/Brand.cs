using System;

namespace Fast.Domain.Entities
{
	public class Brand
	{
		public long ID { get; set; }
		public string Code { get; set; }
		public string Description { get; set; }
		public Nullable<long> LocationID { get; set; }
		public Nullable<long> DeptID { get; set; }
		public Nullable<long> PcID { get; set; }
		public Int32 PackToStick { get; set; }
		public Int32 SlofToPack { get; set; }
		public Int32 BoxToSlof { get; set; }
		public Nullable<double> BeratCigarette { get; set; }
		public Nullable<double> CTW { get; set; }
		public bool IsActive { get; set; }
		public string ModifiedBy { get; set; }
		public Nullable<DateTime> ModifiedDate { get; set; }
        public Nullable<double> CTF { get; set; }
		public string RSCode { get; set; }
	}
}
