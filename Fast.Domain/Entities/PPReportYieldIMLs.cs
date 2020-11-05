using System;

namespace Fast.Domain.Entities
{
	public class PPReportYieldIMLs : BaseEntity
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long? LocationID { get; set; }
	    public string Type { get; set; }
	    public decimal? Value { get; set; }
    }
}
