using System;

namespace Fast.Domain.Entities
{
	public class PPReportYieldTargets : BaseEntity
    {
        public long? LocationID { get; set; }
	    public string Type { get; set; }
	    public string Name { get; set; }
	    public decimal? Target { get; set; }
    }
}
