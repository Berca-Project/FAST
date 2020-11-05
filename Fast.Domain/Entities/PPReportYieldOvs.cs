using System;

namespace Fast.Domain.Entities
{
	public class PPReportYieldOvs : BaseEntity
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long? LocationID { get; set; }
	    public string Type { get; set; }
	    public string Waste { get; set; }
	    public string WasteType { get; set; }
	    public string Area { get; set; }
	    public decimal? OvValue { get; set; }
    }
}
