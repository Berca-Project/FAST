using System;

namespace Fast.Domain.Entities
{
	public class PPReportYields : BaseEntity
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public Nullable<double> Input { get; set; }
	    public Nullable<double> MCInput { get; set; }
	    public Nullable<double> DryInput { get; set; }
	    public Nullable<double> NonBom { get; set; }
	    public Nullable<double> MCNonBom { get; set; }
	    public Nullable<double> DryNonBom { get; set; }
	    public Nullable<double> Casing { get; set; }
	    public Nullable<double> DryMatter { get; set; }
	    public Nullable<double> DryCasing { get; set; }
	    public Nullable<double> Output { get; set; }
	    public Nullable<double> MCOutput { get; set; }
	    public Nullable<double> DryOutput { get; set; }
	    public Nullable<double> DryYield { get; set; }
	    public Nullable<double> WetYield { get; set; }
    }
}
