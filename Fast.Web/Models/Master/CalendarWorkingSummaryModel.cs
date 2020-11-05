namespace Fast.Web.Models
{
	public class CalendarWorkingSummaryModel
	{
		public string ColumnName { get; set; }
		public int Days { get; set; }
		public int Holiday { get; set; }
		public int Leaves { get; set; }
		public int ProdOff { get; set; }
		public int ShiftOff { get; set; }
		public double WorkDays { get; set; }
	}
}