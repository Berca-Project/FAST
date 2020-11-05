namespace Fast.Web.Models
{
	public class SkillModel : BaseModel
	{
		public string Code { get; set; }
		public string SubProcess { get; set; }
		public string Description { get; set; }
		public long LocationID { get; set; }
		public string Location { get; set; }
	}
}