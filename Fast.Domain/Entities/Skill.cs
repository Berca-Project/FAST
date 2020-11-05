namespace Fast.Domain.Entities
{
	public class Skill : BaseEntity
	{
		public string Code { get; set; }
		public string SubProcess { get; set; }
		public string Description { get; set; }
		public long LocationID { get; set; }
	}
}
