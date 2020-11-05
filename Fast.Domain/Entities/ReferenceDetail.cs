namespace Fast.Domain.Entities
{
	public class ReferenceDetail : BaseEntity
	{
		public long ReferenceID { get; set; }
		public string Code { get; set; }
		public string Description { get; set; }
	}
}
