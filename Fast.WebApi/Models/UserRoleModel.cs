namespace Fast.WebApi.Models
{
	public class UserRoleModel : BaseModel
	{
		public long ID { get; set; }
		public string Name { get; set; }
		public string DisplayName { get; set; }
		public string Description { get; set; }
		public bool IsDefault { get; set; }
	}
}