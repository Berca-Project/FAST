namespace Fast.Web.Models
{
	public class ReferenceDetailModel : BaseModel
	{
		public long ReferenceID { get; set; }
		public long ParentID { get; set; }
		public string Code { get; set; }
        public string Description { get; set; } = string.Empty;
	}
}