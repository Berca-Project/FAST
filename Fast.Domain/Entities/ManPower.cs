namespace Fast.Domain.Entities
{
	public class ManPower : BaseEntity
	{
		public long JobTitleID { get; set; }
        public string RoleName { get; set; }
        public long LocationID { get; set; }
		public decimal Value { get; set; }
	}
}
