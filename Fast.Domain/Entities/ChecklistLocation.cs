namespace Fast.Domain.Entities
{
	public class ChecklistLocation : BaseEntity
	{
        public long ChecklistID { get; set; }
        public long LocationID { get; set; }
    }
}
