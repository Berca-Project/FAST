using System;

namespace Fast.Domain.Entities
{
	public class WppPrim : BaseEntity
	{
		public string Location { get; set; }		
		public string Blend { get; set; }
		public string BatchLama { get; set; }
		public DateTime Date { get; set; }		
		public string PONumber { get; set; }
		public string BatchSAP { get; set; }
        public string Others { get; set; }
        public long LocationID { get; set; }
    }
}
