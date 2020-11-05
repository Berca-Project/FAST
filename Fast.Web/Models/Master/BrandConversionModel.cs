namespace Fast.Web.Models
{
	public class BrandConversionModel : BaseModel
    {		
		public string BrandCode { get; set; }		
		public decimal Value1 { get; set; }
		public string UOM1 { get; set; }		
		public decimal Value2 { get; set; }
		public string UOM2 { get; set; }
		public string Notes { get; set; }
	}
}