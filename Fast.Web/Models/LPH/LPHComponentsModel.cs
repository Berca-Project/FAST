namespace Fast.Web.Models.LPH
{
    public class LPHComponentsModel : BaseModel
    {
        public long LPHID { get; set; }
        public string ComponentType { get; set; }
        public string ComponentName { get; set; }
        public string AdditionalClass { get; set; }
        public string ImagePath { get; set; }
        public string Note { get; set; }
        public int OrderNum { get; set; }
        public long? Parent { get; set; }
    }
}