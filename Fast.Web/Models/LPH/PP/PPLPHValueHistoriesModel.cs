namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHValueHistoriesModel : BaseModel
    {
        public long LPHValuesID { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long? UserID { get; set; }
    }
}