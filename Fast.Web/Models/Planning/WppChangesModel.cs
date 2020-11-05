namespace Fast.Web.Models
{
    public class WppChangesModel: BaseModel
    {
        public long WPPID { get; set; }
        public string FieldName { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public string DataType { get; set; }
    }
}