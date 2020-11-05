namespace Fast.Web.Models
{
    public class MppChangesModel: BaseModel
    {
        public long MPPID { get; set; }
        public string FieldName { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public string DataType { get; set; }
    }
}