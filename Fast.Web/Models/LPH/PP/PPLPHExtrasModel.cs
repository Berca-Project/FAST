using System;

namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHExtrasModel : BaseModel
    {
        public long LPHID { get; set; }
        public string HeaderName { get; set; }
        public string FieldName { get; set; }
        public string Value { get; set; }
        public string ValueType { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public int? SubShift { get; set; }
        public long UserID { get; set; }
        public long? RowNumber { get; set; }
        public long? LocationID { get; set; }
    }
}