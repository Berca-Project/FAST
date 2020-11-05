using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Fast.Domain.Entities
{
    public class LPHValues: BaseEntity
    {
        public long LPHComponentID { get; set; }
        public string Value { get; set; }
        public DateTime ValueDate { get; set; }
        public string ValueType { get; set; }
        public long? SubmissionID { get; set; }
        
        
    }
}
