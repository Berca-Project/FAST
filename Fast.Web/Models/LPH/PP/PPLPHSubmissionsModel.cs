using System;

namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHSubmissionsModel : BaseModel
    {
        public long LPHID { get; set; }
        public string LPHHeader { get; set; }
        public DateTime Date { get; set; }
        public string Shift { get; set; }
        public int? SubShift { get; set; }
        public long UserID { get; set; }
        public long LocationID { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string Location { get; set; }
        public string LPHType { get; set; }
		public bool IsCompleted { get; set; }

        public DateTime CreatedAt { get; set; }
        public DateTime? StatusChangedAt { get; set; }
        public string Status { get; set; }

        public string UserFullName { get; set; }
        public Boolean IsComplete { get; set; }

        public int? Flag { get; set; }

    }
}