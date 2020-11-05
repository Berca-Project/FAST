using System;

namespace Fast.Web.Models
{
	public class MachineAllocationModel : BaseModel
	{
		public long MachineID { get; set; }
		public string MachineCode { get; set; }
		public string MachineCategory { get; set; }
        public Nullable<double> Value { get; set; }
    }
}