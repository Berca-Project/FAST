namespace Fast.Web.Models
{
	public class TrainingTitleMachineTypeModel : BaseModel
	{
		public long TrainingTitleID { get; set; }
		public string TrainingTitle { get; set; }
		public long MachineTypeID { get; set; }
		public long[] MachineTypeIDs { get; set; }
		public string MachineType { get; set; }
		public string MachineTypeList { get; set; }
	}
}