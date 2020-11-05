using System.Collections.Generic;

namespace Fast.Web.Models.LPH.PP
{
    public class PPLPHModel : BaseModel
    {
        public string MenuTitle { get; set; }
        public string Header { get; set; }
        public string Type { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public string LPHType { get; set; }
		public PPLPHHeaderModel HeaderModel { get; set; }
	}

	public class PPLPHCompoValModel
    {
        public PPLPHComponentsModel Component { get; set; }
        public PPLPHValuesModel Value { get; set; }
    }

    public class PPLPHEditModel
    {
        public string ViewType { get; set; }
        public LPHHeaderModel Header { get; set; }
        public PPLPHModel LPH { get; set; }
        public List<PPLPHExtrasModel> Extras { get; set; }
        public List<PPLPHCompoValModel> CompoVal { get; set; }
    }
}