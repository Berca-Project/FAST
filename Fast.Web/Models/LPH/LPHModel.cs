using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models.LPH
{
    public class LPHModel : BaseModel
    {
        public string MenuTitle { get; set; }
        public string Header { get; set; }
        public string Type { get; set; }
        public long LocationID { get; set; }
        public string Location { get; set; }
        public string LPHType { get; set; }
		public LPHHeaderModel HeaderModel { get; set; }
	}

	public class LPHCompoValModel
    {
        public LPHComponentsModel Component { get; set; }
        public LPHValuesModel Value { get; set; }
    }

    public class LPHEditModel
    {
        public LPHModel LPH { get; set; }
        public List<LPHExtrasModel> Extras { get; set; }
        public List<LPHCompoValModel> CompoVal { get; set; }
    }
}