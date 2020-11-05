using System.Collections.Generic;

namespace Fast.Web.Models
{
    public class ReferenceModel : BaseModel
    {
		public long ReferenceID { get; set; }
		public string Name { get; set; }
        public string Purpose { get; set; }
		public List<ReferenceDetailModel> ReferenceDetails { get; set; }

        public ReferenceModel()
        {
            ReferenceDetails = new List<ReferenceDetailModel>();
            ReferenceDetails.Add(new ReferenceDetailModel { Code = "", Description = "" });
        }
    }

    public class ReferenceTreeModel : ReferenceModel
    {
		public string Code { get; set; }
		public string Description { get; set; }
		public List<ReferenceModel> Parents { get; set; }
        public ReferenceTreeModel()
        {
            Parents = new List<ReferenceModel>();
        }
    }
}