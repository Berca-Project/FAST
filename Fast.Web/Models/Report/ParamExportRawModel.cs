using System;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class ParamExportRawModel : BaseModel
    {
        public string LPH { get; set; }
        public long ProdCenterID { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

    }
}