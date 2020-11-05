using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;

namespace Fast.Web.Models.Report
{
    public class ReportWasteModel : BaseModel
    {
        public int Week { get; set; }
        public int Year { get; set; }
        public string Area { get; set; }
        public string Waste { get; set; }
        public string WasteType { get; set; }
        public double Total { get; set; }
        public List<WasteDetailModel> WasteDetail { get; set; }
    }
    public class WasteDetailModel
    {
        public string Date { get; set; }
        public int Shift { get; set; }
        public double Value { get; set; }
    }

    public class WasteModel
    {
        public object this[string propertyName]
        {
            get => GetType().GetProperty(propertyName).GetValue(this, null);
            set => GetType().GetProperty(propertyName).SetValue(this, value, null);
        }
        public long LPHID { get; set; }
        public string Area { get; set; }
        public string Waste { get; set; }
        public string WasteType { get; set; }
        public string Frequency { get; set; }
        public double Weight { get; set; }
        public string Remarks { get; set; }
    }

    public class YieldWasteModel
    {
        public int Year { get; set; }
        public int Week { get; set; }
        public Dictionary<string, double> Wastes { get; set; }
    }

}