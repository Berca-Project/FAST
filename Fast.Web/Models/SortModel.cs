using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fast.Web.Models
{
    public class SortModel
    {
        public int id { get; set; }
        public List<SortModel> children { get; set; }
    }
}