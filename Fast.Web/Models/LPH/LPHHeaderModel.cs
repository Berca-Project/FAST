using System;
using System.Collections.Generic;

namespace Fast.Web.Models.LPH
{
	public class LPHHeaderModel
	{
		public LPHHeaderModel()
		{
			Machines = new List<string>();
			Brands = new List<string>();
		}

		public DateTime Date { get; set; }
		public int Week { get; set; }
		public List<string> Machines { get; set; }
		public List<string> Brands { get; set; }
		public string ProdTech { get; set; }
		public string Foreman { get; set; }
		public string Mechanic { get; set; }
		public string Electrician { get; set; }
		public string TeamLeader { get; set; }
		public string Relief { get; set; }
		public string Support { get; set; }
		public string GeneralWorker { get; set; }
		public string Other { get; set; }
		public int Shift { get; set; }
		public string Group { get; set; }
		public int Start { get; set; }
		public int Stop { get; set; }
	}
}