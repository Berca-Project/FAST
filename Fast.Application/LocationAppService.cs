using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using Fast.Infra.CrossCutting.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace Fast.Application
{
	public class LocationAppService : AppServiceBase<Location>, ILocationAppService
	{
		public readonly ILocationService _service;

		public LocationAppService(ILocationService serviceBase) : base(serviceBase)
		{
			_service = serviceBase;
		}

		public string GetLocationFullCode(long locationID)
		{
			string location = string.Empty;

			List<Location> locationList = _service.GetAll().ToList();

			Location lvl1 = locationList.Where(x => x.ID == locationID).FirstOrDefault();
			if (lvl1 != null)
			{
				Location lvl2 = locationList.Where(x => x.ID == lvl1.ParentID).FirstOrDefault();
				if (lvl2 != null)
				{
					Location lvl3 = locationList.Where(x => x.ID == lvl2.ParentID).FirstOrDefault();
					if (lvl3 != null)
					{
						Location lvl4 = locationList.Where(x => x.ID == lvl3.ParentID).FirstOrDefault();
						if (lvl4 != null)
						{
							location = lvl4.Code + "-" + lvl3.Code + "-" + lvl2.Code + "-" + lvl1.Code;
						}
						else
						{
							location = lvl3.Code + "-" + lvl2.Code + "-" + lvl1.Code;
						}
					}
					else
					{
						location = lvl2.Code + "-" + lvl1.Code;
					}
				}
				else
				{
					location = lvl1.Code;
				}
			}

			return location;
		}

		public long GetLocationID(string locationCode)
		{
			if (string.IsNullOrEmpty(locationCode)) return 0;

			string[] locations = locationCode.Split('-');

			Location country = null;
			Location pc = null;
			Location dep = null;
			Location subDep = null;

			for (int i = 0; i < locations.Length; i++)
			{
				if (i == 0)
				{
					country = GetLocation(0, locations[i]);
				}

				if (i == 1)
				{
					pc = GetLocation(country.ID, locations[i]);
				}

				if (i == 2)
				{
					dep = GetLocation(pc.ID, locations[i]);
				}

				if (i == 3)
				{
					subDep = GetLocation(dep.ID, locations[i]);
				}
			}

			if (subDep != null) return subDep.ID;
			if (dep != null) return dep.ID;
			if (pc != null) return pc.ID;
			if (country != null) return country.ID;

			return 0;
		}

		private Location GetLocation(long parentID, string code)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("ParentID", parentID.ToString()));
			filters.Add(new QueryFilter("Code", code));
			filters.Add(new QueryFilter("IsDeleted", "0"));

			Expression<Func<Location, bool>> query = ExpressionBuilder.GetExpression<Location>(filters);
			IList<Location> entities = _service.Find(query).ToList();
			string locStr = entities.Any() ? JsonHelper<Location>.Serialize(entities[0]) : string.Empty;

			Location result = JsonConvert.DeserializeObject<Location>(locStr);

			return result;
		}
	}
}