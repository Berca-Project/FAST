using Fast.Infra.CrossCutting.Common;
using System.Collections.Generic;

namespace Fast.Application.Interfaces
{
	public interface IReferenceAppService : IAppServiceBase
	{
		string GetDetailAll(ReferenceEnum referenceType, bool isExcludeDeleted = false);
		string GetDetailDescById(long id, bool asNoTracking = false);
		string GetDetailCodeById(long id, bool asNoTracking = false);
		void AddDetail(string data);
		string GetDetailById(long id, bool asNoTracking = false);
		string GetDetailBy(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string GetDetailBy(string propertyName, long propertyValue, bool isExcludeDeleted = false);
		string GetDetailAll(bool isExcludeDeleted = false);
		void UpdateDetail(string data);
		void RemoveDetail(long id);
		string GetDetail(ICollection<QueryFilter> filters);
		string GetDetail(ICollection<QueryFilter> filters, bool asNoTracking = false);
		string FindDetail(ICollection<QueryFilter> filters);
		string FindDetailBy(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string FindDetailBy(string propertyName, long propertyValue, bool isExcludeDeleted = false);
	}
}