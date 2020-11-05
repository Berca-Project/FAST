using Fast.Infra.CrossCutting.Common;
using System.Collections.Generic;

namespace Fast.Application.Interfaces
{
	public interface IAppServiceBase
	{
		long Add(string data);
		void AddRange(string dataList);
		string GetById(long id, bool asNoTracking = false);
		string GetBy(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string GetBy(string propertyName, long propertyValue, bool isExcludeDeleted = false);
		string GetLastBy(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string GetLastBy(string propertyName, long propertyValue, bool isExcludeDeleted = false);
		string GetByNoTracking(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string GetByNoTracking(string propertyName, long propertyValue, bool isExcludeDeleted = false);
		string GetLastByNoTracking(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string GetLastByNoTracking(string propertyName, long propertyValue, bool isExcludeDeleted = false);
		string GetAll(bool isExcludeDeleted = false);
		string Update(string data);
		void Remove(long id);
		void RemoveRange(string dataList);
		string Get(ICollection<QueryFilter> filters);
		string Get(ICollection<QueryFilter> filters, bool asNoTracking = false);
		string GetLast(ICollection<QueryFilter> filters, bool asNoTracking = false);
		string Find(ICollection<QueryFilter> filters);
		string FindNoTracking(ICollection<QueryFilter> filters);
		string FindBy(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string FindByNoTracking(string propertyName, string propertyValue, bool isExcludeDeleted = false);
		string FindBy(string propertyName, long propertyValue, bool isExcludeDeleted = false);
	}
}
