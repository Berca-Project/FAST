using System.Collections.Generic;
using Fast.Application;
using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using Fast.Infra.CrossCutting.Common;
using Newtonsoft.Json;

namespace Fast.Application
{
	public class ReferenceAppService : AppServiceBase<Reference>, IReferenceAppService
	{
		private readonly IReferenceDetailAppService _referenceDetailAppService;

		public ReferenceAppService(IReferenceService serviceBase, IReferenceDetailAppService referenceDetailAppService) : base(serviceBase)
		{
			_referenceDetailAppService = referenceDetailAppService;
		}
	
		public string GetDetailAll(ReferenceEnum referenceType, bool isExcludeDeleted = false)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("ReferenceID", (int)referenceType));
			if (isExcludeDeleted)
				filters.Add(new QueryFilter("IsDeleted", "0"));

			return _referenceDetailAppService.Find(filters);
		}

		public string GetDetailDescById(long id, bool asNoTracking = false)
		{
			string userType = _referenceDetailAppService.GetById(id, asNoTracking);
			ReferenceDetail userTypeObject = string.IsNullOrEmpty(userType) ? new ReferenceDetail() : JsonConvert.DeserializeObject<ReferenceDetail>(userType);

			return userTypeObject.Description;
		}

		public string GetDetailCodeById(long id, bool asNoTracking = false)
		{
			string userType = _referenceDetailAppService.GetById(id, asNoTracking);
			ReferenceDetail userTypeObject = string.IsNullOrEmpty(userType) ? new ReferenceDetail() : JsonConvert.DeserializeObject<ReferenceDetail>(userType);

			return userTypeObject.Code;
		}

		public void AddDetail(string data)
		{
			_referenceDetailAppService.Add(data);
		}

		public string FindDetail(ICollection<QueryFilter> filters)
		{
			return _referenceDetailAppService.Find(filters);
		}

		public string FindDetailBy(string propertyName, string propertyValue, bool isExcludeDeleted = false)
		{
			return _referenceDetailAppService.FindBy(propertyName, propertyValue, isExcludeDeleted);
		}

		public string FindDetailBy(string propertyName, long propertyValue, bool isExcludeDeleted = false)
		{
			return _referenceDetailAppService.FindBy(propertyName, propertyValue.ToString(), isExcludeDeleted);
		}

		public string GetDetail(ICollection<QueryFilter> filters)
		{
			return _referenceDetailAppService.Get(filters);
		}

		public string GetDetail(ICollection<QueryFilter> filters, bool asNoTracking = false)
		{
			return _referenceDetailAppService.Get(filters, asNoTracking);
		}

		public string GetDetailAll(bool isExcludeDeleted = false)
		{
			return _referenceDetailAppService.GetAll();
		}

		public string GetDetailBy(string propertyName, string propertyValue, bool isExcludeDeleted = false)
		{
			return _referenceDetailAppService.GetBy(propertyName, propertyValue, isExcludeDeleted);
		}

		public string GetDetailBy(string propertyName, long propertyValue, bool isExcludeDeleted = false)
		{
			return _referenceDetailAppService.GetBy(propertyName, propertyValue.ToString(), isExcludeDeleted);
		}

		public string GetDetailById(long id, bool asNoTracking = false)
		{
			return _referenceDetailAppService.GetById(id, asNoTracking);
		}

		public void RemoveDetail(long id)
		{
			_referenceDetailAppService.Remove(id);
		}

		public void UpdateDetail(string data)
		{
			_referenceDetailAppService.Update(data);
		}
	}
}