using Fast.Application;
using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using Fast.Infra.CrossCutting.Common;
using Newtonsoft.Json;

namespace Fast.Application
{
	public class RoleAppService : AppServiceBase<Role>, IRoleAppService
	{
		public readonly IRoleService _service;

		public RoleAppService(IRoleService serviceBase) : base(serviceBase)
		{
			_service = serviceBase;
		}

		public string GetNameById(long id)
		{
			Role role = _service.GetById(id);
			return role == null ? string.Empty : role.Name;
		}

		public string GetByName(string name, bool asNoTracking = false)
		{
			Role role = _service.GetByName(name, asNoTracking);
			string result = role != null ? JsonHelper<Role>.Serialize(role) : string.Empty;

			return result;
		}

        public void RemoveEntity(string roleName)
        {
            Role role = _service.GetByName(roleName, false);                  
            _service.Remove(role);
            _service.Complete();
        }
    }
}