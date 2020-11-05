using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class AccessRightAppService : AppServiceBase<AccessRight>, IAccessRightAppService
	{
		public AccessRightAppService(IAccessRightService serviceBase) : base(serviceBase)
		{
		}
	}
}