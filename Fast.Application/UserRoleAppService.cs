using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class UserRoleAppService : AppServiceBase<UserRole>, IUserRoleAppService
	{
		public UserRoleAppService(IUserRoleService serviceBase) : base(serviceBase)
		{
		}
	}
}