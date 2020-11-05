using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class UserRoleService : ServiceBase<UserRole>, IUserRoleService
	{
		public UserRoleService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
