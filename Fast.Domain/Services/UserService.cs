using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class UserService : ServiceBase<User>, IUserService
	{
		public UserService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
