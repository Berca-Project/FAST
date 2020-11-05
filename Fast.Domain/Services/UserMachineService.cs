using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class UserMachineService : ServiceBase<UserMachine>, IUserMachineService
	{
		public UserMachineService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
