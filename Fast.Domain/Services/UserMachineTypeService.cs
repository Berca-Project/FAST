using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class UserMachineTypeService : ServiceBase<UserMachineType>, IUserMachineTypeService
	{
		public UserMachineTypeService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
