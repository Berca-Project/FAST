using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class UserMachineAppService : AppServiceBase<UserMachine>, IUserMachineAppService
	{
		public UserMachineAppService(IUserMachineService serviceBase) : base(serviceBase)
		{
		}
	}
}