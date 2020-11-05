using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class UserMachineTypeAppService : AppServiceBase<UserMachineType>, IUserMachineTypeAppService
	{
		public UserMachineTypeAppService(IUserMachineTypeService serviceBase) : base(serviceBase)
		{
		}
	}
}