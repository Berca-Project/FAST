using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class UserLogAppService : AppServiceBase<UserLog>, IUserLogAppService
	{		
		public UserLogAppService(IUserLogService serviceBase) : base(serviceBase)
		{			
		}
	}
}