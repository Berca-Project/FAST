using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;
using System;

namespace Fast.Domain.Services
{
	public class UserLogService : ServiceBase<UserLog>, IUserLogService
	{
		public UserLogService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
