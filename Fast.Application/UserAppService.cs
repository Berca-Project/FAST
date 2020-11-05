using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class UserAppService : AppServiceBase<User>, IUserAppService
	{
        private readonly IUserService _userService;
        public UserAppService(IUserService serviceBase) : base(serviceBase)
		{
            _userService = serviceBase;

        }

        public string GetNameById(long id)
        {
            User user = _userService.GetById(id);
            return user == null ? string.Empty : user.UserName;
        }
    }
}