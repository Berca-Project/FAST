using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class AccessRightService : ServiceBase<AccessRight>, IAccessRightService
	{
		public AccessRightService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
