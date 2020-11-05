using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class ShuttleRequestService : ServiceBase<ShuttleRequest>, IShuttleRequestService
	{
		public ShuttleRequestService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
