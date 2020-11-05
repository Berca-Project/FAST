using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class ShuttleRequestAppService : AppServiceBase<ShuttleRequest>, IShuttleRequestAppService
	{
		public ShuttleRequestAppService(IShuttleRequestService serviceBase) : base(serviceBase)
		{
		}
	}
}