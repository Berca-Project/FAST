using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class WppStpAppService : AppServiceBase<WppStp>, IWppStpAppService
	{
		public WppStpAppService(IWppStpService serviceBase) : base(serviceBase)
		{
		}
	}
}