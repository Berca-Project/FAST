using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class WppAppService : AppServiceBase<Wpp>, IWppAppService
	{
		public WppAppService(IWppService serviceBase) : base(serviceBase)
		{
		}
	}
}