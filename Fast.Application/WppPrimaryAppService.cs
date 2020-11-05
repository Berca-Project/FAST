using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class WppPrimaryAppService : AppServiceBase<WppPrimary>, IWppPrimaryAppService
	{
		public WppPrimaryAppService(IWppPrimaryService serviceBase) : base(serviceBase)
		{
		}
	}
}