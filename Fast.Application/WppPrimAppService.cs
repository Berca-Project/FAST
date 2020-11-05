using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class WppPrimAppService : AppServiceBase<WppPrim>, IWppPrimAppService
	{
		public WppPrimAppService(IWppPrimService serviceBase) : base(serviceBase)
		{
		}
	}
}