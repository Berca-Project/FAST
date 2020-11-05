using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class ManPowerAppService : AppServiceBase<ManPower>, IManPowerAppService
	{
		public ManPowerAppService(IManPowerService serviceBase) : base(serviceBase)
		{
		}
	}
}