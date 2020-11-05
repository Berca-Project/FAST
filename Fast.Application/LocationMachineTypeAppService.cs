using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class LocationMachineTypeAppService : AppServiceBase<LocationMachineType>, ILocationMachineTypeAppService
	{
		public LocationMachineTypeAppService(ILocationMachineTypeService serviceBase) : base(serviceBase)
		{
		}
	}
}