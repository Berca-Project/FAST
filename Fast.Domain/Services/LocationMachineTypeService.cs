using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class LocationMachineTypeService : ServiceBase<LocationMachineType>, ILocationMachineTypeService
	{
		public LocationMachineTypeService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
