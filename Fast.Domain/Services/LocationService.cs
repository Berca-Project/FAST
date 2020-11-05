using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class LocationService : ServiceBase<Location>, ILocationService
    {
		public LocationService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
