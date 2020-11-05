using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class LPHLocationsService: ServiceBase<LPHLocations>, ILPHLocationsService
    {
        public LPHLocationsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
