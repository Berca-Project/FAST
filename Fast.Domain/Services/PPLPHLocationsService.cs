using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class PPLPHLocationsService : ServiceBase<PPLPHLocations>, IPPLPHLocationsService
    {
        public PPLPHLocationsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
