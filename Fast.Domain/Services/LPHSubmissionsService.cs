using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class LPHSubmissionsService: ServiceBase<LPHSubmissions>, ILPHSubmissionsService
    {
        public LPHSubmissionsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
