using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class PPLPHSubmissionsService : ServiceBase<PPLPHSubmissions>, IPPLPHSubmissionsService
    {
        public PPLPHSubmissionsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
