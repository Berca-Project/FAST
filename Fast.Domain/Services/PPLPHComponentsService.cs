using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class PPLPHComponentsService : ServiceBase<PPLPHComponents>, IPPLPHComponentsService
    {
        public PPLPHComponentsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
