using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class MppChangesService: ServiceBase<MppChanges>, IMppChangesService
    {
        public MppChangesService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
