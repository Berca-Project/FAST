using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class LPHApprovalsService: ServiceBase<LPHApprovals>, ILPHApprovalsService
    {
        public LPHApprovalsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
