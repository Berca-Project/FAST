using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class PPLPHApprovalsService: ServiceBase<PPLPHApprovals>, IPPLPHApprovalsService
    {
        public PPLPHApprovalsService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
