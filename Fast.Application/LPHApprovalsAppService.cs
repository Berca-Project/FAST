using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class LPHApprovalsAppService: AppServiceBase<LPHApprovals>, ILPHApprovalsAppService
    {
        private readonly ILPHApprovalsService _lphApprovalsService;

        public LPHApprovalsAppService(ILPHApprovalsService lphApprovalsService) : base(lphApprovalsService)
        {
            _lphApprovalsService = lphApprovalsService;
        }
      
    }
}
