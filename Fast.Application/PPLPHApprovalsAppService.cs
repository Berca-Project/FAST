using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPLPHApprovalsAppService: AppServiceBase<PPLPHApprovals>, IPPLPHApprovalsAppService
    {
        private readonly IPPLPHApprovalsService _PPLphApprovalsService;

        public PPLPHApprovalsAppService(IPPLPHApprovalsService PPLphApprovalsService) : base(PPLphApprovalsService)
        {
            _PPLphApprovalsService = PPLphApprovalsService;
        }
      
    }
}
