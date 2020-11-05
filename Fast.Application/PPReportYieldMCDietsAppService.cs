using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldMCDietsAppService : AppServiceBase<PPReportYieldMCDiets>, IPPReportYieldMCDietsAppService
    {
        private readonly IPPReportYieldMCDietsService _PPReportYieldMCDietsService;
        public PPReportYieldMCDietsAppService(IPPReportYieldMCDietsService PPReportYieldMCDietsService) : base(PPReportYieldMCDietsService)
        {
            _PPReportYieldMCDietsService = PPReportYieldMCDietsService;
        }
    }
}