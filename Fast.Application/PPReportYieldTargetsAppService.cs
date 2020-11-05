using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldTargetsAppService : AppServiceBase<PPReportYieldTargets>, IPPReportYieldTargetsAppService
    {
        private readonly IPPReportYieldTargetsService _ppReportYieldTargetsService;
        public PPReportYieldTargetsAppService(IPPReportYieldTargetsService ppReportYieldTargetsService) : base(ppReportYieldTargetsService)
        {
            _ppReportYieldTargetsService = ppReportYieldTargetsService;
        }
    }
}