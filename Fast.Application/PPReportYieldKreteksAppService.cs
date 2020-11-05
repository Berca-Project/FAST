using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldKreteksAppService : AppServiceBase<PPReportYieldKreteks>, IPPReportYieldKreteksAppService
    {
        private readonly IPPReportYieldKreteksService _PPReportYieldKreteksService;
        public PPReportYieldKreteksAppService(IPPReportYieldKreteksService PPReportYieldKreteksService) : base(PPReportYieldKreteksService)
        {
            _PPReportYieldKreteksService = PPReportYieldKreteksService;
        }
    }
}