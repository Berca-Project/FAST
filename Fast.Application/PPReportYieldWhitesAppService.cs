using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldWhitesAppService : AppServiceBase<PPReportYieldWhites>, IPPReportYieldWhitesAppService
    {
        private readonly IPPReportYieldWhitesService _PPReportYieldWhitesService;
        public PPReportYieldWhitesAppService(IPPReportYieldWhitesService PPReportYieldWhitesService) : base(PPReportYieldWhitesService)
        {
            _PPReportYieldWhitesService = PPReportYieldWhitesService;
        }
    }
}