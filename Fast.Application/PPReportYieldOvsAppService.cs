using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldOvsAppService : AppServiceBase<PPReportYieldOvs>, IPPReportYieldOvsAppService
    {
        private readonly IPPReportYieldOvsService _ppReportYieldOvsService;
        public PPReportYieldOvsAppService(IPPReportYieldOvsService ppReportYieldOvsService) : base(ppReportYieldOvsService)
        {
            _ppReportYieldOvsService = ppReportYieldOvsService;
        }
    }
}