using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldIMLsAppService : AppServiceBase<PPReportYieldIMLs>, IPPReportYieldIMLsAppService
    {
        private readonly IPPReportYieldIMLsService _ppReportYieldIMLsService;
        public PPReportYieldIMLsAppService(IPPReportYieldIMLsService ppReportYieldIMLsService) : base(ppReportYieldIMLsService)
        {
            _ppReportYieldIMLsService = ppReportYieldIMLsService;
        }
    }
}