using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPReportYieldsAppService : AppServiceBase<PPReportYields>, IPPReportYieldsAppService
    {
        private readonly IPPReportYieldsService _PPReportYieldsService;
        public PPReportYieldsAppService(IPPReportYieldsService PPReportYieldsService) : base(PPReportYieldsService)
        {
            _PPReportYieldsService = PPReportYieldsService;
        }
    }
}