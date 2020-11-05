using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPLPHSubmissionsAppService : AppServiceBase<PPLPHSubmissions>, IPPLPHSubmissionsAppService
    {
        private readonly IPPLPHSubmissionsService _PPLphSubmissionsService;

        public PPLPHSubmissionsAppService(IPPLPHSubmissionsService PPLphSubmissionsService) : base(PPLphSubmissionsService)
        {
            _PPLphSubmissionsService = PPLphSubmissionsService;
        }
    }
}
