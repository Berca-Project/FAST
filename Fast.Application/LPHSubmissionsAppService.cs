using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class LPHSubmissionsAppService: AppServiceBase<LPHSubmissions>, ILPHSubmissionsAppService
    {
        private readonly ILPHSubmissionsService _lphSubmissionsService;

        public LPHSubmissionsAppService(ILPHSubmissionsService lphSubmissionsService) : base(lphSubmissionsService)
        {
            _lphSubmissionsService = lphSubmissionsService;
        }
    }
}
