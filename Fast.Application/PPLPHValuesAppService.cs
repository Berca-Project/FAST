using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class PPLPHValuesAppService : AppServiceBase<PPLPHValues>, IPPLPHValuesAppService
    {
        private readonly IPPLPHValuesService _PPLphValuesService;
        public PPLPHValuesAppService(IPPLPHValuesService PPLphValuesService) : base(PPLphValuesService)
        {
            _PPLphValuesService = PPLphValuesService;
        }
    }
}
