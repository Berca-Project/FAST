using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class InputDailyAppService: AppServiceBase<InputDaily>, IInputDailyAppService
    {
        public InputDailyAppService(IInputDailyService serviceBase) : base(serviceBase)
        {
        }
    }
}
