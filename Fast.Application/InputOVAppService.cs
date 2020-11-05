using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class InputOVAppService: AppServiceBase<InputOV>, IInputOVAppService
    {
        public InputOVAppService(IInputOVService serviceBase) : base(serviceBase)
        {
        }
    }
}
