using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class InputTargetAppService:AppServiceBase<InputTarget>, IInputTargetAppService
    {
        public InputTargetAppService(IInputTargetService serviceBase) : base(serviceBase)
        {
        }
    }
}
