using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistLocationAppService : AppServiceBase<ChecklistLocation>, IChecklistLocationAppService
    {
        public ChecklistLocationAppService(IChecklistLocationService serviceBase) : base(serviceBase)
        {
        }
    }
}