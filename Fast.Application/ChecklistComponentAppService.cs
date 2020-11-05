using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistComponentAppService : AppServiceBase<ChecklistComponent>, IChecklistComponentAppService
    {
        public ChecklistComponentAppService(IChecklistComponentService serviceBase) : base(serviceBase)
        {
        }
    }
}