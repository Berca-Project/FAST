using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistAppService : AppServiceBase<Checklist>, IChecklistAppService
    {
        public ChecklistAppService(IChecklistService serviceBase) : base(serviceBase)
        {
        }
    }
}