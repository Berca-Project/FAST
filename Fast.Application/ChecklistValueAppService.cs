using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
    public class ChecklistValueAppService : AppServiceBase<ChecklistValue>, IChecklistValueAppService
    {
        public ChecklistValueAppService(IChecklistValueService serviceBase) : base(serviceBase)
        {
        }
    }
}